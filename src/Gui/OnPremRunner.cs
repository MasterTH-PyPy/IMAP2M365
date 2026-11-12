using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace IMAP2M365.Gui;

/// <summary>Ein Postfach auf dem Exchange on-prem (für den Abgleich mit den Zielpostfächern).</summary>
public sealed record OnPremMailbox(string PrimarySmtp, string Alias, string DisplayName, string Type);

/// <summary>Stand der Migrationsrechte des Admin-Kontos auf dem Exchange on-prem.</summary>
public sealed record OnPremStatus(
    bool HasImpersonation, bool OwnRoleAssignment, bool OwnPolicyExists, string AssignedPolicy);

/// <summary>
/// Verwaltet Exchange on-prem über die Exchange-Remote-PowerShell (WinRM-Endpunkt
/// http(s)://server/PowerShell/): Postfächer auslesen, dem Migrationskonto die Rolle
/// ApplicationImpersonation geben, eine EWS-Throttling-Policy zuweisen – und alles wieder entfernen.
///
/// Der Exchange-Endpunkt läuft im eingeschränkten Sprachmodus. Deshalb wird jeder Aufruf als
/// einzelnes Cmdlet mit festen Werten geschickt (keine Variablen, keine Skriptlogik auf dem Server);
/// gefiltert und ausgewertet wird lokal.
/// </summary>
public sealed class OnPremRunner : IDisposable
{
    public const string RoleAssignmentName = "IMAP2M365-Migration";
    public const string ThrottlingPolicyName = "IMAP2M365-Migration";

    private readonly Runspace _runspace;
    private bool _connected;

    public Action<string>? Log { get; set; }
    public bool IsConnected => _connected;

    public OnPremRunner()
    {
        var iss = InitialSessionState.CreateDefault2();
        iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.RemoteSigned;
        _runspace = RunspaceFactory.CreateRunspace(iss);
        _runspace.Open();
    }

    private PowerShell NewPs()
    {
        var ps = PowerShell.Create();
        ps.Runspace = _runspace;
        return ps;
    }

    private static string Errors(PowerShell ps) =>
        string.Join("; ", ps.Streams.Error.Select(e => e.ToString()));

    /// <summary>Wert als PowerShell-Literal in einfachen Anführungszeichen.</summary>
    private static string Q(string value) => "'" + value.Replace("'", "''") + "'";

    /// <summary>
    /// Baut die Remote-Sitzung zum Exchange auf.
    /// </summary>
    /// <param name="psUri">z. B. http://exchange01.firma.local/PowerShell/</param>
    /// <param name="auth">Kerberos (Standard in der Domäne), Negotiate oder Basic.</param>
    public void Connect(string psUri, string user, string password, string auth, bool skipCertificateChecks)
    {
        using var ps = NewPs();
        ps.AddScript(@"
            param($uri, $user, $pw, $auth, $skip)
            $sec  = ConvertTo-SecureString $pw -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($user, $sec)
            if ($global:OnPremSession) { Remove-PSSession $global:OnPremSession -ErrorAction SilentlyContinue }
            $opt = if ($skip) { New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck } else { New-PSSessionOption }
            $global:OnPremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri `
                -Authentication $auth -Credential $cred -SessionOption $opt -ErrorAction Stop")
          .AddParameter("uri", psUri)
          .AddParameter("user", user)
          .AddParameter("pw", password)
          .AddParameter("auth", auth)
          .AddParameter("skip", skipCertificateChecks);
        ps.Invoke();
        if (ps.HadErrors)
            throw new InvalidOperationException("Remote-PowerShell zum Exchange fehlgeschlagen: " + Errors(ps));

        _connected = true;
        Log?.Invoke($"Mit Exchange on-prem verbunden ({psUri}, {auth}).");
    }

    /// <summary>Führt genau ein Cmdlet (mit festen Werten) auf dem Exchange aus.</summary>
    private Collection<PSObject> Remote(string command, string what)
    {
        if (!_connected)
            throw new InvalidOperationException("Nicht mit dem Exchange on-prem verbunden.");

        using var ps = NewPs();
        ps.AddScript(@"
            param($cmd)
            Invoke-Command -Session $global:OnPremSession -ScriptBlock ([scriptblock]::Create($cmd)) -ErrorAction Stop")
          .AddParameter("cmd", command);
        var result = ps.Invoke();
        if (ps.HadErrors)
            throw new InvalidOperationException($"{what} fehlgeschlagen: {Errors(ps)}");
        return result;
    }

    private static string Prop(PSObject o, string name) => o.Properties[name]?.Value?.ToString() ?? "";

    public List<OnPremMailbox> LoadMailboxes(bool includeShared)
    {
        var types = includeShared ? "UserMailbox,SharedMailbox" : "UserMailbox";
        var rows = Remote($"Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails {types}", "Get-Mailbox");

        var list = rows
            .Select(o => new OnPremMailbox(
                Prop(o, "PrimarySmtpAddress"), Prop(o, "Alias"), Prop(o, "DisplayName"), Prop(o, "RecipientTypeDetails")))
            .Where(m => m.PrimarySmtp.Contains('@'))
            .ToList();

        Log?.Invoke($"{list.Count} Postfächer vom Exchange on-prem gelesen.");
        return list;
    }

    /// <summary>Liest, welche Migrationsrechte das Konto aktuell hat (ändert nichts).</summary>
    public OnPremStatus GetStatus(string adminUser)
    {
        // Rolle: egal über welche Zuweisung (direkt oder über eine Gruppe) – Hauptsache wirksam.
        var anyRole = Remote(
            $"Get-ManagementRoleAssignment -Role ApplicationImpersonation -RoleAssignee {Q(adminUser)} -ErrorAction SilentlyContinue",
            "Get-ManagementRoleAssignment");
        var ownRole = anyRole.Any(o => string.Equals(Prop(o, "Name"), RoleAssignmentName, StringComparison.OrdinalIgnoreCase));

        var policies = Remote("Get-ThrottlingPolicy", "Get-ThrottlingPolicy");
        var ownPolicy = policies.Any(o => string.Equals(Prop(o, "Name"), ThrottlingPolicyName, StringComparison.OrdinalIgnoreCase));

        var assoc = Remote($"Get-ThrottlingPolicyAssociation -Identity {Q(adminUser)}", "Get-ThrottlingPolicyAssociation");
        var assigned = assoc.Count > 0 ? Prop(assoc[0], "ThrottlingPolicyId") : "";

        return new OnPremStatus(anyRole.Count > 0, ownRole, ownPolicy, assigned);
    }

    /// <summary>
    /// Gibt dem Konto ApplicationImpersonation (eine Zuweisung für alle Postfächer), sofern es die
    /// Rolle nicht schon hat. Gibt true zurück, wenn eine neue Zuweisung angelegt wurde.
    /// </summary>
    public bool GrantImpersonation(string adminUser)
    {
        var status = GetStatus(adminUser);
        if (status.HasImpersonation)
        {
            Log?.Invoke(status.OwnRoleAssignment
                ? "ApplicationImpersonation ist bereits zugewiesen (eigene Zuweisung)."
                : "ApplicationImpersonation ist bereits über eine andere Zuweisung vorhanden – nichts angelegt.");
            return false;
        }

        Remote(
            $"New-ManagementRoleAssignment -Name {Q(RoleAssignmentName)} -Role ApplicationImpersonation -User {Q(adminUser)}",
            "New-ManagementRoleAssignment");
        Log?.Invoke($"Rolle ApplicationImpersonation zugewiesen (Zuweisung '{RoleAssignmentName}').");
        return true;
    }

    /// <summary>
    /// Legt eine EWS-Throttling-Policy ohne Limits an und weist sie dem Migrationskonto zu.
    /// Hat das Konto bereits eine fremde Policy, bleibt diese unangetastet (false).
    /// Parameter gelten für Exchange 2013 und neuer.
    /// </summary>
    public bool ApplyThrottling(string adminUser)
    {
        var status = GetStatus(adminUser);
        var foreign = status.AssignedPolicy.Length > 0
                      && !status.AssignedPolicy.Contains(ThrottlingPolicyName, StringComparison.OrdinalIgnoreCase);
        if (foreign)
        {
            Log?.Invoke($"Throttling: Konto hat bereits die Policy '{status.AssignedPolicy}' – wird nicht überschrieben.");
            return false;
        }

        if (!status.OwnPolicyExists)
        {
            Remote(
                $"New-ThrottlingPolicy -Name {Q(ThrottlingPolicyName)} -ThrottlingPolicyScope Regular " +
                "-EwsMaxConcurrency Unlimited -EwsMaxBurst Unlimited -EwsRechargeRate Unlimited " +
                "-EwsCutoffBalance Unlimited -EwsMaxSubscriptions Unlimited",
                "New-ThrottlingPolicy");
            Log?.Invoke($"Throttling-Policy '{ThrottlingPolicyName}' angelegt (EWS ohne Limits).");
        }

        if (status.AssignedPolicy.Length == 0)
        {
            Remote(
                $"Set-ThrottlingPolicyAssociation -Identity {Q(adminUser)} -ThrottlingPolicy {Q(ThrottlingPolicyName)}",
                "Set-ThrottlingPolicyAssociation");
            Log?.Invoke("Throttling-Policy dem Migrationskonto zugewiesen.");
        }
        else
        {
            Log?.Invoke("Throttling-Policy ist dem Migrationskonto bereits zugewiesen.");
        }
        return true;
    }

    /// <summary>
    /// Baut zurück, was dieses Tool angelegt hat: eigene Rollenzuweisung, eigene Policy-Zuordnung
    /// und die eigene Policy. Fremde Zuweisungen/Policies bleiben bestehen.
    /// </summary>
    public void RevokeAll(string adminUser)
    {
        var status = GetStatus(adminUser);

        if (status.OwnRoleAssignment)
        {
            Remote($"Remove-ManagementRoleAssignment -Identity {Q(RoleAssignmentName)} -Confirm:$false",
                "Remove-ManagementRoleAssignment");
            Log?.Invoke($"Rollenzuweisung '{RoleAssignmentName}' entfernt.");
        }
        else if (status.HasImpersonation)
        {
            Log?.Invoke("ApplicationImpersonation stammt aus einer fremden Zuweisung – bleibt bestehen (nicht von diesem Tool angelegt).");
        }
        else
        {
            Log?.Invoke("Keine Rollenzuweisung vorhanden – nichts zu entfernen.");
        }

        var ownAssigned = status.AssignedPolicy.Contains(ThrottlingPolicyName, StringComparison.OrdinalIgnoreCase);
        if (ownAssigned)
        {
            Remote($"Set-ThrottlingPolicyAssociation -Identity {Q(adminUser)} -ThrottlingPolicy $null",
                "Set-ThrottlingPolicyAssociation");
            Log?.Invoke("Throttling-Zuordnung des Migrationskontos zurückgesetzt.");
        }

        if (status.OwnPolicyExists)
        {
            Remote($"Remove-ThrottlingPolicy -Identity {Q(ThrottlingPolicyName)} -Confirm:$false", "Remove-ThrottlingPolicy");
            Log?.Invoke($"Throttling-Policy '{ThrottlingPolicyName}' entfernt.");
        }
    }

    public void Dispose()
    {
        try
        {
            using var ps = NewPs();
            ps.AddScript("if ($global:OnPremSession) { Remove-PSSession $global:OnPremSession -ErrorAction SilentlyContinue }");
            ps.Invoke();
        }
        catch { /* best effort */ }
        _runspace.Dispose();
    }
}
