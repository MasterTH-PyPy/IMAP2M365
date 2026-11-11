using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace IMAP2M365.Gui;

/// <summary>
/// Führt Exchange-Online-Cmdlets in-process über das PowerShell SDK aus.
/// Erwartet das Modul ExchangeOnlineManagement (wird bei Bedarf installiert, falls
/// AutoInstallModule = true). Rückgaben kommen als PSObject zurück und werden typisiert
/// ausgewertet – so landet die Postfachliste direkt im Grid.
/// </summary>
public sealed class ExoRunner : IDisposable
{
    private readonly Runspace _runspace;
    private bool _connected;

    public Action<string>? Log { get; set; }
    public bool AutoInstallModule { get; set; } = true;

    public ExoRunner()
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

    private void CheckErrors(PowerShell ps, string op)
    {
        if (ps.HadErrors)
        {
            var msg = string.Join("; ", ps.Streams.Error.Select(e => e.ToString()));
            throw new InvalidOperationException($"{op} fehlgeschlagen: {msg}");
        }
    }

    public void EnsureModule()
    {
        using var ps = NewPs();
        ps.AddScript("Get-Module -ListAvailable -Name ExchangeOnlineManagement");
        var found = ps.Invoke();
        if (found.Count == 0)
        {
            if (!AutoInstallModule)
                throw new InvalidOperationException(
                    "Modul ExchangeOnlineManagement fehlt. Einmalig installieren: " +
                    "Install-Module ExchangeOnlineManagement -Scope CurrentUser");
            Log?.Invoke("Installiere Modul ExchangeOnlineManagement (CurrentUser) ...");

            using var install = NewPs();
            // Nur der moderne PSResourceGet-Pfad (PS7): nicht-interaktiv, kein NuGet-Provider nötig.
            // Der alte PowerShellGet/Install-Module-Pfad wird bewusst NICHT genutzt – im gehosteten
            // Runspace greift er auf PowerShellGet 1.0.0.1 (WinPS 5.1) zu und erzeugt Clobber-Fehler.
            install.AddScript(@"
                $ErrorActionPreference = 'Stop'
                $ConfirmPreference = 'None'
                try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}
                Import-Module Microsoft.PowerShell.PSResourceGet -ErrorAction SilentlyContinue
                if (-not (Get-Command Install-PSResource -ErrorAction SilentlyContinue)) {
                    throw 'PSResourceGet im Host nicht verfügbar'
                }
                Install-PSResource -Name ExchangeOnlineManagement -Scope CurrentUser -TrustRepository -Quiet");
            install.Invoke();
            if (install.HadErrors)
            {
                var msg = string.Join("; ", install.Streams.Error.Select(e => e.ToString()));
                throw new InvalidOperationException(
                    "Automatische Installation nicht möglich (" + msg + "). " +
                    "Bitte das Modul einmalig in einer normalen PowerShell installieren und das Tool danach neu verbinden:\r\n" +
                    "    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber");
            }
        }
        using var import = NewPs();
        import.AddScript("Import-Module ExchangeOnlineManagement -ErrorAction Stop");
        import.Invoke();
        CheckErrors(import, "Import-Module");
    }

    /// <summary>Interaktiver Admin-Login (Browser/Device-Code via Connect-ExchangeOnline).</summary>
    public void Connect(string organization, string? userPrincipalName = null)
    {
        using var ps = NewPs();
        ps.AddCommand("Connect-ExchangeOnline")
          .AddParameter("Organization", organization)
          .AddParameter("ShowBanner", false);
        if (!string.IsNullOrWhiteSpace(userPrincipalName))
            ps.AddParameter("UserPrincipalName", userPrincipalName);
        ps.Invoke();
        CheckErrors(ps, "Connect-ExchangeOnline");
        _connected = true;
        Log?.Invoke($"Mit Exchange Online verbunden ({organization}).");
    }

    /// <summary>
    /// Liest Postfächer aus – User + Shared – inkl. IMAP-Status (Get-CASMailbox).
    /// </summary>
    public List<MailboxRow> LoadMailboxes(bool includeShared = true)
    {
        EnsureConnected();
        var types = includeShared ? "'UserMailbox','SharedMailbox'" : "'UserMailbox'";

        using var ps = NewPs();
        ps.AddScript($@"
            $boxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails {types} |
                     Select-Object PrimarySmtpAddress, RecipientTypeDetails
            foreach ($b in $boxes) {{
                $cas = Get-EXOCASMailbox -Identity $b.PrimarySmtpAddress | Select-Object ImapEnabled
                [pscustomobject]@{{
                    Smtp = [string]$b.PrimarySmtpAddress
                    Type = [string]$b.RecipientTypeDetails
                    Imap = [bool]$cas.ImapEnabled
                }}
            }}");
        var results = ps.Invoke();
        CheckErrors(ps, "Get-EXOMailbox");

        var rows = new List<MailboxRow>();
        foreach (var o in results)
        {
            rows.Add(new MailboxRow
            {
                TargetSmtp = o.Properties["Smtp"]?.Value?.ToString() ?? "",
                MailboxType = o.Properties["Type"]?.Value?.ToString() ?? "UserMailbox",
                ImapEnabled = o.Properties["Imap"]?.Value is bool b && b,
                SourceUser = o.Properties["Smtp"]?.Value?.ToString() ?? "" // Vorbelegung, editierbar
            });
        }
        Log?.Invoke($"{rows.Count} Postfächer geladen ({(includeShared ? "User + Shared" : "nur User")}).");
        return rows;
    }

    /// <summary>IMAP für ein Postfach einschalten.</summary>
    public void EnableImap(string smtp)
    {
        EnsureConnected();
        using var ps = NewPs();
        ps.AddCommand("Set-CASMailbox").AddParameter("Identity", smtp).AddParameter("ImapEnabled", true);
        ps.Invoke();
        CheckErrors(ps, $"Set-CASMailbox ({smtp})");
        Log?.Invoke($"IMAP aktiviert: {smtp}");
    }

    /// <summary>Service-Principal der App in Exchange registrieren (idempotent).</summary>
    public void RegisterServicePrincipal(string appId, string objectId, string displayName = "IMAP2M365 Migration")
    {
        EnsureConnected();
        using var ps = NewPs();
        ps.AddScript($@"
            $sp = Get-ServicePrincipal -ErrorAction SilentlyContinue | Where-Object {{ $_.AppId -eq '{appId}' }}
            if (-not $sp) {{
                New-ServicePrincipal -AppId '{appId}' -ServiceId '{objectId}' -DisplayName '{displayName}'
            }}");
        ps.Invoke();
        CheckErrors(ps, "New-ServicePrincipal");
        Log?.Invoke($"ServicePrincipal registriert/vorhanden: {appId}");
    }

    /// <summary>FullAccess des Service-Principals auf ein Zielpostfach (auch Shared).</summary>
    public void GrantMailboxPermission(string smtp, string objectId)
    {
        EnsureConnected();
        using var ps = NewPs();
        ps.AddCommand("Add-MailboxPermission")
          .AddParameter("Identity", smtp)
          .AddParameter("User", objectId)
          .AddParameter("AccessRights", "FullAccess")
          .AddParameter("AutoMapping", false);
        ps.Invoke();
        // Add-MailboxPermission meldet "already exists" als non-terminating; tolerieren.
        if (ps.HadErrors && !ps.Streams.Error.Any(e => e.ToString().Contains("already", StringComparison.OrdinalIgnoreCase)))
            CheckErrors(ps, $"Add-MailboxPermission ({smtp})");
        Log?.Invoke($"Postfachrecht erteilt: {smtp}");
    }

    /// <summary>Optionale ApplicationAccessPolicy – begrenzt App-Zugriff auf eine Gruppe.</summary>
    public void SetApplicationAccessPolicy(string appId, string scopeGroupId)
    {
        EnsureConnected();
        using var ps = NewPs();
        ps.AddCommand("New-ApplicationAccessPolicy")
          .AddParameter("AppId", appId)
          .AddParameter("PolicyScopeGroupId", scopeGroupId)
          .AddParameter("AccessRight", "RestrictAccess")
          .AddParameter("Description", "IMAP2M365 nur auf Migrations-Postfächer");
        ps.Invoke();
        CheckErrors(ps, "New-ApplicationAccessPolicy");
        Log?.Invoke($"ApplicationAccessPolicy gesetzt (Gruppe {scopeGroupId}).");
    }

    /// <summary>Prüft, ob die App auf ein Postfach zugreifen darf.</summary>
    public bool TestAccess(string appId, string smtp)
    {
        EnsureConnected();
        using var ps = NewPs();
        ps.AddCommand("Test-ApplicationAccessPolicy").AddParameter("AppId", appId).AddParameter("Identity", smtp);
        var r = ps.Invoke();
        CheckErrors(ps, $"Test-ApplicationAccessPolicy ({smtp})");
        var access = r.FirstOrDefault()?.Properties["AccessCheckResult"]?.Value?.ToString();
        return string.Equals(access, "Granted", StringComparison.OrdinalIgnoreCase);
    }

    private void EnsureConnected()
    {
        if (!_connected)
            throw new InvalidOperationException("Nicht mit Exchange Online verbunden. Erst Connect() aufrufen.");
    }

    public void Dispose()
    {
        try
        {
            if (_connected)
            {
                using var ps = NewPs();
                ps.AddCommand("Disconnect-ExchangeOnline").AddParameter("Confirm", false);
                ps.Invoke();
            }
        }
        catch { /* best effort */ }
        _runspace.Dispose();
    }
}
