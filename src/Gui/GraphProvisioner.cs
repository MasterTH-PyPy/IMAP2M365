using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;

namespace IMAP2M365.Gui;

/// <summary>
/// Legt die Entra-App-Registrierung für app-only XOAUTH2 an und erteilt Admin-Consent –
/// über reines Graph-REST (HttpClient), ohne Microsoft.Graph/Kiota.
/// Admin meldet sich einmal interaktiv im Browser an (Konto-Auswahl). Kein Device-Code,
/// weil der in vielen Tenants per Conditional Access gesperrt ist (AADSTS530035).
///
/// Benötigte delegierte Rechte des anmeldenden Admins:
///   Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All.
/// </summary>
public sealed class GraphProvisioner
{
    private const string Base = "https://graph.microsoft.com/v1.0";

    // Ressourcen-AppIds (konstant in jedem Tenant)
    private const string ExchangeOnlineResourceAppId = "00000002-0000-0ff1-ce00-000000000000"; // Office 365 Exchange Online
    private const string GraphResourceAppId          = "00000003-0000-0000-c000-000000000000"; // Microsoft Graph

    // AppRole-Namen (value) – die IDs werden dynamisch aus dem Ressourcen-ServicePrincipal
    // nachgeschlagen, weil hart kodierte GUIDs je nach Tenant/Zeitpunkt nicht passen
    // ("Permission being assigned was not found on application").
    private const string ImapAccessAsAppRole = "IMAP.AccessAsApp"; // Office 365 Exchange Online
    private const string MailReadWriteRole   = "Mail.ReadWrite";   // Microsoft Graph

    // Delegated-Scopes, die der anmeldende Admin für die Provisionierung braucht.
    // Explizit anfordern (nicht ".default") -> löst beim ersten Login den Consent-Dialog aus.
    private static readonly string[] GraphScopes =
    {
        "https://graph.microsoft.com/Application.ReadWrite.All",
        "https://graph.microsoft.com/AppRoleAssignment.ReadWrite.All",
        "https://graph.microsoft.com/Directory.ReadWrite.All"
    };

    private readonly HttpClient _http = new();

    public Action<string>? Log { get; set; }

    private InteractiveBrowserCredential? _cred;
    private string? _credTenant;

    private TokenCredential Credential(string tenantId)
    {
        // Pro Instanz nur EIN Browser-Login; Folgeaufrufe nutzen den Token-Cache des Credentials.
        if (_cred is not null && _credTenant == tenantId) return _cred;

        Log?.Invoke("Browser zur Anmeldung wird geöffnet – bitte Admin-Konto auswählen.");
        _cred = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
        {
            TenantId = tenantId,
            ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e", // Microsoft Graph Command Line Tools (öffentlicher Client)
            RedirectUri = new Uri("http://localhost")          // Auth-Code-Flow mit lokalem Redirect
        });
        _credTenant = tenantId;
        return _cred;
    }

    public async Task<(string appId, string spObjectId, string secret)> ProvisionAsync(
        string tenantId, string displayName, bool includeGraphTarget, CancellationToken ct)
    {
        var cred = Credential(tenantId);

        // Rollen-IDs dynamisch aus den Ressourcen-ServicePrincipals auflösen (keine GUID-Rate-rei).
        var exo = await ResolveRoleAsync(cred, ExchangeOnlineResourceAppId, ImapAccessAsAppRole, ct).ConfigureAwait(false);
        (string SpId, string RoleId)? graph = includeGraphTarget
            ? await ResolveRoleAsync(cred, GraphResourceAppId, MailReadWriteRole, ct).ConfigureAwait(false)
            : null;

        // 1) App-Registrierung mit requiredResourceAccess (aufgelöste Rollen-IDs)
        var resourceAccess = new List<object>
        {
            new
            {
                resourceAppId = ExchangeOnlineResourceAppId,
                resourceAccess = new[] { new { id = exo.RoleId, type = "Role" } }
            }
        };
        if (graph is not null)
            resourceAccess.Add(new
            {
                resourceAppId = GraphResourceAppId,
                resourceAccess = new[] { new { id = graph.Value.RoleId, type = "Role" } }
            });

        var appJson = await PostAsync(cred, "/applications", new
        {
            displayName,
            signInAudience = "AzureADMyOrg",
            requiredResourceAccess = resourceAccess
        }, ct).ConfigureAwait(false);

        var appObjectId = appJson.GetProperty("id").GetString()!;
        var appId = appJson.GetProperty("appId").GetString()!;
        Log?.Invoke($"App angelegt: {appId}");

        // 2) Client-Secret
        var secretJson = await PostAsync(cred, $"/applications/{appObjectId}/addPassword", new
        {
            passwordCredential = new { displayName = "IMAP2M365", endDateTime = DateTimeOffset.UtcNow.AddYears(1) }
        }, ct).ConfigureAwait(false);
        var secret = secretJson.GetProperty("secretText").GetString() ?? "";
        Log?.Invoke("Client-Secret erzeugt (1 Jahr gültig).");

        // 3) ServicePrincipal zur App (kurz retrien, App-Propagation abwarten)
        string spObjectId = await CreateServicePrincipalAsync(cred, appId, ct).ConfigureAwait(false);

        // 4) Admin-Consent = appRoleAssignment je Ressource (aufgelöste SP-/Rollen-IDs)
        await AssignAppRoleAsync(cred, spObjectId, exo.SpId, exo.RoleId, ct).ConfigureAwait(false);
        if (graph is not null)
            await AssignAppRoleAsync(cred, spObjectId, graph.Value.SpId, graph.Value.RoleId, ct).ConfigureAwait(false);

        Log?.Invoke("Admin-Consent erteilt.");
        return (appId, spObjectId, secret);
    }

    /// <summary>Listet vorhandene App-Registrierungen des Tenants (zur Auswahl einer bestehenden App).</summary>
    public async Task<List<(string DisplayName, string AppId, string AppObjectId)>> ListApplicationsAsync(
        string tenantId, CancellationToken ct)
    {
        var cred = Credential(tenantId);
        var res = await GetAsync(cred, "/applications?$select=id,appId,displayName&$top=200", ct).ConfigureAwait(false);

        var list = new List<(string, string, string)>();
        foreach (var a in res.GetProperty("value").EnumerateArray())
        {
            list.Add((
                a.TryGetProperty("displayName", out var dn) ? dn.GetString() ?? "(ohne Namen)" : "(ohne Namen)",
                a.GetProperty("appId").GetString() ?? "",
                a.GetProperty("id").GetString() ?? ""));
        }
        return list;
    }

    /// <summary>
    /// Bereitet eine bestehende App für app-only XOAUTH2 vor: frisches Secret erzeugen
    /// (bestehende sind nicht auslesbar), ServicePrincipal holen/anlegen, Consent best-effort.
    /// </summary>
    public async Task<(string spObjectId, string secret)> PrepareExistingAppAsync(
        string tenantId, string appObjectId, string appId, bool includeGraphTarget, CancellationToken ct)
    {
        var cred = Credential(tenantId);

        var secretJson = await PostAsync(cred, $"/applications/{appObjectId}/addPassword", new
        {
            passwordCredential = new { displayName = "IMAP2M365", endDateTime = DateTimeOffset.UtcNow.AddYears(1) }
        }, ct).ConfigureAwait(false);
        var secret = secretJson.GetProperty("secretText").GetString() ?? "";

        var spRes = await GetAsync(cred, $"/servicePrincipals?$filter=appId eq '{appId}'&$select=id&$top=1", ct)
            .ConfigureAwait(false);
        var spArr = spRes.GetProperty("value");
        var spObjectId = spArr.GetArrayLength() > 0
            ? spArr[0].GetProperty("id").GetString()!
            : await CreateServicePrincipalAsync(cred, appId, ct).ConfigureAwait(false);

        await TryAssignRoleAsync(cred, spObjectId, ExchangeOnlineResourceAppId, ImapAccessAsAppRole, ct).ConfigureAwait(false);
        if (includeGraphTarget)
            await TryAssignRoleAsync(cred, spObjectId, GraphResourceAppId, MailReadWriteRole, ct).ConfigureAwait(false);

        return (spObjectId, secret);
    }

    private async Task TryAssignRoleAsync(
        TokenCredential cred, string clientSpId, string resourceAppId, string roleValue, CancellationToken ct)
    {
        try
        {
            var (spId, roleId) = await ResolveRoleAsync(cred, resourceAppId, roleValue, ct).ConfigureAwait(false);
            await AssignAppRoleAsync(cred, clientSpId, spId, roleId, ct).ConfigureAwait(false);
        }
        catch (HttpRequestException ex) when (ex.Message.Contains("already", StringComparison.OrdinalIgnoreCase)
                                              || ex.Message.Contains("Permission being assigned already exists", StringComparison.OrdinalIgnoreCase))
        {
            // appRoleAssignment existiert bereits -> ok
        }
    }

    /// <summary>Findet (ServicePrincipal-Id, AppRole-Id) einer Ressource anhand des Rollennamens (value).</summary>
    private async Task<(string SpId, string RoleId)> ResolveRoleAsync(
        TokenCredential cred, string resourceAppId, string roleValue, CancellationToken ct)
    {
        var res = await GetAsync(cred, $"/servicePrincipals?$filter=appId eq '{resourceAppId}'&$select=id,appRoles&$top=1", ct)
            .ConfigureAwait(false);
        var arr = res.GetProperty("value");
        if (arr.GetArrayLength() == 0)
            throw new InvalidOperationException($"Ressourcen-ServicePrincipal {resourceAppId} nicht gefunden (Ressource im Tenant nicht vorhanden?).");

        var sp = arr[0];
        var spId = sp.GetProperty("id").GetString()!;
        foreach (var role in sp.GetProperty("appRoles").EnumerateArray())
        {
            if (string.Equals(role.GetProperty("value").GetString(), roleValue, StringComparison.OrdinalIgnoreCase))
            {
                var allowsApp = role.GetProperty("allowedMemberTypes").EnumerateArray()
                    .Any(t => string.Equals(t.GetString(), "Application", StringComparison.OrdinalIgnoreCase));
                if (allowsApp)
                    return (spId, role.GetProperty("id").GetString()!);
            }
        }
        throw new InvalidOperationException($"App-Rolle '{roleValue}' an Ressource {resourceAppId} nicht gefunden.");
    }

    private async Task AssignAppRoleAsync(
        TokenCredential cred, string clientSpId, string resourceSpId, string appRoleId, CancellationToken ct)
    {
        await PostAsync(cred, $"/servicePrincipals/{clientSpId}/appRoleAssignments", new
        {
            principalId = clientSpId,
            resourceId = resourceSpId,
            appRoleId
        }, ct).ConfigureAwait(false);
    }

    private async Task<string> CreateServicePrincipalAsync(TokenCredential cred, string appId, CancellationToken ct)
    {
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                var sp = await PostAsync(cred, "/servicePrincipals", new { appId }, ct).ConfigureAwait(false);
                return sp.GetProperty("id").GetString()!;
            }
            catch (HttpRequestException) when (attempt < 5)
            {
                await Task.Delay(TimeSpan.FromSeconds(5), ct).ConfigureAwait(false); // App noch nicht propagiert
            }
        }
    }

    // --- REST-Helfer ---

    private async Task<string> TokenAsync(TokenCredential cred, CancellationToken ct)
    {
        var t = await cred.GetTokenAsync(new TokenRequestContext(GraphScopes), ct).ConfigureAwait(false);
        return t.Token;
    }

    private async Task<JsonElement> PostAsync(TokenCredential cred, string path, object body, CancellationToken ct)
    {
        using var req = new HttpRequestMessage(HttpMethod.Post, Base + path)
        {
            Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
        };
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await TokenAsync(cred, ct).ConfigureAwait(false));
        using var resp = await _http.SendAsync(req, ct).ConfigureAwait(false);
        var text = await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
        if (!resp.IsSuccessStatusCode)
            throw new HttpRequestException($"POST {path} -> {(int)resp.StatusCode}: {text}");
        return JsonDocument.Parse(text).RootElement.Clone();
    }

    private async Task<JsonElement> GetAsync(TokenCredential cred, string path, CancellationToken ct)
    {
        using var req = new HttpRequestMessage(HttpMethod.Get, Base + path);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await TokenAsync(cred, ct).ConfigureAwait(false));
        using var resp = await _http.SendAsync(req, ct).ConfigureAwait(false);
        var text = await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
        if (!resp.IsSuccessStatusCode)
            throw new HttpRequestException($"GET {path} -> {(int)resp.StatusCode}: {text}");
        return JsonDocument.Parse(text).RootElement.Clone();
    }
}
