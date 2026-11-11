using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace IMAP2M365.Gui;

/// <summary>
/// Legt die Entra-App-Registrierung für app-only XOAUTH2 an und erteilt Admin-Consent.
/// Admin meldet sich interaktiv per Device-Code an (einmalig) – danach läuft alles über Graph.
///
/// Benötigte delegierte Rechte des anmeldenden Admins:
///   Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All (Global Admin / Cloud App Admin hat sie).
/// </summary>
public sealed class GraphProvisioner
{
    // Ressourcen-AppIds (konstant in jedem Tenant)
    private const string ExchangeOnlineResourceAppId = "00000002-0000-0ff1-ce00-000000000000"; // Office 365 Exchange Online
    private const string GraphResourceAppId          = "00000003-0000-0000-c000-000000000000"; // Microsoft Graph

    // AppRole-IDs
    private const string ImapAccessAsApp = "a4b8392a-d8d1-4954-a029-8e668a39a170"; // Office365 EXO: IMAP.AccessAsApp (App)
    private const string MailReadWrite   = "e2a3a72e-5f79-4c64-b1b1-878b674786c9"; // Graph: Mail.ReadWrite (App)

    public Action<string>? Log { get; set; }

    private GraphServiceClient Client(string tenantId)
    {
        var cred = new DeviceCodeCredential(new DeviceCodeCredentialOptions
        {
            TenantId = tenantId,
            ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e", // Microsoft Graph PowerShell (öffentlicher Client)
            DeviceCodeCallback = (info, ct) => { Log?.Invoke(info.Message); return Task.CompletedTask; }
        });
        return new GraphServiceClient(cred, new[] { "https://graph.microsoft.com/.default" });
    }

    /// <summary>
    /// Erstellt App + ServicePrincipal, setzt die benötigten Application-Permissions,
    /// erzeugt ein Client-Secret und erteilt Admin-Consent (appRoleAssignment).
    /// Gibt (AppId, ObjectId des ServicePrincipals, Secret) zurück – wandert in config.enc.
    /// </summary>
    public async Task<(string appId, string spObjectId, string secret)> ProvisionAsync(
        string tenantId, string displayName, bool includeGraphTarget, CancellationToken ct)
    {
        var graph = Client(tenantId);

        // 1) App-Registrierung mit requiredResourceAccess
        var resources = new List<RequiredResourceAccess>
        {
            new()
            {
                ResourceAppId = ExchangeOnlineResourceAppId,
                ResourceAccess = new() { new() { Id = Guid.Parse(ImapAccessAsApp), Type = "Role" } }
            }
        };
        if (includeGraphTarget)
            resources.Add(new()
            {
                ResourceAppId = GraphResourceAppId,
                ResourceAccess = new() { new() { Id = Guid.Parse(MailReadWrite), Type = "Role" } }
            });

        var app = await graph.Applications.PostAsync(new Application
        {
            DisplayName = displayName,
            SignInAudience = "AzureADMyOrg",
            RequiredResourceAccess = resources
        }, cancellationToken: ct) ?? throw new InvalidOperationException("App-Registrierung fehlgeschlagen.");
        Log?.Invoke($"App angelegt: {app.AppId}");

        // 2) Client-Secret
        var secret = await graph.Applications[app.Id].AddPassword.PostAsync(
            new Microsoft.Graph.Applications.Item.AddPassword.AddPasswordPostRequestBody
            {
                PasswordCredential = new PasswordCredential
                {
                    DisplayName = "IMAP2M365",
                    EndDateTime = DateTimeOffset.UtcNow.AddYears(1)
                }
            }, cancellationToken: ct);
        Log?.Invoke("Client-Secret erzeugt (1 Jahr gültig).");

        // 3) ServicePrincipal zur App
        var sp = await graph.ServicePrincipals.PostAsync(new ServicePrincipal { AppId = app.AppId }, cancellationToken: ct)
                 ?? throw new InvalidOperationException("ServicePrincipal-Anlage fehlgeschlagen.");

        // 4) Admin-Consent = appRoleAssignment je Ressource
        await GrantConsentAsync(graph, sp.Id!, ExchangeOnlineResourceAppId, ImapAccessAsApp, ct);
        if (includeGraphTarget)
            await GrantConsentAsync(graph, sp.Id!, GraphResourceAppId, MailReadWrite, ct);

        Log?.Invoke("Admin-Consent erteilt.");
        return (app.AppId!, sp.Id!, secret?.SecretText ?? "");
    }

    private async Task GrantConsentAsync(
        GraphServiceClient graph, string clientSpId, string resourceAppId, string appRoleId, CancellationToken ct)
    {
        // ServicePrincipal der Ressource (z. B. EXO) im Tenant finden
        var res = await graph.ServicePrincipals.GetAsync(r =>
        {
            r.QueryParameters.Filter = $"appId eq '{resourceAppId}'";
            r.QueryParameters.Top = 1;
        }, ct);
        var resourceSp = res?.Value?.FirstOrDefault()
            ?? throw new InvalidOperationException($"Ressourcen-ServicePrincipal {resourceAppId} nicht gefunden.");

        await graph.ServicePrincipals[clientSpId].AppRoleAssignments.PostAsync(new AppRoleAssignment
        {
            PrincipalId = Guid.Parse(clientSpId),
            ResourceId = Guid.Parse(resourceSp.Id!),
            AppRoleId = Guid.Parse(appRoleId)
        }, cancellationToken: ct);
    }
}
