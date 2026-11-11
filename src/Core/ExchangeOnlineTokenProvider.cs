using Microsoft.Identity.Client;

namespace IMAP2M365.Core;

/// <summary>
/// App-only Token-Beschaffung via MSAL Client-Credentials.
/// Pro Resource-Scope wird eine eigene ConfidentialClientApplication-Tokenkette geführt;
/// MSAL cached intern und erneuert automatisch.
/// Für EXO-IMAP ist der Scope "https://outlook.office365.com/.default",
/// für Graph "https://graph.microsoft.com/.default".
/// Der XOAUTH2-Username ist NICHT in diesem Token kodiert – er wird beim Login
/// als Ziel-UPN mitgegeben (SaslMechanismOAuth2(upn, token)); das app-only-Token
/// berechtigt die App, im Namen genau der Postfächer zu handeln, die per
/// Add-MailboxPermission / ApplicationAccessPolicy freigegeben sind.
/// </summary>
public sealed class ExchangeOnlineTokenProvider : ITokenProvider
{
    private readonly IConfidentialClientApplication _app;

    public ExchangeOnlineTokenProvider(string tenantId, string clientId, string clientSecret)
    {
        _app = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
            .Build();
    }

    public async Task<string> GetTokenAsync(string scope, CancellationToken ct)
    {
        // .default fordert alle der App statisch zugewiesenen Application-Permissions an.
        var result = await _app
            .AcquireTokenForClient(new[] { scope })
            .ExecuteAsync(ct)
            .ConfigureAwait(false);

        return result.AccessToken;
    }

    public const string ExchangeImapScope = "https://outlook.office365.com/.default";
    public const string GraphScope        = "https://graph.microsoft.com/.default";
}
