using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;
using IMAP2M365.Core;

namespace IMAP2M365.Targets;

/// <summary>
/// Ziel = Exchange Online via Microsoft Graph (app-only, Mail.ReadWrite).
/// Nachrichten werden als MIME importiert (POST .../messages, Content-Type text/plain,
/// base64), damit receivedDateTime aus den Headern erhalten bleibt; isRead/flag werden
/// anschließend per PATCH gesetzt. Throttling (429) wird mit Retry-After abgefangen.
/// </summary>
public sealed class GraphTarget : IMailTarget
{
    private const string Base = "https://graph.microsoft.com/v1.0";

    private readonly TokenCredential _cred;
    private readonly HttpClient _http = new();
    private string _userId = "";

    // Kanonische Ordnernamen -> Graph-Wellknown-Folder
    private static readonly Dictionary<string, string> WellKnown = new(StringComparer.OrdinalIgnoreCase)
    {
        ["INBOX"] = "inbox",
        ["Sent Items"] = "sentitems",
        ["Drafts"] = "drafts",
        ["Deleted Items"] = "deleteditems",
        ["Junk Email"] = "junkemail",
        ["Archive"] = "archive"
    };

    public GraphTarget(string tenantId, string clientId, string clientSecret)
    {
        _cred = new ClientSecretCredential(tenantId, clientId, clientSecret);
    }

    public Task ConnectAsync(string targetMailbox, CancellationToken ct)
    {
        _userId = targetMailbox; // Graph adressiert Postfächer per UPN/SMTP
        return Task.CompletedTask;
    }

    private async Task<string> TokenAsync(CancellationToken ct)
    {
        var token = await _cred.GetTokenAsync(
            new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }), ct).ConfigureAwait(false);
        return token.Token;
    }

    /// <summary>Sendet eine Anfrage mit Bearer-Token und 429-Retry (Retry-After beachtet).</summary>
    private async Task<HttpResponseMessage> SendAsync(
        Func<HttpRequestMessage> build, CancellationToken ct, int maxRetries = 5)
    {
        for (int attempt = 0; ; attempt++)
        {
            var req = build();
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await TokenAsync(ct).ConfigureAwait(false));
            var resp = await _http.SendAsync(req, ct).ConfigureAwait(false);

            if (resp.StatusCode != (HttpStatusCode)429 && resp.StatusCode != HttpStatusCode.ServiceUnavailable)
                return resp;

            if (attempt >= maxRetries)
                return resp;

            var wait = resp.Headers.RetryAfter?.Delta
                       ?? TimeSpan.FromSeconds(Math.Min(60, Math.Pow(2, attempt)));
            resp.Dispose();
            await Task.Delay(wait, ct).ConfigureAwait(false);
        }
    }

    public async Task<string> EnsureFolderAsync(IReadOnlyList<string> pathSegments, CancellationToken ct)
    {
        // Einzelner kanonischer Name -> Wellknown-Folder direkt verwenden
        if (pathSegments.Count == 1 && WellKnown.TryGetValue(pathSegments[0], out var wk))
            return wk;

        // Sonst Pfad unterhalb msgfolderroot auflösen/anlegen
        string parent = "msgfolderroot";
        foreach (var segment in pathSegments)
        {
            parent = await FindOrCreateChildAsync(parent, segment, ct).ConfigureAwait(false);
        }
        return parent;
    }

    private async Task<string> FindOrCreateChildAsync(string parentId, string displayName, CancellationToken ct)
    {
        var filter = Uri.EscapeDataString($"displayName eq '{displayName.Replace("'", "''")}'");
        using var findResp = await SendAsync(() => new HttpRequestMessage(
            HttpMethod.Get, $"{Base}/users/{_userId}/mailFolders/{parentId}/childFolders?$filter={filter}&$select=id&$top=1"), ct)
            .ConfigureAwait(false);

        if (findResp.IsSuccessStatusCode)
        {
            using var doc = JsonDocument.Parse(await findResp.Content.ReadAsStringAsync(ct).ConfigureAwait(false));
            var arr = doc.RootElement.GetProperty("value");
            if (arr.GetArrayLength() > 0)
                return arr[0].GetProperty("id").GetString()!;
        }

        // anlegen
        var body = JsonSerializer.Serialize(new { displayName });
        using var createResp = await SendAsync(() => new HttpRequestMessage(
            HttpMethod.Post, $"{Base}/users/{_userId}/mailFolders/{parentId}/childFolders")
        { Content = new StringContent(body, Encoding.UTF8, "application/json") }, ct).ConfigureAwait(false);

        createResp.EnsureSuccessStatusCode();
        using var cdoc = JsonDocument.Parse(await createResp.Content.ReadAsStringAsync(ct).ConfigureAwait(false));
        return cdoc.RootElement.GetProperty("id").GetString()!;
    }

    public async Task<bool> ExistsAsync(string folderHandle, string messageId, CancellationToken ct)
    {
        if (string.IsNullOrEmpty(messageId)) return false;
        var id = messageId.StartsWith('<') ? messageId : $"<{messageId}>";
        var filter = Uri.EscapeDataString($"internetMessageId eq '{id.Replace("'", "''")}'");

        using var resp = await SendAsync(() => new HttpRequestMessage(
            HttpMethod.Get, $"{Base}/users/{_userId}/mailFolders/{folderHandle}/messages?$filter={filter}&$select=id&$top=1"), ct)
            .ConfigureAwait(false);

        if (!resp.IsSuccessStatusCode) return false;
        using var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false));
        return doc.RootElement.GetProperty("value").GetArrayLength() > 0;
    }

    public async Task AppendAsync(string folderHandle, SourceMessage message, CancellationToken ct)
    {
        // MIME serialisieren und base64-kodiert als text/plain importieren
        await using var ms = new MemoryStream();
        await message.Mime.WriteToAsync(ms, ct).ConfigureAwait(false);
        var b64 = Convert.ToBase64String(ms.ToArray());

        using var resp = await SendAsync(() => new HttpRequestMessage(
            HttpMethod.Post, $"{Base}/users/{_userId}/mailFolders/{folderHandle}/messages")
        { Content = new StringContent(b64, Encoding.UTF8, "text/plain") }, ct).ConfigureAwait(false);

        resp.EnsureSuccessStatusCode();
        using var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false));
        var newId = doc.RootElement.GetProperty("id").GetString()!;

        // isRead und Kennzeichnung nachziehen (receivedDateTime kommt aus dem MIME selbst)
        var patch = JsonSerializer.Serialize(new
        {
            isRead = message.IsSeen,
            flag = new { flagStatus = message.IsFlagged ? "flagged" : "notFlagged" }
        });
        using var patchResp = await SendAsync(() => new HttpRequestMessage(
            HttpMethod.Patch, $"{Base}/users/{_userId}/messages/{newId}")
        { Content = new StringContent(patch, Encoding.UTF8, "application/json") }, ct).ConfigureAwait(false);
        // PATCH ist best-effort; ein Fehler hier soll die Migration nicht abbrechen.
    }

    public ValueTask DisposeAsync()
    {
        _http.Dispose();
        return ValueTask.CompletedTask;
    }
}
