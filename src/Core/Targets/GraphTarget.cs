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
        // Kanonischer erster Abschnitt -> Wellknown-Folder (sprachunabhängig), sonst unter msgfolderroot.
        string parent = "msgfolderroot";
        var start = 0;
        if (pathSegments.Count > 0 && WellKnown.TryGetValue(pathSegments[0], out var wk))
        {
            parent = wk;
            start = 1;
        }

        for (var i = start; i < pathSegments.Count; i++)
            parent = await FindOrCreateChildAsync(parent, pathSegments[i], ct).ConfigureAwait(false);
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

    // Graph begrenzt den Request-Body auf 4 MB. Unterhalb -> fidelitätstreuer MIME-Import.
    // Darüber -> Nachricht aus MIME-Teilen bauen + große Anhänge per Upload-Session.
    private const int MimeImportLimit = 3 * 1024 * 1024;       // Sicherheitsabstand unter 4 MB
    private const int InlineAttachmentLimit = 3 * 1024 * 1024; // größere Anhänge via Session
    private const int UploadChunk = 4 * 327680;                // 1,25 MiB, Vielfaches von 320 KiB

    public async Task AppendAsync(string folderHandle, SourceMessage message, CancellationToken ct)
    {
        await using var ms = new MemoryStream();
        await message.Mime.WriteToAsync(ms, ct).ConfigureAwait(false);

        if (ms.Length <= MimeImportLimit)
            await AppendViaMimeAsync(folderHandle, message, ms, ct).ConfigureAwait(false);
        else
            await AppendLargeAsync(folderHandle, message, ct).ConfigureAwait(false);
    }

    /// <summary>Kleine Nachrichten: roher MIME-Import, höchste Treue.</summary>
    private async Task AppendViaMimeAsync(string folderHandle, SourceMessage message, MemoryStream mime, CancellationToken ct)
    {
        var b64 = Convert.ToBase64String(mime.ToArray());
        using var resp = await SendAsync(() => new HttpRequestMessage(
            HttpMethod.Post, $"{Base}/users/{_userId}/mailFolders/{folderHandle}/messages")
        { Content = new StringContent(b64, Encoding.UTF8, "text/plain") }, ct).ConfigureAwait(false);

        resp.EnsureSuccessStatusCode();
        using var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false));
        var newId = doc.RootElement.GetProperty("id").GetString()!;

        var patch = JsonSerializer.Serialize(new
        {
            isRead = message.IsSeen,
            flag = new { flagStatus = message.IsFlagged ? "flagged" : "notFlagged" }
        });
        using var _ = await SendAsync(() => new HttpRequestMessage(
            HttpMethod.Patch, $"{Base}/users/{_userId}/messages/{newId}")
        { Content = new StringContent(patch, Encoding.UTF8, "application/json") }, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Große Nachrichten: Nachricht aus MIME-Teilen bauen (Betreff/Body/Empfänger),
    /// Empfangs-/Sendedatum über MAPI-Extended-Properties setzen, dann kleine Anhänge
    /// inline, große Anhänge per createUploadSession hochladen.
    /// Treue-Hinweis: exakte MIME-Struktur/Custom-Header gehen hier verloren (Graph-Limit).
    /// </summary>
    private async Task AppendLargeAsync(string folderHandle, SourceMessage message, CancellationToken ct)
    {
        var mime = message.Mime;

        var (inlineAtts, largeAtts) = SplitAttachments(mime);

        // MAPI: PR_MESSAGE_DELIVERY_TIME (0x0E06) und PR_CLIENT_SUBMIT_TIME (0x0039), beide PT_SYSTIME
        var receivedIso = message.InternalDate.UtcDateTime.ToString("o");
        var sentIso = (mime.Date == default ? message.InternalDate : mime.Date).UtcDateTime.ToString("o");

        var body = new
        {
            subject = mime.Subject ?? "",
            body = BuildBody(mime),
            from = ToRecipient(mime.From.Mailboxes.FirstOrDefault()),
            toRecipients = ToRecipients(mime.To.Mailboxes),
            ccRecipients = ToRecipients(mime.Cc.Mailboxes),
            bccRecipients = ToRecipients(mime.Bcc.Mailboxes),
            isRead = message.IsSeen,
            flag = new { flagStatus = message.IsFlagged ? "flagged" : "notFlagged" },
            internetMessageId = string.IsNullOrEmpty(mime.MessageId) ? null : $"<{mime.MessageId}>",
            singleValueExtendedProperties = new object[]
            {
                new { id = "SystemTime 0x0E06", value = receivedIso },
                new { id = "SystemTime 0x0039", value = sentIso }
            },
            attachments = inlineAtts // kleine Anhänge direkt mitgeben
        };

        var json = JsonSerializer.Serialize(body, new JsonSerializerOptions { DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull });
        using var resp = await SendAsync(() => new HttpRequestMessage(
            HttpMethod.Post, $"{Base}/users/{_userId}/mailFolders/{folderHandle}/messages")
        { Content = new StringContent(json, Encoding.UTF8, "application/json") }, ct).ConfigureAwait(false);

        resp.EnsureSuccessStatusCode();
        using var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false));
        var msgId = doc.RootElement.GetProperty("id").GetString()!;

        foreach (var (name, contentType, bytes) in largeAtts)
            await UploadLargeAttachmentAsync(msgId, name, contentType, bytes, ct).ConfigureAwait(false);
    }

    private async Task UploadLargeAttachmentAsync(string msgId, string name, string contentType, byte[] bytes, CancellationToken ct)
    {
        // 1) Upload-Session eröffnen
        var sessionBody = JsonSerializer.Serialize(new
        {
            AttachmentItem = new { attachmentType = "file", name, size = bytes.Length, contentType }
        });
        using var sessResp = await SendAsync(() => new HttpRequestMessage(
            HttpMethod.Post, $"{Base}/users/{_userId}/messages/{msgId}/attachments/createUploadSession")
        { Content = new StringContent(sessionBody, Encoding.UTF8, "application/json") }, ct).ConfigureAwait(false);
        sessResp.EnsureSuccessStatusCode();
        using var sdoc = JsonDocument.Parse(await sessResp.Content.ReadAsStringAsync(ct).ConfigureAwait(false));
        var uploadUrl = sdoc.RootElement.GetProperty("uploadUrl").GetString()!;

        // 2) In Blöcken (Vielfache von 320 KiB) hochladen; uploadUrl ist vorsigniert, kein Bearer nötig
        for (int offset = 0; offset < bytes.Length; offset += UploadChunk)
        {
            var len = Math.Min(UploadChunk, bytes.Length - offset);
            var slice = new ReadOnlyMemory<byte>(bytes, offset, len);

            using var put = new HttpRequestMessage(HttpMethod.Put, uploadUrl)
            {
                Content = new ReadOnlyMemoryContent(slice)
            };
            put.Content.Headers.ContentLength = len;
            put.Content.Headers.ContentRange = new ContentRangeHeaderValue(offset, offset + len - 1, bytes.Length);

            using var putResp = await _http.SendAsync(put, ct).ConfigureAwait(false);
            if (!putResp.IsSuccessStatusCode && putResp.StatusCode != HttpStatusCode.Created && (int)putResp.StatusCode != 202)
                putResp.EnsureSuccessStatusCode();
        }
    }

    private static (List<object> inline, List<(string name, string contentType, byte[] bytes)> large) SplitAttachments(MimeKit.MimeMessage mime)
    {
        var inline = new List<object>();
        var large = new List<(string, string, byte[])>();

        foreach (var att in mime.Attachments.OfType<MimeKit.MimePart>())
        {
            using var mem = new MemoryStream();
            att.Content.DecodeTo(mem);
            var bytes = mem.ToArray();
            var name = att.FileName ?? "attachment";
            var ctype = att.ContentType?.MimeType ?? "application/octet-stream";

            if (bytes.Length <= InlineAttachmentLimit)
                inline.Add(new Dictionary<string, object>
                {
                    ["@odata.type"] = "#microsoft.graph.fileAttachment",
                    ["name"] = name,
                    ["contentType"] = ctype,
                    ["contentBytes"] = Convert.ToBase64String(bytes)
                });
            else
                large.Add((name, ctype, bytes));
        }
        return (inline, large);
    }

    private static object BuildBody(MimeKit.MimeMessage mime)
        => !string.IsNullOrEmpty(mime.HtmlBody)
            ? new { contentType = "HTML", content = mime.HtmlBody }
            : new { contentType = "Text", content = mime.TextBody ?? "" };

    private static object[] ToRecipients(IEnumerable<MimeKit.MailboxAddress> addrs)
        => addrs.Select(a => ToRecipient(a)!).Where(r => r != null).ToArray();

    private static object? ToRecipient(MimeKit.MailboxAddress? a)
        => a is null ? null : new { emailAddress = new { address = a.Address, name = a.Name ?? a.Address } };

    public async IAsyncEnumerable<TargetMessage> ListAsync(
        string folderHandle, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct)
    {
        var url = $"{Base}/users/{_userId}/mailFolders/{folderHandle}/messages?$select=id,internetMessageId&$top=200";
        while (!string.IsNullOrEmpty(url))
        {
            var page = new List<TargetMessage>();
            string next;

            using (var resp = await SendAsync(() => new HttpRequestMessage(HttpMethod.Get, url), ct).ConfigureAwait(false))
            {
                resp.EnsureSuccessStatusCode();
                var body = await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
                using var doc = JsonDocument.Parse(body);
                foreach (var m in doc.RootElement.GetProperty("value").EnumerateArray())
                {
                    var id = m.GetProperty("id").GetString() ?? "";
                    var mid = m.TryGetProperty("internetMessageId", out var x) ? x.GetString() ?? "" : "";
                    page.Add(new TargetMessage(mid, id));
                }
                next = doc.RootElement.TryGetProperty("@odata.nextLink", out var nl) ? nl.GetString() ?? "" : "";
            }

            foreach (var tm in page) yield return tm;
            url = next;
        }
    }

    public async Task DeleteAsync(string folderHandle, string deleteHandle, CancellationToken ct)
    {
        // DELETE auf eine Nachricht verschiebt sie nach „Gelöschte Elemente" (wiederherstellbar).
        using var resp = await SendAsync(() => new HttpRequestMessage(
            HttpMethod.Delete, $"{Base}/users/{_userId}/messages/{deleteHandle}"), ct).ConfigureAwait(false);
        if (!resp.IsSuccessStatusCode && resp.StatusCode != HttpStatusCode.NoContent)
            resp.EnsureSuccessStatusCode();
    }

    public ValueTask DisposeAsync()
    {
        _http.Dispose();
        return ValueTask.CompletedTask;
    }
}
