using System.Net;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Xml.Linq;

namespace IMAP2M365.Ews;

/// <summary>Fehler aus einer EWS-Antwort (ResponseCode + Klartext).</summary>
public sealed class EwsException : Exception
{
    public string Code { get; }
    public EwsException(string code, string message) : base(message) => Code = code;
}

public sealed record EwsFolder(string Id, string ParentId, string DisplayName, string FolderClass, int TotalCount);

public sealed record EwsItemInfo(
    string Id, string InternetMessageId, bool IsRead, bool IsDraft,
    DateTimeOffset? Received, long Size, int FlagStatus, int LastVerb);

/// <summary>
/// Schlanker EWS-Client (SOAP über HttpClient, keine Fremdbibliothek) für Exchange on-prem.
/// Anmeldung mit einem Admin-/Dienstkonto (Basic/NTLM/Negotiate, je nach Server); der Zugriff auf
/// das jeweilige Postfach läuft über ApplicationImpersonation (<see cref="ImpersonatedSmtp"/>).
/// </summary>
public sealed class EwsClient : IDisposable
{
    public static readonly XNamespace Soap = "http://schemas.xmlsoap.org/soap/envelope/";
    public static readonly XNamespace T = "http://schemas.microsoft.com/exchange/services/2006/types";
    public static readonly XNamespace M = "http://schemas.microsoft.com/exchange/services/2006/messages";

    /// <summary>Kanonische Ordnernamen der Engine -> EWS-DistinguishedFolderId.</summary>
    public static readonly IReadOnlyDictionary<string, string> CanonicalToDistinguished =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["INBOX"] = "inbox",
            ["Sent Items"] = "sentitems",
            ["Drafts"] = "drafts",
            ["Deleted Items"] = "deleteditems",
            ["Junk Email"] = "junkemail"
        };

    private const int MaxRetries = 5;

    private readonly HttpClient _http;
    private readonly string _url;

    /// <summary>SMTP-Adresse des Postfachs, in dessen Namen gearbeitet wird (leer = eigenes Postfach).</summary>
    public string? ImpersonatedSmtp { get; set; }

    public EwsClient(string ewsUrl, string user, string password, bool acceptAnyCertificate)
    {
        _url = ewsUrl;

        // "DOMAIN\user" in Domäne + Benutzer zerlegen; "user@domain" bleibt wie es ist.
        var bs = user.IndexOf('\\');
        var cred = bs > 0
            ? new NetworkCredential(user[(bs + 1)..], password, user[..bs])
            : new NetworkCredential(user, password);

        var handler = new HttpClientHandler { Credentials = cred, PreAuthenticate = true };
        if (acceptAnyCertificate)
            handler.ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator;

        _http = new HttpClient(handler) { Timeout = TimeSpan.FromMinutes(10) };
    }

    public static string DefaultUrl(string server) => $"https://{server.Trim()}/EWS/Exchange.asmx";

    private static string Esc(string s) => SecurityElement.Escape(s) ?? "";

    private string Envelope(string body)
    {
        var imp = string.IsNullOrWhiteSpace(ImpersonatedSmtp)
            ? ""
            : $"<t:ExchangeImpersonation><t:ConnectingSID><t:SmtpAddress>{Esc(ImpersonatedSmtp!)}</t:SmtpAddress></t:ConnectingSID></t:ExchangeImpersonation>";

        return "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
               "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" " +
               "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" " +
               "xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\">" +
               "<soap:Header><t:RequestServerVersion Version=\"Exchange2010_SP2\"/>" + imp + "</soap:Header>" +
               "<soap:Body>" + body + "</soap:Body></soap:Envelope>";
    }

    /// <summary>
    /// Sendet einen EWS-Aufruf und gibt das Antwortdokument zurück. Bei ResponseClass="Error" wird
    /// eine <see cref="EwsException"/> geworfen (außer throwOnError=false). Drosselung
    /// (ErrorServerBusy) wird mit der vom Server genannten Wartezeit wiederholt.
    /// </summary>
    public async Task<XDocument> CallAsync(string body, CancellationToken ct, bool throwOnError = true)
    {
        for (var attempt = 0; ; attempt++)
        {
            using var req = new HttpRequestMessage(HttpMethod.Post, _url)
            {
                Content = new StringContent(Envelope(body), Encoding.UTF8, "text/xml")
            };
            if (!string.IsNullOrWhiteSpace(ImpersonatedSmtp))
                req.Headers.TryAddWithoutValidation("X-AnchorMailbox", ImpersonatedSmtp);

            using var resp = await _http.SendAsync(req, ct).ConfigureAwait(false);
            var text = await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false);

            if (resp.StatusCode == HttpStatusCode.Unauthorized)
                throw new EwsException("Unauthorized", "Anmeldung am EWS abgelehnt (HTTP 401) – Benutzer/Kennwort prüfen.");

            XDocument doc;
            try { doc = XDocument.Parse(text); }
            catch
            {
                throw new EwsException("Http" + (int)resp.StatusCode,
                    $"EWS antwortet nicht mit SOAP (HTTP {(int)resp.StatusCode}) – EWS-URL prüfen.");
            }

            // SOAP-Fault (z. B. Impersonation verweigert, Schemafehler)
            var fault = doc.Descendants(Soap + "Fault").FirstOrDefault();
            if (fault is not null)
            {
                var code = fault.Descendants().FirstOrDefault(e => e.Name.LocalName == "ResponseCode")?.Value ?? "SoapFault";
                var msg = fault.Descendants().FirstOrDefault(e => e.Name.LocalName == "faultstring")?.Value ?? "SOAP-Fehler";
                if (code == "ErrorServerBusy" && attempt < MaxRetries)
                {
                    await Task.Delay(BackOff(fault, attempt), ct).ConfigureAwait(false);
                    continue;
                }
                throw new EwsException(code, msg);
            }

            var error = doc.Descendants().FirstOrDefault(e => (string?)e.Attribute("ResponseClass") == "Error");
            if (error is not null)
            {
                var code = error.Element(M + "ResponseCode")?.Value ?? "Error";
                if (code == "ErrorServerBusy" && attempt < MaxRetries)
                {
                    await Task.Delay(BackOff(error, attempt), ct).ConfigureAwait(false);
                    continue;
                }
                if (throwOnError)
                    throw new EwsException(code, error.Element(M + "MessageText")?.Value ?? code);
            }

            return doc;
        }
    }

    private static TimeSpan BackOff(XElement scope, int attempt)
    {
        var ms = scope.Descendants()
            .FirstOrDefault(e => (string?)e.Attribute("Name") == "BackOffMilliseconds")?.Value;
        return int.TryParse(ms, out var v) && v > 0
            ? TimeSpan.FromMilliseconds(Math.Min(v, 300_000))
            : TimeSpan.FromSeconds(5 * (attempt + 1));
    }

    // ---------- Ordner ----------

    /// <summary>Ids der Standardordner (msgfolderroot, inbox, sentitems, drafts, deleteditems, junkemail).</summary>
    public async Task<Dictionary<string, string>> GetWellKnownFolderIdsAsync(CancellationToken ct)
    {
        string[] names = { "msgfolderroot", "inbox", "sentitems", "drafts", "deleteditems", "junkemail" };
        var ids = string.Concat(names.Select(n => $"<t:DistinguishedFolderId Id=\"{n}\"/>"));
        var doc = await CallAsync(
            "<m:GetFolder><m:FolderShape><t:BaseShape>IdOnly</t:BaseShape></m:FolderShape>" +
            $"<m:FolderIds>{ids}</m:FolderIds></m:GetFolder>", ct, throwOnError: false).ConfigureAwait(false);

        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var messages = doc.Descendants(M + "GetFolderResponseMessage").ToList();
        for (var i = 0; i < messages.Count && i < names.Length; i++)
        {
            var id = messages[i].Descendants(T + "FolderId").FirstOrDefault()?.Attribute("Id")?.Value;
            if (!string.IsNullOrEmpty(id)) result[names[i]] = id;
        }

        if (!result.ContainsKey("msgfolderroot"))
        {
            // Der erste Eintrag trägt den eigentlichen Grund (z. B. ErrorImpersonateUserDenied).
            var first = messages.FirstOrDefault();
            throw new EwsException(
                first?.Element(M + "ResponseCode")?.Value ?? "Error",
                first?.Element(M + "MessageText")?.Value ?? "Postfach nicht erreichbar.");
        }
        return result;
    }

    /// <summary>Alle Ordner unterhalb der Postfachwurzel (tiefe Suche, seitenweise).</summary>
    public async Task<List<EwsFolder>> FindFoldersDeepAsync(CancellationToken ct)
    {
        var all = new List<EwsFolder>();
        for (var offset = 0; ;)
        {
            var doc = await CallAsync(
                "<m:FindFolder Traversal=\"Deep\"><m:FolderShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>" +
                "<t:FieldURI FieldURI=\"folder:DisplayName\"/><t:FieldURI FieldURI=\"folder:ParentFolderId\"/>" +
                "<t:FieldURI FieldURI=\"folder:FolderClass\"/><t:FieldURI FieldURI=\"folder:TotalCount\"/>" +
                "</t:AdditionalProperties></m:FolderShape>" +
                $"<m:IndexedPageFolderView MaxEntriesReturned=\"500\" Offset=\"{offset}\" BasePoint=\"Beginning\"/>" +
                "<m:ParentFolderIds><t:DistinguishedFolderId Id=\"msgfolderroot\"/></m:ParentFolderIds></m:FindFolder>",
                ct).ConfigureAwait(false);

            var root = doc.Descendants(M + "RootFolder").FirstOrDefault();
            var folders = root?.Element(T + "Folders")?.Elements().ToList() ?? new List<XElement>();
            foreach (var f in folders)
            {
                // Nur echte Ordner (keine Such-, Kalender-, Kontakt- oder Aufgabenordner).
                if (f.Name != T + "Folder") continue;
                all.Add(new EwsFolder(
                    f.Element(T + "FolderId")?.Attribute("Id")?.Value ?? "",
                    f.Element(T + "ParentFolderId")?.Attribute("Id")?.Value ?? "",
                    f.Element(T + "DisplayName")?.Value ?? "",
                    f.Element(T + "FolderClass")?.Value ?? "",
                    int.TryParse(f.Element(T + "TotalCount")?.Value, out var n) ? n : 0));
            }

            offset += folders.Count;
            var last = (string?)root?.Attribute("IncludesLastItemInRange") != "false";
            if (last || folders.Count == 0) break;
        }
        return all;
    }

    public async Task<string?> FindChildFolderAsync(string parentId, string displayName, CancellationToken ct)
    {
        var doc = await CallAsync(
            "<m:FindFolder Traversal=\"Shallow\"><m:FolderShape><t:BaseShape>IdOnly</t:BaseShape></m:FolderShape>" +
            "<m:Restriction><t:IsEqualTo><t:FieldURI FieldURI=\"folder:DisplayName\"/>" +
            $"<t:FieldURIOrConstant><t:Constant Value=\"{Esc(displayName)}\"/></t:FieldURIOrConstant></t:IsEqualTo></m:Restriction>" +
            $"<m:ParentFolderIds><t:FolderId Id=\"{Esc(parentId)}\"/></m:ParentFolderIds></m:FindFolder>",
            ct).ConfigureAwait(false);
        return doc.Descendants(T + "FolderId").FirstOrDefault()?.Attribute("Id")?.Value;
    }

    public async Task<string> CreateFolderAsync(string parentId, string displayName, CancellationToken ct)
    {
        var doc = await CallAsync(
            $"<m:CreateFolder><m:ParentFolderId><t:FolderId Id=\"{Esc(parentId)}\"/></m:ParentFolderId>" +
            $"<m:Folders><t:Folder><t:FolderClass>IPF.Note</t:FolderClass><t:DisplayName>{Esc(displayName)}</t:DisplayName></t:Folder></m:Folders>" +
            "</m:CreateFolder>", ct).ConfigureAwait(false);
        return doc.Descendants(T + "FolderId").FirstOrDefault()?.Attribute("Id")?.Value
               ?? throw new EwsException("Error", $"Ordner '{displayName}' konnte nicht angelegt werden.");
    }

    // ---------- Nachrichten ----------

    /// <summary>Listet die Elemente eines Ordners (älteste zuerst, seitenweise).</summary>
    public async IAsyncEnumerable<EwsItemInfo> FindItemsAsync(string folderId, [EnumeratorCancellation] CancellationToken ct)
    {
        for (var offset = 0; ;)
        {
            var doc = await CallAsync(
                "<m:FindItem Traversal=\"Shallow\"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>" +
                "<t:FieldURI FieldURI=\"item:Size\"/><t:FieldURI FieldURI=\"item:DateTimeReceived\"/><t:FieldURI FieldURI=\"item:IsDraft\"/>" +
                "<t:FieldURI FieldURI=\"message:InternetMessageId\"/><t:FieldURI FieldURI=\"message:IsRead\"/>" +
                "<t:ExtendedFieldURI PropertyTag=\"0x1090\" PropertyType=\"Integer\"/>" +
                "<t:ExtendedFieldURI PropertyTag=\"0x1081\" PropertyType=\"Integer\"/>" +
                "</t:AdditionalProperties></m:ItemShape>" +
                $"<m:IndexedPageItemView MaxEntriesReturned=\"200\" Offset=\"{offset}\" BasePoint=\"Beginning\"/>" +
                "<m:SortOrder><t:FieldOrder Order=\"Ascending\"><t:FieldURI FieldURI=\"item:DateTimeReceived\"/></t:FieldOrder></m:SortOrder>" +
                $"<m:ParentFolderIds><t:FolderId Id=\"{Esc(folderId)}\"/></m:ParentFolderIds></m:FindItem>",
                ct).ConfigureAwait(false);

            var root = doc.Descendants(M + "RootFolder").FirstOrDefault();
            var items = root?.Element(T + "Items")?.Elements().ToList() ?? new List<XElement>();

            var page = new List<EwsItemInfo>(items.Count);
            foreach (var it in items)
            {
                var id = it.Element(T + "ItemId")?.Attribute("Id")?.Value;
                if (string.IsNullOrEmpty(id)) continue;
                page.Add(new EwsItemInfo(
                    id,
                    it.Element(T + "InternetMessageId")?.Value ?? "",
                    it.Element(T + "IsRead")?.Value == "true",
                    it.Element(T + "IsDraft")?.Value == "true",
                    DateTimeOffset.TryParse(it.Element(T + "DateTimeReceived")?.Value, out var d) ? d : null,
                    long.TryParse(it.Element(T + "Size")?.Value, out var s) ? s : 0,
                    ExtendedInt(it, "0x1090"),
                    ExtendedInt(it, "0x1081")));
            }

            foreach (var p in page) yield return p;

            offset += items.Count;
            var last = (string?)root?.Attribute("IncludesLastItemInRange") != "false";
            if (last || items.Count == 0) yield break;
        }
    }

    private static int ExtendedInt(XElement item, string tag)
    {
        foreach (var ep in item.Elements(T + "ExtendedProperty"))
        {
            var uri = ep.Element(T + "ExtendedFieldURI");
            if (string.Equals((string?)uri?.Attribute("PropertyTag"), tag, StringComparison.OrdinalIgnoreCase))
                return int.TryParse(ep.Element(T + "Value")?.Value, out var v) ? v : 0;
        }
        return 0;
    }

    public async Task<bool> ExistsByMessageIdAsync(string folderId, string messageId, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(messageId)) return false;
        var id = messageId.StartsWith('<') ? messageId : $"<{messageId}>";
        var doc = await CallAsync(
            "<m:FindItem Traversal=\"Shallow\"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape></m:ItemShape>" +
            "<m:IndexedPageItemView MaxEntriesReturned=\"1\" Offset=\"0\" BasePoint=\"Beginning\"/>" +
            "<m:Restriction><t:IsEqualTo><t:FieldURI FieldURI=\"message:InternetMessageId\"/>" +
            $"<t:FieldURIOrConstant><t:Constant Value=\"{Esc(id)}\"/></t:FieldURIOrConstant></t:IsEqualTo></m:Restriction>" +
            $"<m:ParentFolderIds><t:FolderId Id=\"{Esc(folderId)}\"/></m:ParentFolderIds></m:FindItem>",
            ct).ConfigureAwait(false);
        return doc.Descendants(T + "ItemId").Any();
    }

    /// <summary>Holt den vollständigen MIME-Inhalt eines Elements.</summary>
    public async Task<byte[]> GetMimeAsync(string itemId, CancellationToken ct)
    {
        var doc = await CallAsync(
            "<m:GetItem><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:IncludeMimeContent>true</t:IncludeMimeContent></m:ItemShape>" +
            $"<m:ItemIds><t:ItemId Id=\"{Esc(itemId)}\"/></m:ItemIds></m:GetItem>", ct).ConfigureAwait(false);
        var b64 = doc.Descendants(T + "MimeContent").FirstOrDefault()?.Value
                  ?? throw new EwsException("NoMime", "Element liefert keinen MIME-Inhalt.");
        return Convert.FromBase64String(b64);
    }

    /// <summary>
    /// Legt eine Nachricht aus MIME im Ordner an. Gelesen-Status, Entwurf, Kennzeichnung sowie
    /// Empfangs-/Sendedatum werden über MAPI-Eigenschaften gesetzt, damit sie erhalten bleiben.
    /// </summary>
    public async Task CreateMimeItemAsync(
        string folderId, byte[] mime, bool isRead, bool isDraft, bool flagged,
        DateTimeOffset received, DateTimeOffset sent, CancellationToken ct)
    {
        // PR_MESSAGE_FLAGS: 1 = gelesen, 8 = ungesendet (Entwurf). Ohne 8 gilt die Mail als zugestellt.
        var flags = (isRead ? 1 : 0) | (isDraft ? 8 : 0);
        var sb = new StringBuilder();
        sb.Append("<m:CreateItem MessageDisposition=\"SaveOnly\">");
        sb.Append($"<m:SavedItemFolderId><t:FolderId Id=\"{Esc(folderId)}\"/></m:SavedItemFolderId>");
        sb.Append("<m:Items><t:Message>");
        sb.Append("<t:MimeContent CharacterSet=\"UTF-8\">").Append(Convert.ToBase64String(mime)).Append("</t:MimeContent>");
        sb.Append(ExtProp("0x0E07", "Integer", flags.ToString()));
        sb.Append(ExtProp("0x0E06", "SystemTime", Iso(received)));
        sb.Append(ExtProp("0x0039", "SystemTime", Iso(sent)));
        if (flagged) sb.Append(ExtProp("0x1090", "Integer", "2"));
        sb.Append("</t:Message></m:Items></m:CreateItem>");

        await CallAsync(sb.ToString(), ct).ConfigureAwait(false);
    }

    private static string ExtProp(string tag, string type, string value) =>
        $"<t:ExtendedProperty><t:ExtendedFieldURI PropertyTag=\"{tag}\" PropertyType=\"{type}\"/><t:Value>{Esc(value)}</t:Value></t:ExtendedProperty>";

    private static string Iso(DateTimeOffset d) => d.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ");

    /// <summary>Verschiebt ein Element nach "Gelöschte Elemente" (wiederherstellbar).</summary>
    public async Task DeleteItemAsync(string itemId, CancellationToken ct)
    {
        await CallAsync(
            "<m:DeleteItem DeleteType=\"MoveToDeletedItems\">" +
            $"<m:ItemIds><t:ItemId Id=\"{Esc(itemId)}\"/></m:ItemIds></m:DeleteItem>", ct).ConfigureAwait(false);
    }

    public void Dispose() => _http.Dispose();
}
