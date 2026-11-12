using IMAP2M365.Core;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;

namespace IMAP2M365.Targets;

/// <summary>
/// Ziel = IMAP-Server per APPEND. Zwei Betriebsarten:
///  - Exchange Online mit app-only XOAUTH2 (Postfach = SASL-Benutzer, Token der App),
///  - beliebiger IMAP-Server mit Benutzer/Kennwort (für den Weg aus der Cloud zurück).
/// Überträgt rohes MIME 1:1; Flags und INTERNALDATE (Empfangsdatum) bleiben erhalten.
/// </summary>
public sealed class ImapAppendTarget : IMailTarget
{
    public const string ExchangeOnlineHost = "outlook.office365.com";

    private readonly ITokenProvider? _tokens;
    private readonly string _host;
    private readonly int _port;
    private readonly bool _useSsl;
    private readonly string _password = "";
    private readonly ImapClient _client = new();
    private string _mailbox = "";

    /// <summary>Exchange Online, app-only XOAUTH2.</summary>
    public ImapAppendTarget(ITokenProvider tokens)
    {
        _tokens = tokens;
        _host = ExchangeOnlineHost;
        _port = 993;
        _useSsl = true;
    }

    /// <summary>Beliebiger IMAP-Server mit Benutzer/Kennwort (Benutzer = Zielpostfach aus ConnectAsync).</summary>
    public ImapAppendTarget(string host, int port, bool useSsl, string password, bool acceptAnyCertificate = false)
    {
        _host = host;
        _port = port;
        _useSsl = useSsl;
        _password = password;
        if (acceptAnyCertificate)
            _client.ServerCertificateValidationCallback = (_, _, _, _) => true;
    }

    public async Task ConnectAsync(string targetMailbox, CancellationToken ct)
    {
        _mailbox = targetMailbox;
        var opt = _useSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTlsWhenAvailable;
        await _client.ConnectAsync(_host, _port, opt, ct).ConfigureAwait(false);

        if (_tokens is not null)
        {
            // app-only: UPN des Zielpostfachs + App-Token. Erfordert vorab registrierten
            // ServicePrincipal + Mailbox-Permission in Exchange.
            var token = await _tokens.GetTokenAsync(ExchangeOnlineTokenProvider.ExchangeImapScope, ct)
                .ConfigureAwait(false);
            await _client.AuthenticateAsync(new SaslMechanismOAuth2(targetMailbox, token), ct).ConfigureAwait(false);
        }
        else
        {
            await _client.AuthenticateAsync(targetMailbox, _password, ct).ConfigureAwait(false);
        }
    }

    public async Task<string> EnsureFolderAsync(IReadOnlyList<string> pathSegments, CancellationToken ct)
    {
        IMailFolder current = _client.GetFolder(_client.PersonalNamespaces[0]);
        var start = 0;

        // Kanonischer erster Abschnitt (INBOX, Sent Items …) -> echter Standardordner des Servers,
        // unabhängig von dessen Sprache/Benennung (SPECIAL-USE). Sonst normal per Name.
        if (pathSegments.Count > 0 && ResolveSpecial(pathSegments[0]) is { } special)
        {
            current = special;
            start = 1;
        }

        for (var i = start; i < pathSegments.Count; i++)
        {
            IMailFolder child;
            try
            {
                child = await current.GetSubfolderAsync(pathSegments[i], ct).ConfigureAwait(false);
            }
            catch (FolderNotFoundException)
            {
                child = await current.CreateAsync(pathSegments[i], isMessageFolder: true, ct).ConfigureAwait(false);
            }
            current = child;
        }
        return current.FullName;
    }

    private IMailFolder? ResolveSpecial(string canonical)
    {
        if (canonical.Equals("INBOX", StringComparison.OrdinalIgnoreCase)) return _client.Inbox;

        SpecialFolder? kind = canonical.ToLowerInvariant() switch
        {
            "sent items" => SpecialFolder.Sent,
            "drafts" => SpecialFolder.Drafts,
            "deleted items" => SpecialFolder.Trash,
            "junk email" => SpecialFolder.Junk,
            "archive" => SpecialFolder.Archive,
            _ => null
        };
        if (kind is null) return null;

        // Server ohne SPECIAL-USE/XLIST: kein Standardordner bekannt -> Aufrufer legt per Name an.
        try { return _client.GetFolder(kind.Value); }
        catch { return null; }
    }

    public async Task<bool> ExistsAsync(string folderHandle, string messageId, CancellationToken ct)
    {
        if (string.IsNullOrEmpty(messageId)) return false;

        var folder = await _client.GetFolderAsync(folderHandle, ct).ConfigureAwait(false);
        await folder.OpenAsync(FolderAccess.ReadOnly, ct).ConfigureAwait(false);

        var query = MailKit.Search.SearchQuery.HeaderContains("Message-Id", messageId);
        var hits = await folder.SearchAsync(query, ct).ConfigureAwait(false);
        return hits.Count > 0;
    }

    public async Task AppendAsync(string folderHandle, SourceMessage message, CancellationToken ct)
    {
        var folder = await _client.GetFolderAsync(folderHandle, ct).ConfigureAwait(false);

        var flags = MessageFlags.None;
        if (message.IsSeen) flags |= MessageFlags.Seen;
        if (message.IsFlagged) flags |= MessageFlags.Flagged;
        if (message.IsAnswered) flags |= MessageFlags.Answered;
        if (message.IsDraft) flags |= MessageFlags.Draft;

        // Exchange Online weist manche APPENDs mit Flags+Datum ab ("Command Argument Error").
        // Deshalb gestuft: erst mit Flags + INTERNALDATE, dann nur Flags, zuletzt nackt.
        try
        {
            await folder.AppendAsync(message.Mime, flags, message.InternalDate, ct).ConfigureAwait(false);
        }
        catch (MailKit.Net.Imap.ImapCommandException)
        {
            try
            {
                await folder.AppendAsync(message.Mime, flags, ct).ConfigureAwait(false);
            }
            catch (MailKit.Net.Imap.ImapCommandException)
            {
                await folder.AppendAsync(message.Mime, MessageFlags.None, ct).ConfigureAwait(false);
            }
        }
    }

    public async IAsyncEnumerable<TargetMessage> ListAsync(
        string folderHandle, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct)
    {
        var folder = await _client.GetFolderAsync(folderHandle, ct).ConfigureAwait(false);
        await folder.OpenAsync(FolderAccess.ReadOnly, ct).ConfigureAwait(false);

        var uids = await folder.SearchAsync(MailKit.Search.SearchQuery.All, ct).ConfigureAwait(false);
        if (uids.Count == 0) yield break;

        var summaries = await folder.FetchAsync(uids, MessageSummaryItems.UniqueId | MessageSummaryItems.Envelope, ct)
            .ConfigureAwait(false);
        foreach (var s in summaries)
            yield return new TargetMessage(s.Envelope?.MessageId ?? "", s.UniqueId.Id.ToString());
    }

    public async Task DeleteAsync(string folderHandle, string deleteHandle, CancellationToken ct)
    {
        var folder = await _client.GetFolderAsync(folderHandle, ct).ConfigureAwait(false);
        if (!folder.IsOpen || folder.Access != FolderAccess.ReadWrite)
            await folder.OpenAsync(FolderAccess.ReadWrite, ct).ConfigureAwait(false);

        var uid = new UniqueId(uint.Parse(deleteHandle));

        // Wenn möglich in den Papierkorb verschieben (wiederherstellbar), sonst endgültig löschen.
        IMailFolder? trash = null;
        try { trash = _client.GetFolder(SpecialFolder.Trash); } catch { /* kein SPECIAL-USE */ }

        if (trash is not null)
        {
            await folder.MoveToAsync(uid, trash, ct).ConfigureAwait(false);
        }
        else
        {
            await folder.AddFlagsAsync(uid, MessageFlags.Deleted, true, ct).ConfigureAwait(false);
            await folder.ExpungeAsync(ct).ConfigureAwait(false);
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (_client.IsConnected)
            await _client.DisconnectAsync(true).ConfigureAwait(false);
        _client.Dispose();
    }
}
