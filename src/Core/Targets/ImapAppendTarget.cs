using IMAP2M365.Core;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;

namespace IMAP2M365.Targets;

/// <summary>
/// Ziel = Exchange Online via IMAP-APPEND mit app-only XOAUTH2.
/// Überträgt rohes MIME 1:1; Flags und INTERNALDATE (Empfangsdatum) bleiben erhalten.
/// </summary>
public sealed class ImapAppendTarget : IMailTarget
{
    private const string Host = "outlook.office365.com";
    private const int Port = 993;

    private readonly ITokenProvider _tokens;
    private readonly ImapClient _client = new();
    private string _mailbox = "";

    public ImapAppendTarget(ITokenProvider tokens) => _tokens = tokens;

    public async Task ConnectAsync(string targetMailbox, CancellationToken ct)
    {
        _mailbox = targetMailbox;
        await _client.ConnectAsync(Host, Port, SecureSocketOptions.SslOnConnect, ct)
            .ConfigureAwait(false);

        var token = await _tokens.GetTokenAsync(ExchangeOnlineTokenProvider.ExchangeImapScope, ct)
            .ConfigureAwait(false);

        // app-only: UPN des Zielpostfachs + App-Token. Erfordert vorab registrierten
        // ServicePrincipal + Mailbox-Permission in Exchange (siehe prep/-Skript).
        var sasl = new SaslMechanismOAuth2(targetMailbox, token);
        await _client.AuthenticateAsync(sasl, ct).ConfigureAwait(false);
    }

    public async Task<string> EnsureFolderAsync(IReadOnlyList<string> pathSegments, CancellationToken ct)
    {
        var personal = _client.GetFolder(_client.PersonalNamespaces[0]);
        IMailFolder current = _client.Inbox; // Startpunkt für Root-nahe Pfade

        // Sonderfall: "INBOX" direkt
        if (pathSegments.Count == 1 && pathSegments[0].Equals("INBOX", StringComparison.OrdinalIgnoreCase))
            return _client.Inbox.FullName;

        current = personal;
        foreach (var segment in pathSegments)
        {
            IMailFolder child;
            try
            {
                child = await current.GetSubfolderAsync(segment, ct).ConfigureAwait(false);
            }
            catch (FolderNotFoundException)
            {
                child = await current.CreateAsync(segment, isMessageFolder: true, ct).ConfigureAwait(false);
            }
            current = child;
        }
        return current.FullName;
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
                await folder.AppendAsync(message.Mime, ct).ConfigureAwait(false);
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
