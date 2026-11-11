using System.Runtime.CompilerServices;
using IMAP2M365.Core;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;

namespace IMAP2M365.Sources;

/// <summary>
/// Quelle = beliebiger IMAP-Server. Login per Basic-Auth (Alt-Server, z. B. i-MSCP)
/// oder optional XOAUTH2, falls die Quelle selbst nur noch OAuth erlaubt.
/// Nachrichten werden gestreamt gelesen, damit auch große Postfächer nicht den RAM sprengen.
/// </summary>
public sealed class ImapSource : IMailSource
{
    private readonly ImapClient _client = new();
    private readonly string _host;
    private readonly int _port;
    private readonly bool _useSsl;
    private readonly string _user;
    private readonly string _password; // leer => XOAUTH2
    private readonly ITokenProvider? _oauth;
    private readonly string? _oauthScope;

    public ImapSource(string host, int port, bool useSsl, string user, string password)
    {
        _host = host; _port = port; _useSsl = useSsl; _user = user; _password = password;
    }

    public ImapSource(string host, int port, bool useSsl, string user, ITokenProvider oauth, string scope)
    {
        _host = host; _port = port; _useSsl = useSsl; _user = user; _password = "";
        _oauth = oauth; _oauthScope = scope;
    }

    public async Task ConnectAsync(CancellationToken ct)
    {
        var opt = _useSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTlsWhenAvailable;
        await _client.ConnectAsync(_host, _port, opt, ct).ConfigureAwait(false);

        if (_oauth is not null)
        {
            var token = await _oauth.GetTokenAsync(_oauthScope!, ct).ConfigureAwait(false);
            await _client.AuthenticateAsync(new SaslMechanismOAuth2(_user, token), ct).ConfigureAwait(false);
        }
        else
        {
            await _client.AuthenticateAsync(_user, _password, ct).ConfigureAwait(false);
        }
    }

    public async Task<IReadOnlyList<FolderInfo>> ListFoldersAsync(CancellationToken ct)
    {
        var personal = _client.GetFolder(_client.PersonalNamespaces[0]);
        var all = await personal.GetSubfoldersAsync(subscribedOnly: false, ct).ConfigureAwait(false);

        var result = new List<FolderInfo> { ToInfo(_client.Inbox) };
        foreach (var f in all)
            await CollectAsync(f, result, ct).ConfigureAwait(false);
        return result;
    }

    private async Task CollectAsync(IMailFolder folder, List<FolderInfo> acc, CancellationToken ct)
    {
        if ((folder.Attributes & FolderAttributes.NonExistent) == 0 &&
            (folder.Attributes & FolderAttributes.NoSelect) == 0)
            acc.Add(ToInfo(folder));

        var subs = await folder.GetSubfoldersAsync(false, ct).ConfigureAwait(false);
        foreach (var s in subs)
            await CollectAsync(s, acc, ct).ConfigureAwait(false);
    }

    private static FolderInfo ToInfo(IMailFolder f)
    {
        var segments = f.FullName.Split(f.DirectorySeparator);
        return new FolderInfo(f.FullName, segments);
    }

    public async IAsyncEnumerable<SourceMessage> ReadFolderAsync(
        FolderInfo folder, [EnumeratorCancellation] CancellationToken ct)
    {
        var mf = await _client.GetFolderAsync(folder.FullName, ct).ConfigureAwait(false);
        await mf.OpenAsync(FolderAccess.ReadOnly, ct).ConfigureAwait(false);

        for (int i = 0; i < mf.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var summary = (await mf.FetchAsync(i, i, MessageSummaryItems.Flags | MessageSummaryItems.InternalDate, ct)
                .ConfigureAwait(false))[0];
            var mime = await mf.GetMessageAsync(i, ct).ConfigureAwait(false);

            var flags = summary.Flags ?? MessageFlags.None;
            yield return new SourceMessage(
                Mime: mime,
                InternalDate: summary.InternalDate ?? mime.Date,
                IsSeen: flags.HasFlag(MessageFlags.Seen),
                IsFlagged: flags.HasFlag(MessageFlags.Flagged),
                IsAnswered: flags.HasFlag(MessageFlags.Answered),
                IsDraft: flags.HasFlag(MessageFlags.Draft));
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (_client.IsConnected)
            await _client.DisconnectAsync(true).ConfigureAwait(false);
        _client.Dispose();
    }
}
