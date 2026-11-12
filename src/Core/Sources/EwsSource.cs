using System.Runtime.CompilerServices;
using IMAP2M365.Core;
using IMAP2M365.Ews;
using MimeKit;

namespace IMAP2M365.Sources;

/// <summary>
/// Quelle = Exchange on-prem über EWS. Ein Admin-/Dienstkonto mit der Rolle
/// ApplicationImpersonation liest das jeweilige Postfach; IMAP muss dafür nicht aktiv sein.
/// Standardordner werden auf die kanonischen Namen der Engine abgebildet (INBOX, Sent Items …),
/// damit das Ziel sie unabhängig von der Postfachsprache richtig zuordnet.
/// </summary>
public sealed class EwsSource : IMailSource
{
    private readonly EwsClient _ews;
    private readonly string _mailbox;
    private readonly Dictionary<string, int> _counts = new(StringComparer.Ordinal);
    private Dictionary<string, string> _wellKnown = new();

    public EwsSource(string ewsUrl, string adminUser, string adminPassword, bool acceptAnyCertificate, string mailboxSmtp)
    {
        _ews = new EwsClient(ewsUrl, adminUser, adminPassword, acceptAnyCertificate);
        _mailbox = mailboxSmtp;
    }

    public async Task ConnectAsync(CancellationToken ct)
    {
        _ews.ImpersonatedSmtp = _mailbox;
        _wellKnown = await _ews.GetWellKnownFolderIdsAsync(ct).ConfigureAwait(false);
    }

    public async Task<IReadOnlyList<FolderInfo>> ListFoldersAsync(CancellationToken ct)
    {
        var all = await _ews.FindFoldersDeepAsync(ct).ConfigureAwait(false);
        var byId = all.ToDictionary(f => f.Id, StringComparer.Ordinal);
        var rootId = _wellKnown["msgfolderroot"];

        // FolderId -> kanonischer Name für die Standardordner
        var canonical = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (name, distinguished) in EwsClient.CanonicalToDistinguished)
            if (_wellKnown.TryGetValue(distinguished, out var id)) canonical[id] = name;

        var result = new List<FolderInfo>();
        foreach (var f in all)
        {
            // Nur Mailordner (Klasse leer oder IPF.Note*); Kalender/Kontakte/Aufgaben bleiben außen vor.
            if (f.FolderClass.Length > 0 && !f.FolderClass.StartsWith("IPF.Note", StringComparison.OrdinalIgnoreCase))
                continue;

            var segments = new List<string>();
            var current = f;
            var guard = 0;
            while (guard++ < 64)
            {
                segments.Insert(0, canonical.TryGetValue(current.Id, out var c) ? c : current.DisplayName);
                if (current.ParentId == rootId || !byId.TryGetValue(current.ParentId, out var parent)) break;
                current = parent;
            }

            _counts[f.Id] = f.TotalCount;
            result.Add(new FolderInfo(string.Join("/", segments), segments, f.Id));
        }
        return result;
    }

    public Task<int> GetMessageCountAsync(FolderInfo folder, CancellationToken ct) =>
        Task.FromResult(folder.Handle is not null && _counts.TryGetValue(folder.Handle, out var n) ? n : 0);

    public async IAsyncEnumerable<SourceMessage> ReadFolderAsync(
        FolderInfo folder, [EnumeratorCancellation] CancellationToken ct)
    {
        if (folder.Handle is null) yield break;

        await foreach (var info in _ews.FindItemsAsync(folder.Handle, ct).ConfigureAwait(false))
        {
            ct.ThrowIfCancellationRequested();

            var bytes = await TryGetMimeAsync(info.Id, ct).ConfigureAwait(false);
            if (bytes is null) continue; // Element ohne MIME-Darstellung (selten) -> überspringen

            MimeMessage mime;
            using (var ms = new MemoryStream(bytes))
                mime = await MimeMessage.LoadAsync(ms, ct).ConfigureAwait(false);

            yield return new SourceMessage(
                Mime: mime,
                InternalDate: info.Received ?? mime.Date,
                IsSeen: info.IsRead,
                IsFlagged: info.FlagStatus == 2,
                IsAnswered: info.LastVerb is 102 or 103,
                IsDraft: info.IsDraft,
                Size: info.Size > 0 ? info.Size : bytes.Length);
        }
    }

    private async Task<byte[]?> TryGetMimeAsync(string itemId, CancellationToken ct)
    {
        try { return await _ews.GetMimeAsync(itemId, ct).ConfigureAwait(false); }
        catch (EwsException) { return null; }
    }

    public ValueTask DisposeAsync()
    {
        _ews.Dispose();
        return ValueTask.CompletedTask;
    }
}
