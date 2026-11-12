using System.Runtime.CompilerServices;
using IMAP2M365.Core;
using IMAP2M365.Ews;

namespace IMAP2M365.Targets;

/// <summary>
/// Ziel = Exchange on-prem über EWS (ApplicationImpersonation). Nachrichten werden als MIME
/// angelegt; Gelesen-Status, Kennzeichnung und Empfangsdatum bleiben über MAPI-Eigenschaften erhalten.
/// </summary>
public sealed class EwsTarget : IMailTarget
{
    private readonly EwsClient _ews;
    private readonly Dictionary<string, string> _folderCache = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, string> _wellKnown = new();

    public EwsTarget(string ewsUrl, string adminUser, string adminPassword, bool acceptAnyCertificate)
    {
        _ews = new EwsClient(ewsUrl, adminUser, adminPassword, acceptAnyCertificate);
    }

    public async Task ConnectAsync(string targetMailbox, CancellationToken ct)
    {
        _ews.ImpersonatedSmtp = targetMailbox;
        _wellKnown = await _ews.GetWellKnownFolderIdsAsync(ct).ConfigureAwait(false);
    }

    public async Task<string> EnsureFolderAsync(IReadOnlyList<string> pathSegments, CancellationToken ct)
    {
        var key = string.Join("\u0001", pathSegments);
        if (_folderCache.TryGetValue(key, out var cached)) return cached;

        // Kanonischer erster Abschnitt -> vorhandener Standardordner, sonst unter der Postfachwurzel.
        var parent = _wellKnown["msgfolderroot"];
        var start = 0;
        if (pathSegments.Count > 0
            && EwsClient.CanonicalToDistinguished.TryGetValue(pathSegments[0], out var distinguished)
            && _wellKnown.TryGetValue(distinguished, out var wellKnownId))
        {
            parent = wellKnownId;
            start = 1;
        }

        for (var i = start; i < pathSegments.Count; i++)
        {
            parent = await _ews.FindChildFolderAsync(parent, pathSegments[i], ct).ConfigureAwait(false)
                     ?? await _ews.CreateFolderAsync(parent, pathSegments[i], ct).ConfigureAwait(false);
        }

        _folderCache[key] = parent;
        return parent;
    }

    public Task<bool> ExistsAsync(string folderHandle, string messageId, CancellationToken ct) =>
        _ews.ExistsByMessageIdAsync(folderHandle, messageId, ct);

    public async Task AppendAsync(string folderHandle, SourceMessage message, CancellationToken ct)
    {
        using var ms = new MemoryStream();
        await message.Mime.WriteToAsync(ms, ct).ConfigureAwait(false);

        var sent = message.Mime.Date == default ? message.InternalDate : message.Mime.Date;
        await _ews.CreateMimeItemAsync(
            folderHandle, ms.ToArray(), message.IsSeen, message.IsDraft, message.IsFlagged,
            message.InternalDate, sent, ct).ConfigureAwait(false);
    }

    public async IAsyncEnumerable<TargetMessage> ListAsync(
        string folderHandle, [EnumeratorCancellation] CancellationToken ct)
    {
        await foreach (var item in _ews.FindItemsAsync(folderHandle, ct).ConfigureAwait(false))
            yield return new TargetMessage(item.InternetMessageId, item.Id);
    }

    public Task DeleteAsync(string folderHandle, string deleteHandle, CancellationToken ct) =>
        _ews.DeleteItemAsync(deleteHandle, ct);

    public ValueTask DisposeAsync()
    {
        _ews.Dispose();
        return ValueTask.CompletedTask;
    }
}
