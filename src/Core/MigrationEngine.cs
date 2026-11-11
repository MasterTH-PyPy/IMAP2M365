namespace IMAP2M365.Core;

public sealed record MigrationProgress(
    int MailboxIndex, int MailboxTotal, string Mailbox,
    int FolderIndex, int FolderTotal, string Folder,
    int MessageIndex, int MessageTotal,
    long MessagesCopied, long MessagesSkipped);

/// <summary>
/// Fährt einen Quell-IMAP gegen ein Ziel-Backend: Ordner mappen, Nachrichten streamen,
/// im Dry-Run nur zählen, sonst APPEND mit Dedup (Message-Id) für saubere Resumes.
/// Quelle/Ziel kommen fertig verbunden rein (Postfach-UPN steht beim Target schon).
/// </summary>
public sealed class MigrationEngine
{
    private readonly MigrationConfig _cfg;
    private readonly Func<IReadOnlyList<string>, IReadOnlyList<string>> _mapFolder;
    private readonly IProgress<MigrationProgress>? _progress;
    private readonly Action<string>? _log;

    public MigrationEngine(
        MigrationConfig cfg,
        Func<IReadOnlyList<string>, IReadOnlyList<string>>? mapFolder = null,
        IProgress<MigrationProgress>? progress = null,
        Action<string>? log = null)
    {
        _cfg = cfg;
        _mapFolder = mapFolder ?? (p => p); // 1:1, bis die GUI ein Mapping liefert
        _progress = progress;
        _log = log;
    }

    public async Task<(long copied, long skipped)> RunAsync(
        string mailbox, IMailSource source, IMailTarget target, CancellationToken ct,
        int mailboxIndex = 1, int mailboxTotal = 1)
    {
        long copied = 0, skipped = 0;

        var folders = await source.ListFoldersAsync(ct).ConfigureAwait(false);
        var selectable = folders
            .Where(f => !_cfg.ExcludedFolders.Any(x => f.PathSegments.Any(s => s.Equals(x, StringComparison.OrdinalIgnoreCase))))
            .ToList();

        var testMode = _cfg.TestMessageLimit > 0;

        // Für den Spiegelabgleich: je Zielordner-Handle die aktuell in der Quelle erwarteten Message-Ids.
        var expectedByHandle = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);

        for (int fi = 0; fi < selectable.Count; fi++)
        {
            ct.ThrowIfCancellationRequested();
            var folder = selectable[fi];
            var targetPath = _mapFolder(folder.PathSegments);

            var handle = await target.EnsureFolderAsync(targetPath, ct).ConfigureAwait(false);
            if (_cfg.Mirror && !expectedByHandle.ContainsKey(handle))
                expectedByHandle[handle] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            _log?.Invoke($"[{mailbox}] {folder.FullName} -> {string.Join('/', targetPath)}{(testMode ? " (Testmodus)" : _cfg.Mirror ? " (Spiegel)" : "")}");

            var msgTotal = await source.GetMessageCountAsync(folder, ct).ConfigureAwait(false);
            var msgIdx = 0;

            await foreach (var msg in source.ReadFolderAsync(folder, ct).ConfigureAwait(false))
            {
                ct.ThrowIfCancellationRequested();
                msgIdx++;

                if (_cfg.Mirror) expectedByHandle[handle].Add(Normalize(msg.MessageId));

                _progress?.Report(new MigrationProgress(
                    mailboxIndex, mailboxTotal, mailbox, fi + 1, selectable.Count, folder.FullName,
                    msgIdx, msgTotal, copied, skipped));

                if (await target.ExistsAsync(handle, msg.MessageId, ct).ConfigureAwait(false))
                {
                    skipped++;
                    continue;
                }

                if (_cfg.Debug) _log?.Invoke($"  → kopiere: {Trim(msg.Mime.Subject)}");
                await target.AppendAsync(handle, msg, ct).ConfigureAwait(false);
                copied++;

                if (testMode && copied >= _cfg.TestMessageLimit)
                {
                    _log?.Invoke($"[{mailbox}] Testmodus: {copied} Nachrichten kopiert – Abbruch wie vorgesehen.");
                    return (copied, skipped);
                }
            }
        }

        long deleted = 0;
        if (_cfg.Mirror)
            deleted = await ReconcileAsync(mailbox, target, expectedByHandle, ct).ConfigureAwait(false);

        _log?.Invoke($"[{mailbox}] fertig: {copied} kopiert, {skipped} übersprungen" +
                     (_cfg.Mirror ? $", {deleted} im Ziel entfernt (Spiegel)" : testMode ? " (Testmodus)" : ""));
        return (copied, skipped);
    }

    /// <summary>
    /// Spiegelabgleich: In jedem gespiegelten Zielordner die Nachrichten löschen (in den Papierkorb),
    /// deren Message-Id in der Quelle (an dieser Stelle) nicht mehr vorkommt – also verschobene oder
    /// am Quellserver gelöschte Mails. Nachrichten ohne Message-Id bleiben sicherheitshalber erhalten.
    /// </summary>
    private async Task<long> ReconcileAsync(
        string mailbox, IMailTarget target, Dictionary<string, HashSet<string>> expectedByHandle, CancellationToken ct)
    {
        long deleted = 0;
        foreach (var (handle, expected) in expectedByHandle)
        {
            ct.ThrowIfCancellationRequested();

            // Erst komplett auflisten, dann löschen (vermeidet Zustandskonflikte beim Blättern/Umschalten).
            var toDelete = new List<string>();
            await foreach (var tm in target.ListAsync(handle, ct).ConfigureAwait(false))
            {
                var id = Normalize(tm.MessageId);
                if (id.Length == 0) continue;                 // ohne Id nicht abgleichbar -> behalten
                if (!expected.Contains(id)) toDelete.Add(tm.DeleteHandle);
            }

            foreach (var del in toDelete)
            {
                ct.ThrowIfCancellationRequested();
                await target.DeleteAsync(handle, del, ct).ConfigureAwait(false);
                deleted++;
            }
            if (toDelete.Count > 0)
                _log?.Invoke($"[{mailbox}] Spiegel: {toDelete.Count} überzählige Nachrichten entfernt.");
        }
        return deleted;
    }

    private static string Normalize(string messageId) =>
        messageId.Trim().Trim('<', '>').ToLowerInvariant();

    private static string Trim(string? subject) =>
        string.IsNullOrEmpty(subject) ? "(kein Betreff)" : (subject.Length > 80 ? subject[..80] + "…" : subject);
    }
}
