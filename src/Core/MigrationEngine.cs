namespace IMAP2M365.Core;

public sealed record MigrationProgress(
    string Mailbox, string Folder, int FolderDone, int FolderTotal, long MessagesCopied, long MessagesSkipped);

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
        string mailbox, IMailSource source, IMailTarget target, CancellationToken ct)
    {
        long copied = 0, skipped = 0;

        var folders = await source.ListFoldersAsync(ct).ConfigureAwait(false);
        var selectable = folders
            .Where(f => !_cfg.ExcludedFolders.Any(x => f.PathSegments.Any(s => s.Equals(x, StringComparison.OrdinalIgnoreCase))))
            .ToList();

        var testMode = _cfg.TestMessageLimit > 0;

        for (int fi = 0; fi < selectable.Count; fi++)
        {
            ct.ThrowIfCancellationRequested();
            var folder = selectable[fi];
            var targetPath = _mapFolder(folder.PathSegments);

            var handle = await target.EnsureFolderAsync(targetPath, ct).ConfigureAwait(false);

            _log?.Invoke($"[{mailbox}] {folder.FullName} -> {string.Join('/', targetPath)}{(testMode ? " (Testmodus)" : "")}");

            await foreach (var msg in source.ReadFolderAsync(folder, ct).ConfigureAwait(false))
            {
                ct.ThrowIfCancellationRequested();

                if (await target.ExistsAsync(handle, msg.MessageId, ct).ConfigureAwait(false))
                {
                    skipped++;
                    continue;
                }

                await target.AppendAsync(handle, msg, ct).ConfigureAwait(false);
                copied++;

                if ((copied & 0x3F) == 0)
                    _progress?.Report(new MigrationProgress(mailbox, folder.FullName, fi + 1, selectable.Count, copied, skipped));

                if (testMode && copied >= _cfg.TestMessageLimit)
                {
                    _log?.Invoke($"[{mailbox}] Testmodus: {copied} Nachrichten kopiert – Abbruch wie vorgesehen.");
                    _progress?.Report(new MigrationProgress(mailbox, folder.FullName, fi + 1, selectable.Count, copied, skipped));
                    return (copied, skipped);
                }
            }

            _progress?.Report(new MigrationProgress(mailbox, folder.FullName, fi + 1, selectable.Count, copied, skipped));
        }

        _log?.Invoke($"[{mailbox}] fertig: {copied} kopiert, {skipped} übersprungen{(testMode ? " (Testmodus)" : "")}");
        return (copied, skipped);
    }
}
