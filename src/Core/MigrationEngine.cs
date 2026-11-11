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

        for (int fi = 0; fi < selectable.Count; fi++)
        {
            ct.ThrowIfCancellationRequested();
            var folder = selectable[fi];
            var targetPath = _mapFolder(folder.PathSegments);

            string handle = "";
            if (!_cfg.DryRun)
                handle = await target.EnsureFolderAsync(targetPath, ct).ConfigureAwait(false);

            _log?.Invoke($"[{mailbox}] {folder.FullName} -> {string.Join('/', targetPath)}{(_cfg.DryRun ? " (dry-run)" : "")}");

            await foreach (var msg in source.ReadFolderAsync(folder, ct).ConfigureAwait(false))
            {
                ct.ThrowIfCancellationRequested();

                if (_cfg.DryRun)
                {
                    copied++; // im Dry-Run zählt alles als "würde kopiert"
                    continue;
                }

                if (await target.ExistsAsync(handle, msg.MessageId, ct).ConfigureAwait(false))
                {
                    skipped++;
                    continue;
                }

                await target.AppendAsync(handle, msg, ct).ConfigureAwait(false);
                copied++;

                if ((copied & 0x3F) == 0) // alle 64 Nachrichten Fortschritt melden
                    _progress?.Report(new MigrationProgress(mailbox, folder.FullName, fi + 1, selectable.Count, copied, skipped));
            }

            _progress?.Report(new MigrationProgress(mailbox, folder.FullName, fi + 1, selectable.Count, copied, skipped));
        }

        _log?.Invoke($"[{mailbox}] fertig: {copied} kopiert, {skipped} übersprungen{(_cfg.DryRun ? " (dry-run)" : "")}");
        return (copied, skipped);
    }
}
