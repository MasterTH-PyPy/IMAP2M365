using MimeKit;

namespace IMAP2M365.Core;

/// <summary>Eine zu übertragende Nachricht samt der Metadaten, die erhalten bleiben sollen.</summary>
public sealed record SourceMessage(
    MimeMessage Mime,
    DateTimeOffset InternalDate, // Empfangsdatum -> APPEND INTERNALDATE
    bool IsSeen,
    bool IsFlagged,
    bool IsAnswered,
    bool IsDraft)
{
    /// <summary>Message-Id-Header, normalisiert; für Dedup beim Resume.</summary>
    public string MessageId => Mime.MessageId ?? string.Empty;
}

public sealed record FolderInfo(string FullName, IReadOnlyList<string> PathSegments);

/// <summary>Quelle: beliebiger IMAP-Server (Basic- oder XOAUTH2-Login).</summary>
public interface IMailSource : IAsyncDisposable
{
    Task ConnectAsync(CancellationToken ct);
    Task<IReadOnlyList<FolderInfo>> ListFoldersAsync(CancellationToken ct);

    /// <summary>Liefert Nachrichten eines Ordners gestreamt (schont Speicher bei großen Postfächern).</summary>
    IAsyncEnumerable<SourceMessage> ReadFolderAsync(FolderInfo folder, CancellationToken ct);
}

/// <summary>Ziel-Backend. IMAP-APPEND oder Graph – austauschbar.</summary>
public interface IMailTarget : IAsyncDisposable
{
    Task ConnectAsync(string targetMailbox, CancellationToken ct);

    /// <summary>Stellt sicher, dass der (gemappte) Zielordner existiert, und gibt einen Handle zurück.</summary>
    Task<string> EnsureFolderAsync(IReadOnlyList<string> pathSegments, CancellationToken ct);

    /// <summary>Existiert eine Nachricht mit dieser Message-Id im Zielordner schon? (Dedup bei Resume.)</summary>
    Task<bool> ExistsAsync(string folderHandle, string messageId, CancellationToken ct);

    Task AppendAsync(string folderHandle, SourceMessage message, CancellationToken ct);
}

/// <summary>Liefert app-only Access-Token für den angefragten Resource-Scope.</summary>
public interface ITokenProvider
{
    /// <param name="scope">z. B. https://outlook.office365.com/.default oder https://graph.microsoft.com/.default</param>
    Task<string> GetTokenAsync(string scope, CancellationToken ct);
}

/// <summary>Konfiguration eines Migrationslaufs (aus AES-256-config.enc).</summary>
public sealed class MigrationConfig
{
    // Ziel-Tenant (app-only)
    public string TenantId { get; init; } = "";
    public string ClientId { get; init; } = "";
    public string ClientSecret { get; init; } = "";

    // Quelle (Default; pro Postfach in der CSV überschreibbar)
    public string SourceHost { get; init; } = "";
    public int SourcePort { get; init; } = 993;
    public bool SourceUseSsl { get; init; } = true;

    // Verhalten
    public bool DryRun { get; init; } = true;
    /// <summary>Testmodus: nach so vielen kopierten Nachrichten pro Postfach abbrechen (0 = unbegrenzt).</summary>
    public int TestMessageLimit { get; init; } = 0;
    public IReadOnlyList<string> ExcludedFolders { get; init; } = new[] { "Junk", "Trash", "Deleted Items" };
    public TargetKind Target { get; init; } = TargetKind.ImapAppend;
}

public enum TargetKind { ImapAppend, Graph }

/// <summary>Ein Postfach-Mapping aus der Batch-CSV: Quelle -> Ziel-UPN.</summary>
public sealed record MailboxMapping(
    string SourceUser,
    string SourcePassword,   // oder leer bei XOAUTH2-Quelle
    string TargetUpn);
