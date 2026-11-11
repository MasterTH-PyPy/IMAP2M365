using System.Text.Json.Serialization;

namespace IMAP2M365.Gui;

/// <summary>Persistierte Konfiguration (landet verschlüsselt in config.enc).</summary>
public sealed class AppConfig
{
    // Ziel-Tenant / app-only
    public string TenantId { get; set; } = "";
    public string ClientId { get; set; } = "";
    public string ClientSecret { get; set; } = "";
    public string TargetDomain { get; set; } = ""; // z. B. contoso.onmicrosoft.com

    // Zentrale Quell-IMAP-Defaults
    public string SourceHost { get; set; } = "";
    public int SourcePort { get; set; } = 993;
    public bool SourceUseSsl { get; set; } = true;

    // Verhalten
    public bool DryRun { get; set; } = true;
    public List<string> ExcludedFolders { get; set; } = new() { "Junk", "Junk-E-Mail", "Trash", "Deleted Items", "Gelöschte Elemente" };
    public string TargetKind { get; set; } = "ImapAppend"; // ImapAppend | Graph
    public string? ScopeGroupId { get; set; } // optionale ApplicationAccessPolicy-Gruppe

    // Postfachzeilen inkl. Quell-Credentials (maskiert in der GUI, verschlüsselt auf Platte)
    public List<MailboxRow> Mailboxes { get; set; } = new();
}

/// <summary>Eine Zeile im Setup-Grid: Ziel aus EXO ausgelesen, Quelle vom Benutzer ergänzt.</summary>
public sealed class MailboxRow
{
    // Ziel (aus EXO)
    public string TargetSmtp { get; set; } = "";
    public string MailboxType { get; set; } = "UserMailbox"; // UserMailbox | SharedMailbox
    public bool ImapEnabled { get; set; }                     // aus Get-CASMailbox
    public bool PermissionGranted { get; set; }               // Add-MailboxPermission erledigt?

    // Quelle (vom Benutzer; Server/Port leer = zentraler Default)
    public string SourceUser { get; set; } = "";
    public string SourcePassword { get; set; } = "";
    public string? SourceHostOverride { get; set; }
    public int? SourcePortOverride { get; set; }

    [JsonIgnore] public bool IsShared => MailboxType.Equals("SharedMailbox", StringComparison.OrdinalIgnoreCase);
}
