using System.Text.Json.Serialization;

namespace IMAP2M365.Gui;

/// <summary>Persistierte Konfiguration (landet verschlüsselt in config.enc).</summary>
public sealed class AppConfig
{
    // Ziel-Tenant / app-only
    public string TenantId { get; set; } = "";
    public string ClientId { get; set; } = "";
    public string ClientSecret { get; set; } = "";
    public string SpObjectId { get; set; } = ""; // ObjectId des ServicePrincipals (aus Provisionierung)
    public string TargetDomain { get; set; } = ""; // z. B. contoso.onmicrosoft.com

    // Zentrale Quell-IMAP-Defaults
    public string SourceHost { get; set; } = "";
    public int SourcePort { get; set; } = 993;
    public bool SourceUseSsl { get; set; } = true;

    // Quelle: "Imap" (beliebiger IMAP-Server) oder "OnPremEws" (Exchange on-prem per EWS + Impersonation)
    public string SourceKind { get; set; } = "Imap";
    public string OnPremEwsUrl { get; set; } = "";      // https://server/EWS/Exchange.asmx
    public string OnPremPsUri { get; set; } = "";       // http://server/PowerShell/
    public string OnPremPsAuth { get; set; } = "Kerberos";
    public string OnPremAdminUser { get; set; } = "";   // Migrationskonto (DOMAIN\user oder UPN)
    public string OnPremAdminPassword { get; set; } = "";
    public bool OnPremAcceptAnyCert { get; set; }

    [JsonIgnore] public bool SourceIsOnPrem => SourceKind.Equals("OnPremEws", StringComparison.OrdinalIgnoreCase);

    // Verhalten
    public bool DryRun { get; set; } = true;
    public List<string> ExcludedFolders { get; set; } = new() { "Junk", "Junk-E-Mail", "Trash", "Deleted Items", "Gelöschte Elemente" };
    public string TargetKind { get; set; } = "ImapAppend"; // ImapAppend | Graph
    public string? ScopeGroupId { get; set; } // optionale ApplicationAccessPolicy-Gruppe

    // Profil/Kunde (zur Anzeige; Dateiname bestimmt den tatsächlichen Speicherort)
    public string ProfileName { get; set; } = "";

    // Kumulierte Statistik (über alle Läufe dieses Profils)
    public long StatMessagesCopied { get; set; }
    public long StatBytesDown { get; set; }
    public long StatBytesUp { get; set; }

    // Postfachzeilen inkl. Quell-Credentials (maskiert in der GUI, verschlüsselt auf Platte)
    public List<MailboxRow> Mailboxes { get; set; } = new();

    /// <summary>Übernimmt alle Werte eines geladenen Profils in diese (von den Reitern gehaltene) Instanz.</summary>
    public void CopyFrom(AppConfig other)
    {
        TenantId = other.TenantId;
        ClientId = other.ClientId;
        ClientSecret = other.ClientSecret;
        SpObjectId = other.SpObjectId;
        TargetDomain = other.TargetDomain;
        SourceHost = other.SourceHost;
        SourcePort = other.SourcePort;
        SourceUseSsl = other.SourceUseSsl;
        SourceKind = other.SourceKind;
        OnPremEwsUrl = other.OnPremEwsUrl;
        OnPremPsUri = other.OnPremPsUri;
        OnPremPsAuth = other.OnPremPsAuth;
        OnPremAdminUser = other.OnPremAdminUser;
        OnPremAdminPassword = other.OnPremAdminPassword;
        OnPremAcceptAnyCert = other.OnPremAcceptAnyCert;
        DryRun = other.DryRun;
        ExcludedFolders = other.ExcludedFolders;
        TargetKind = other.TargetKind;
        ScopeGroupId = other.ScopeGroupId;
        ProfileName = other.ProfileName;
        StatMessagesCopied = other.StatMessagesCopied;
        StatBytesDown = other.StatBytesDown;
        StatBytesUp = other.StatBytesUp;
        Mailboxes = other.Mailboxes;
    }
}

/// <summary>Eine Zeile im Setup-Grid: Ziel aus EXO ausgelesen, Quelle vom Benutzer ergänzt.</summary>
public sealed class MailboxRow
{
    // Aktiv? Deaktivierte Konten werden bei Einrichtung und Migration übersprungen.
    // Default true -> alte Profile ohne dieses Feld bleiben aktiv (STJ überschreibt fehlende Felder nicht).
    public bool Enabled { get; set; } = true;

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
