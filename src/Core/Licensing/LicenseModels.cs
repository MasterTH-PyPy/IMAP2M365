namespace IMAP2M365.Licensing;

/// <summary>
/// Aktivierungsanfrage, die der Kunde erzeugt und an den Lizenzgeber schickt.
/// Wird mit dem öffentlichen Schlüssel hybrid verschlüsselt -> nur der Lizenzgeber liest sie.
/// </summary>
public sealed class ActivationRequest
{
    public string MachineId { get; set; } = "";       // Computer-ID (Hash, node-locked)
    public string MachineName { get; set; } = "";      // Rechnername (Klartext, zur Zuordnung)
    public bool IsVirtualMachine { get; set; }         // VM erkannt?
    public string VmHint { get; set; } = "";           // woran VM erkannt wurde (Hersteller/Modell)

    // Vom Kunden eingegeben
    public string Company { get; set; } = "";
    public string Contact { get; set; } = "";
    public string Address { get; set; } = "";
    public string Email { get; set; } = "";

    public string ToolVersion { get; set; } = "";
    public DateTimeOffset CreatedUtc { get; set; } = DateTimeOffset.UtcNow;
}

/// <summary>
/// Vom Lizenzgeber signierte Lizenz. An die Computer-ID gebunden, unbefristet,
/// mit maximaler Kontenzahl. Der signierte Teil ist <see cref="Payload"/>.
/// </summary>
public sealed class License
{
    public LicensePayload Payload { get; set; } = new();
    public string Signature { get; set; } = "";       // Base64, RSA-PSS SHA-256 über die Payload-JSON
}

public sealed class LicensePayload
{
    public string LicenseId { get; set; } = Guid.NewGuid().ToString("N");
    public string MachineId { get; set; } = "";        // muss zur Computer-ID des Rechners passen
    public int MaxAccounts { get; set; }               // wie viele Postfächer migriert werden dürfen

    // Kundenbezug (zur Anzeige/Nachweis)
    public string Company { get; set; } = "";
    public string Contact { get; set; } = "";
    public string Email { get; set; } = "";

    public DateTimeOffset IssuedUtc { get; set; } = DateTimeOffset.UtcNow;
}
