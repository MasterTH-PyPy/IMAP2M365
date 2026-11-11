namespace IMAP2M365.Licensing;

/// <summary>
/// Hochsprachige Lizenzlogik auf diesem Rechner: Lizenz laden/aktivieren, Restkontingent und
/// Verbrauch pro Postfach. Dateien liegen im übergebenen Verzeichnis (z. B. ProgramData\IMAP2M365).
/// </summary>
public sealed class LicenseManager
{
    private readonly string _dir;
    private readonly string _machineId;

    private License? _license;
    private LicenseState? _state;

    public LicenseManager(string storageDir, string machineId)
    {
        _dir = storageDir;
        _machineId = machineId;
        Directory.CreateDirectory(_dir);
        Load();
    }

    private string LicensePath => Path.Combine(_dir, "license.key");
    private string StatePath => Path.Combine(_dir, "license-state.dat");

    public bool HasValidLicense => _license is not null;
    public string? Company => _license?.Payload.Company;
    public int MaxAccounts => _license?.Payload.MaxAccounts ?? 0;
    public int ConsumedCount => _state?.ConsumedCount ?? 0;
    public int Remaining => Math.Max(0, MaxAccounts - ConsumedCount);

    private void Load()
    {
        _license = null;
        _state = null;

        if (!File.Exists(LicensePath)) return;
        var lic = LicenseCrypto.VerifyLicense(File.ReadAllText(LicensePath).Trim(), EmbeddedPublicKey.Pem);
        if (lic is null) return;
        if (!string.Equals(lic.Payload.MachineId, _machineId, StringComparison.Ordinal)) return; // andere Maschine
        _license = lic;

        // Zustand laden oder neu anlegen
        if (File.Exists(StatePath))
            _state = LicenseState.Decrypt(File.ReadAllBytes(StatePath), _machineId, lic.Payload.LicenseId);
        _state ??= new LicenseState { LicenseId = lic.Payload.LicenseId, MachineId = _machineId };
    }

    /// <summary>Aktiviert einen Lizenzschlüssel (Signatur + Bindung an diese Maschine prüfen).</summary>
    public (bool ok, string message) Activate(string licenseKey)
    {
        var lic = LicenseCrypto.VerifyLicense(licenseKey.Trim(), EmbeddedPublicKey.Pem);
        if (lic is null)
            return (false, "Ungültiger Lizenzschlüssel (Signatur fehlgeschlagen).");
        if (!string.Equals(lic.Payload.MachineId, _machineId, StringComparison.Ordinal))
            return (false, "Diese Lizenz gehört zu einem anderen Rechner (Computer-ID passt nicht).");

        File.WriteAllText(LicensePath, licenseKey.Trim());
        _license = lic;

        // frischen Zustand behalten, falls vorhandener gültig ist, sonst neu
        if (File.Exists(StatePath))
            _state = LicenseState.Decrypt(File.ReadAllBytes(StatePath), _machineId, lic.Payload.LicenseId);
        _state ??= new LicenseState { LicenseId = lic.Payload.LicenseId, MachineId = _machineId };
        Persist();

        return (true, $"Lizenz aktiviert: {lic.Payload.MaxAccounts} Postfächer für {lic.Payload.Company}.");
    }

    public bool IsConsumed(string mailbox) =>
        _state?.ConsumedMailboxes.Contains(Norm(mailbox)) == true;

    /// <summary>
    /// Versucht, für ein Postfach eine Lizenz zu ziehen. Bereits verbrauchte Postfächer sind frei
    /// (idempotent). Gibt false zurück, wenn kein Kontingent mehr frei ist.
    /// </summary>
    public bool TryConsume(string mailbox)
    {
        if (_license is null || _state is null) return false;
        var m = Norm(mailbox);
        if (_state.ConsumedMailboxes.Contains(m)) return true;     // schon gezählt
        if (Remaining <= 0) return false;                          // Kontingent leer
        _state.ConsumedMailboxes.Add(m);
        Persist();
        return true;
    }

    private void Persist()
    {
        if (_state is null) return;
        File.WriteAllBytes(StatePath, _state.Encrypt());
    }

    private static string Norm(string mailbox) => mailbox.Trim().ToLowerInvariant();
}
