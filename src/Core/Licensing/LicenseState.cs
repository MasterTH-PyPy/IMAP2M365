using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace IMAP2M365.Licensing;

/// <summary>
/// Verbrauchszustand auf diesem Rechner: welche Postfächer bereits eine Lizenz gezogen haben.
/// An MachineId + LicenseId gebunden und AES-256-GCM-verschlüsselt (tamper-evident). Kopieren auf
/// einen anderen Rechner ist wirkungslos (anderer Schlüssel). Löschen = Reset bleibt ein Restrisiko.
/// </summary>
public sealed class LicenseState
{
    public string LicenseId { get; set; } = "";
    public string MachineId { get; set; } = "";
    /// <summary>Bereits verbrauchte Zielpostfächer (SMTP, klein geschrieben).</summary>
    public List<string> ConsumedMailboxes { get; set; } = new();

    public int ConsumedCount => ConsumedMailboxes.Count;

    private static readonly JsonSerializerOptions Json = new() { WriteIndented = false };

    private static byte[] DeriveKey(string machineId, string licenseId)
    {
        using var sha = SHA256.Create();
        return sha.ComputeHash(Encoding.UTF8.GetBytes($"IMAP2M365-state|{machineId}|{licenseId}"));
    }

    public byte[] Encrypt()
    {
        var plain = JsonSerializer.SerializeToUtf8Bytes(this, Json);
        var key = DeriveKey(MachineId, LicenseId);
        var nonce = RandomNumberGenerator.GetBytes(12);
        var cipher = new byte[plain.Length];
        var tag = new byte[16];
        using (var gcm = new AesGcm(key, 16))
            gcm.Encrypt(nonce, plain, cipher, tag);

        using var ms = new MemoryStream();
        ms.Write(nonce);
        ms.Write(tag);
        ms.Write(cipher);
        return ms.ToArray();
    }

    /// <summary>Entschlüsselt und prüft gegen erwartete MachineId/LicenseId. Bei Manipulation/Fehlbindung null.</summary>
    public static LicenseState? Decrypt(byte[] data, string machineId, string licenseId)
    {
        try
        {
            if (data.Length < 28) return null;
            var nonce = data.AsSpan(0, 12).ToArray();
            var tag = data.AsSpan(12, 16).ToArray();
            var cipher = data.AsSpan(28).ToArray();

            var key = DeriveKey(machineId, licenseId);
            var plain = new byte[cipher.Length];
            using (var gcm = new AesGcm(key, 16))
                gcm.Decrypt(nonce, cipher, tag, plain);

            var state = JsonSerializer.Deserialize<LicenseState>(plain, Json);
            if (state is null) return null;
            // Bindung hart prüfen
            if (!string.Equals(state.MachineId, machineId, StringComparison.Ordinal)) return null;
            if (!string.Equals(state.LicenseId, licenseId, StringComparison.Ordinal)) return null;
            return state;
        }
        catch { return null; }
    }
}
