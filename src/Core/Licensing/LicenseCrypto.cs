using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace IMAP2M365.Licensing;

/// <summary>
/// Kryptografie für das Lizenzsystem (RSA-3072):
///  - Lizenz signieren/prüfen  (RSA-PSS SHA-256)
///  - Aktivierungsanfrage hybrid ver-/entschlüsseln (AES-256-GCM + RSA-OAEP-SHA256)
///  - Schlüsselpaar erzeugen (für den Generator)
///
/// Tool-Seite nutzt nur den öffentlichen Schlüssel (Prüfen + Verschlüsseln).
/// Generator-Seite nutzt den privaten Schlüssel (Signieren + Entschlüsseln).
/// </summary>
public static class LicenseCrypto
{
    private const string ReqPrefix = "IMAP2M365-REQ1.";
    private const string LicPrefix = "IMAP2M365-LIC1.";

    private static readonly JsonSerializerOptions Json = new() { WriteIndented = false };

    // ---------- Schlüsselpaar (Generator, einmalig) ----------

    public static (string publicPem, string privatePem) CreateKeyPair()
    {
        using var rsa = RSA.Create(3072);
        return (rsa.ExportSubjectPublicKeyInfoPem(), rsa.ExportPkcs8PrivateKeyPem());
    }

    // ---------- Lizenz signieren / prüfen ----------

    public static string SignLicense(LicensePayload payload, string privatePem)
    {
        var payloadBytes = JsonSerializer.SerializeToUtf8Bytes(payload, Json);

        using var rsa = RSA.Create();
        rsa.ImportFromPem(privatePem);
        var sig = rsa.SignData(payloadBytes, HashAlgorithmName.SHA256, RSASignaturePadding.Pss);

        var license = new License { Payload = payload, Signature = Convert.ToBase64String(sig) };
        var licJson = JsonSerializer.SerializeToUtf8Bytes(license, Json);
        return LicPrefix + Base64Url(licJson);
    }

    /// <summary>Prüft Signatur und gibt die Lizenz zurück, oder null bei ungültig.</summary>
    public static License? VerifyLicense(string licenseKey, string publicPem)
    {
        try
        {
            if (!licenseKey.StartsWith(LicPrefix, StringComparison.Ordinal)) return null;
            var licJson = Base64UrlDecode(licenseKey[LicPrefix.Length..]);
            var license = JsonSerializer.Deserialize<License>(licJson, Json);
            if (license is null) return null;

            var payloadBytes = JsonSerializer.SerializeToUtf8Bytes(license.Payload, Json);
            using var rsa = RSA.Create();
            rsa.ImportFromPem(publicPem);
            var ok = rsa.VerifyData(payloadBytes, Convert.FromBase64String(license.Signature),
                HashAlgorithmName.SHA256, RSASignaturePadding.Pss);
            return ok ? license : null;
        }
        catch { return null; }
    }

    // ---------- Aktivierungsanfrage hybrid ver-/entschlüsseln ----------

    public static string EncryptRequest(ActivationRequest request, string publicPem)
    {
        var plain = JsonSerializer.SerializeToUtf8Bytes(request, Json);

        // AES-256-GCM mit Zufallsschlüssel
        var aesKey = RandomNumberGenerator.GetBytes(32);
        var nonce = RandomNumberGenerator.GetBytes(12);
        var cipher = new byte[plain.Length];
        var tag = new byte[16];
        using (var gcm = new AesGcm(aesKey, 16))
            gcm.Encrypt(nonce, plain, cipher, tag);

        // AES-Schlüssel mit RSA-OAEP(SHA-256) verschlüsseln
        using var rsa = RSA.Create();
        rsa.ImportFromPem(publicPem);
        var encKey = rsa.Encrypt(aesKey, RSAEncryptionPadding.OaepSHA256);

        // Paket: [2B encKeyLen][encKey][12B nonce][16B tag][cipher]
        using var ms = new MemoryStream();
        ms.WriteByte((byte)(encKey.Length >> 8));
        ms.WriteByte((byte)(encKey.Length & 0xFF));
        ms.Write(encKey);
        ms.Write(nonce);
        ms.Write(tag);
        ms.Write(cipher);
        return ReqPrefix + Base64Url(ms.ToArray());
    }

    public static ActivationRequest DecryptRequest(string requestString, string privatePem)
    {
        var s = requestString.Trim();
        if (s.StartsWith(ReqPrefix, StringComparison.Ordinal)) s = s[ReqPrefix.Length..];
        var raw = Base64UrlDecode(s);

        int pos = 0;
        int encKeyLen = (raw[pos] << 8) | raw[pos + 1]; pos += 2;
        var encKey = raw.AsSpan(pos, encKeyLen).ToArray(); pos += encKeyLen;
        var nonce = raw.AsSpan(pos, 12).ToArray(); pos += 12;
        var tag = raw.AsSpan(pos, 16).ToArray(); pos += 16;
        var cipher = raw.AsSpan(pos).ToArray();

        using var rsa = RSA.Create();
        rsa.ImportFromPem(privatePem);
        var aesKey = rsa.Decrypt(encKey, RSAEncryptionPadding.OaepSHA256);

        var plain = new byte[cipher.Length];
        using (var gcm = new AesGcm(aesKey, 16))
            gcm.Decrypt(nonce, cipher, tag, plain);

        return JsonSerializer.Deserialize<ActivationRequest>(plain, Json)
               ?? throw new InvalidDataException("Aktivierungsanfrage konnte nicht gelesen werden.");
    }

    // ---------- Base64Url ----------

    private static string Base64Url(byte[] data) =>
        Convert.ToBase64String(data).Replace('+', '-').Replace('/', '_').TrimEnd('=');

    private static byte[] Base64UrlDecode(string s)
    {
        s = s.Replace('-', '+').Replace('_', '/');
        switch (s.Length % 4) { case 2: s += "=="; break; case 3: s += "="; break; }
        return Convert.FromBase64String(s);
    }
}
