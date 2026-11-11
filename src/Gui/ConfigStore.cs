using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace IMAP2M365.Gui;

/// <summary>
/// AES-256-verschlüsselte Konfig-Ablage (config.enc), wie in den anderen Tools der Suite.
/// Schlüsselableitung aus einem Master-Passwort via PBKDF2; Salt + IV werden der Datei vorangestellt.
/// Format: [16 Byte Salt][16 Byte IV][AES-256-CBC Ciphertext].
/// </summary>
public static class ConfigStore
{
    private const int SaltLen = 16;
    private const int IvLen = 16;
    private const int KeyLen = 32;           // AES-256
    private const int Iterations = 200_000;

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true
    };

    public static void Save(string path, AppConfig config, string masterPassword)
    {
        var plaintext = JsonSerializer.SerializeToUtf8Bytes(config, JsonOpts);

        var salt = RandomNumberGenerator.GetBytes(SaltLen);
        using var kdf = new Rfc2898DeriveBytes(masterPassword, salt, Iterations, HashAlgorithmName.SHA256);
        var key = kdf.GetBytes(KeyLen);

        using var aes = Aes.Create();
        aes.KeySize = 256;
        aes.Key = key;
        aes.GenerateIV();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.PKCS7;

        using var enc = aes.CreateEncryptor();
        var cipher = enc.TransformFinalBlock(plaintext, 0, plaintext.Length);

        using var fs = File.Create(path);
        fs.Write(salt);
        fs.Write(aes.IV);
        fs.Write(cipher);
    }

    public static AppConfig Load(string path, string masterPassword)
    {
        var raw = File.ReadAllBytes(path);
        if (raw.Length < SaltLen + IvLen)
            throw new InvalidDataException("config.enc ist beschädigt oder leer.");

        var salt = raw.AsSpan(0, SaltLen).ToArray();
        var iv = raw.AsSpan(SaltLen, IvLen).ToArray();
        var cipher = raw.AsSpan(SaltLen + IvLen).ToArray();

        using var kdf = new Rfc2898DeriveBytes(masterPassword, salt, Iterations, HashAlgorithmName.SHA256);
        var key = kdf.GetBytes(KeyLen);

        using var aes = Aes.Create();
        aes.KeySize = 256;
        aes.Key = key;
        aes.IV = iv;
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.PKCS7;

        try
        {
            using var dec = aes.CreateDecryptor();
            var plain = dec.TransformFinalBlock(cipher, 0, cipher.Length);
            return JsonSerializer.Deserialize<AppConfig>(plain, JsonOpts)
                   ?? throw new InvalidDataException("Konfiguration konnte nicht gelesen werden.");
        }
        catch (CryptographicException)
        {
            throw new UnauthorizedAccessException("Falsches Master-Passwort oder beschädigte config.enc.");
        }
    }

    /// <summary>CSV-Import/-Export der Postfachzeilen (ohne Klartext-Passwörter im Export, außer explizit gewünscht).</summary>
    public static List<MailboxRow> ImportCsv(string path)
    {
        var rows = new List<MailboxRow>();
        foreach (var line in File.ReadLines(path).Skip(1))
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            var c = line.Split(',');
            rows.Add(new MailboxRow
            {
                SourceUser = c.ElementAtOrDefault(0)?.Trim() ?? "",
                SourcePassword = c.ElementAtOrDefault(1)?.Trim() ?? "",
                TargetSmtp = c.ElementAtOrDefault(2)?.Trim() ?? ""
            });
        }
        return rows;
    }

    public static void ExportCsv(string path, IEnumerable<MailboxRow> rows, bool includePasswords)
    {
        var sb = new StringBuilder();
        sb.AppendLine("SourceUser,SourcePassword,TargetUpn");
        foreach (var r in rows)
            sb.AppendLine($"{r.SourceUser},{(includePasswords ? r.SourcePassword : "")},{r.TargetSmtp}");
        File.WriteAllText(path, sb.ToString());
    }
}
