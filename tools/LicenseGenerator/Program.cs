using IMAP2M365.Licensing;

namespace IMAP2M365.LicenseGen;

/// <summary>
/// Konsolen-Generator für den Lizenzgeber. Hält den privaten Schlüssel (private.pem).
/// Menü:
///   1) Schlüsselpaar erzeugen  -> public.pem (ins Tool einbetten) + private.pem (geheim halten)
///   2) Aktivierungsanfrage entschlüsseln  -> zeigt Kundendaten + Computer-ID + VM-Flag
///   3) Lizenz erstellen  -> signierter Lizenzschlüssel für den Kunden
/// </summary>
internal static class Program
{
    private const string PrivatePath = "private.pem";
    private const string PublicPath = "public.pem";

    private static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.WriteLine("IMAP2M365 – Lizenz-Generator\n");

        while (true)
        {
            Console.WriteLine("\n[1] Schlüsselpaar erzeugen   [2] Anfrage entschlüsseln   [3] Lizenz erstellen   [0] Ende");
            Console.Write("> ");
            switch (Console.ReadLine()?.Trim())
            {
                case "1": CreateKeyPair(); break;
                case "2": DecryptRequest(); break;
                case "3": CreateLicense(); break;
                case "0": return;
                default: Console.WriteLine("Unbekannt."); break;
            }
        }
    }

    private static void CreateKeyPair()
    {
        if (File.Exists(PrivatePath))
        {
            Console.Write($"{PrivatePath} existiert bereits. Überschreiben? (j/N) ");
            if (!string.Equals(Console.ReadLine()?.Trim(), "j", StringComparison.OrdinalIgnoreCase)) return;
        }
        var (pub, priv) = LicenseCrypto.CreateKeyPair();
        File.WriteAllText(PublicPath, pub);
        File.WriteAllText(PrivatePath, priv);
        Console.WriteLine($"\nErzeugt:\n  {PrivatePath}  (GEHEIM halten – niemals weitergeben!)\n  {PublicPath}   (in EmbeddedPublicKey.cs des Tools einsetzen)\n");
        Console.WriteLine("Öffentlicher Schlüssel:\n");
        Console.WriteLine(pub);
    }

    private static string? LoadPrivate()
    {
        if (File.Exists(PrivatePath)) return File.ReadAllText(PrivatePath);
        Console.Write("Pfad zur private.pem: ");
        var p = Console.ReadLine()?.Trim('"', ' ');
        if (string.IsNullOrWhiteSpace(p) || !File.Exists(p)) { Console.WriteLine("Nicht gefunden."); return null; }
        return File.ReadAllText(p);
    }

    private static ActivationRequest? ReadAndDecrypt(string privatePem)
    {
        Console.WriteLine("Aktivierungsanfrage einfügen (String oder Pfad zu .req), Enter:");
        var input = Console.ReadLine()?.Trim() ?? "";
        if (File.Exists(input)) input = File.ReadAllText(input);
        try { return LicenseCrypto.DecryptRequest(input, privatePem); }
        catch (Exception ex) { Console.WriteLine("Entschlüsselung fehlgeschlagen: " + ex.Message); return null; }
    }

    private static void DecryptRequest()
    {
        var priv = LoadPrivate(); if (priv is null) return;
        var req = ReadAndDecrypt(priv); if (req is null) return;
        PrintRequest(req);
    }

    private static void PrintRequest(ActivationRequest req)
    {
        Console.WriteLine("\n--- Aktivierungsanfrage ---");
        Console.WriteLine($"  Firma:          {req.Company}");
        Console.WriteLine($"  Ansprechpartner:{req.Contact}");
        Console.WriteLine($"  Adresse:        {req.Address}");
        Console.WriteLine($"  Mail:           {req.Email}");
        Console.WriteLine($"  Computer-ID:    {req.MachineId}");
        Console.WriteLine($"  Rechner:        {req.MachineName}");
        Console.WriteLine($"  VM:             {(req.IsVirtualMachine ? "JA  -> " + req.VmHint : "nein")}");
        Console.WriteLine($"  Tool-Version:   {req.ToolVersion}");
        Console.WriteLine($"  Erstellt (UTC): {req.CreatedUtc:u}");
        Console.WriteLine("---------------------------\n");
    }

    private static void CreateLicense()
    {
        var priv = LoadPrivate(); if (priv is null) return;
        var req = ReadAndDecrypt(priv); if (req is null) return;
        PrintRequest(req);

        if (req.IsVirtualMachine)
        {
            Console.Write("Achtung: VM erkannt. Trotzdem lizenzieren? (j/N) ");
            if (!string.Equals(Console.ReadLine()?.Trim(), "j", StringComparison.OrdinalIgnoreCase)) return;
        }

        Console.Write("Anzahl Postfächer (Kontingent): ");
        if (!int.TryParse(Console.ReadLine()?.Trim(), out var count) || count <= 0)
        {
            Console.WriteLine("Ungültige Anzahl."); return;
        }

        var payload = new LicensePayload
        {
            MachineId = req.MachineId,
            MaxAccounts = count,
            Company = req.Company,
            Contact = req.Contact,
            Email = req.Email
        };

        var key = LicenseCrypto.SignLicense(payload, priv);
        var outFile = $"lizenz_{Sanitize(req.Company)}_{count}.key";
        File.WriteAllText(outFile, key);

        Console.WriteLine($"\nLizenzschlüssel erstellt ({count} Postfächer, gebunden an {req.MachineId}):");
        Console.WriteLine($"  Datei: {outFile}\n");
        Console.WriteLine(key);
        Console.WriteLine("\n-> an den Kunden senden; er fügt ihn im Reiter „Lizenz" ein und aktiviert.");
    }

    private static string Sanitize(string s)
    {
        var clean = new string(s.Where(c => char.IsLetterOrDigit(c)).ToArray());
        return string.IsNullOrEmpty(clean) ? "kunde" : clean;
    }
}
