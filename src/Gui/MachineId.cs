using System.Management;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace IMAP2M365.Gui;

/// <summary>
/// Node-locking: erzeugt eine stabile Computer-ID aus MachineGuid (Registry) + BIOS-UUID (WMI)
/// + Rechnername und erkennt gängige virtuelle Maschinen. Die Computer-ID ist ein Hash – sie
/// gibt keine Hardware-Details preis, bindet die Lizenz aber an diesen Rechner.
/// </summary>
public static class MachineId
{
    public sealed record Info(string Id, string MachineName, bool IsVirtual, string VmHint);

    public static Info Current()
    {
        var machineGuid = ReadMachineGuid();
        var (biosUuid, manufacturer, model) = ReadSystemInfo();
        var name = Environment.MachineName;

        var raw = $"{machineGuid}|{biosUuid}|{name}";
        var id = Hash(raw);

        var (isVm, hint) = DetectVm(manufacturer, model, biosUuid);
        return new Info(id, name, isVm, hint);
    }

    private static string ReadMachineGuid()
    {
        try
        {
            using var key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                .OpenSubKey(@"SOFTWARE\Microsoft\Cryptography", false);
            return key?.GetValue("MachineGuid")?.ToString() ?? "";
        }
        catch { return ""; }
    }

    private static (string uuid, string manufacturer, string model) ReadSystemInfo()
    {
        string uuid = "", manufacturer = "", model = "";
        try
        {
            using (var s = new ManagementObjectSearcher("SELECT UUID FROM Win32_ComputerSystemProduct"))
                foreach (ManagementObject o in s.Get()) { uuid = o["UUID"]?.ToString() ?? ""; break; }

            using (var s = new ManagementObjectSearcher("SELECT Manufacturer, Model FROM Win32_ComputerSystem"))
                foreach (ManagementObject o in s.Get())
                {
                    manufacturer = o["Manufacturer"]?.ToString() ?? "";
                    model = o["Model"]?.ToString() ?? "";
                    break;
                }
        }
        catch { /* WMI nicht verfügbar -> ohne diese Signale */ }
        return (uuid, manufacturer, model);
    }

    private static (bool, string) DetectVm(string manufacturer, string model, string biosUuid)
    {
        var hay = $"{manufacturer} {model}".ToLowerInvariant();
        string[] markers = { "vmware", "virtualbox", "oracle vm", "innotek", "qemu", "kvm", "xen",
                             "hyper-v", "virtual machine", "microsoft corporation|virtual", "parallels", "bochs", "bhyve" };
        foreach (var m in markers)
        {
            if (m.Contains('|'))
            {
                var parts = m.Split('|');
                if (parts.All(p => hay.Contains(p))) return (true, $"{manufacturer} {model}".Trim());
            }
            else if (hay.Contains(m))
            {
                return (true, $"{manufacturer} {model}".Trim());
            }
        }
        return (false, "");
    }

    private static string Hash(string input)
    {
        using var sha = SHA256.Create();
        var h = sha.ComputeHash(Encoding.UTF8.GetBytes(input));
        // 16 Bytes -> 32 Hex, in 4er-Gruppen für Lesbarkeit
        var hex = Convert.ToHexString(h, 0, 16);
        return string.Join('-', Enumerable.Range(0, hex.Length / 4).Select(i => hex.Substring(i * 4, 4)));
    }
}
