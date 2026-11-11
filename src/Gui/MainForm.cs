using System.Reflection;

namespace IMAP2M365.Gui;

/// <summary>
/// Hauptfenster, graue Standard-WinForms-Optik wie in der übrigen Tool-Suite.
/// Zwei Reiter: „Einrichtung" (Wizard) und „Migration". Gemeinsames Log unten,
/// Versionsanzeige (Commit) in der Statusleiste.
/// </summary>
public sealed class MainForm : Form
{
    private readonly AppConfig _config = new();
    private readonly TextBox _log = new()
    {
        Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
        Dock = DockStyle.Fill, Font = new Font("Consolas", 8.5f)
    };
    private readonly StatusStrip _status = new();
    private readonly ToolStripStatusLabel _versionLabel = new();

    private SetupTab _setupTab = null!;
    private MigrationTab _migrationTab = null!;

    public MainForm()
    {
        Text = "IMAP2M365 – IMAP nach Microsoft 365";
        Width = 1040;
        Height = 720;
        StartPosition = FormStartPosition.CenterScreen;

        var tabs = new TabControl { Dock = DockStyle.Fill };

        _setupTab = new SetupTab(_config, Log);
        _migrationTab = new MigrationTab(_config, Log);

        tabs.TabPages.Add(new TabPage("Einrichtung") { Controls = { _setupTab } });
        tabs.TabPages.Add(new TabPage("Migration") { Controls = { _migrationTab } });
        _setupTab.Dock = DockStyle.Fill;
        _migrationTab.Dock = DockStyle.Fill;

        // Log-Panel unten, per Splitter abtrennbar
        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            SplitterDistance = 480
        };
        split.Panel1.Controls.Add(tabs);
        var logBox = new GroupBox { Text = "Protokoll", Dock = DockStyle.Fill };
        logBox.Controls.Add(_log);
        split.Panel2.Controls.Add(logBox);

        Controls.Add(split);

        _versionLabel.Text = "Version: " + BuildVersion();
        _status.Items.Add(_versionLabel);
        _status.Dock = DockStyle.Bottom;
        Controls.Add(_status);
    }

    /// <summary>Thread-sicheres Log, das setup/migration gemeinsam nutzen.</summary>
    public void Log(string line)
    {
        if (InvokeRequired) { BeginInvoke(() => Log(line)); return; }
        _log.AppendText($"{DateTime.Now:HH:mm:ss}  {line}{Environment.NewLine}");
    }

    private static string BuildVersion()
    {
        // Commit-Hash wird beim Build via InformationalVersion eingebettet (wie in der Suite üblich).
        var info = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        return string.IsNullOrWhiteSpace(info) ? "dev" : info;
    }
}
