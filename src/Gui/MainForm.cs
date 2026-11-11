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
    private LicenseTab _licenseTab = null!;

    private readonly MachineId.Info _machine = MachineId.Current();
    private readonly Licensing.LicenseManager _license;

    private const string BaseTitle = "IMAP2M365 – IMAP nach Microsoft 365";
    private string? _currentPath;       // Pfad der aktuell geladenen .enc
    private string? _masterPassword;    // Session-Cache, nicht persistiert

    private static string ProfilesDir
    {
        get
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "IMAP2M365", "Profiles");
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    /// <summary>Lizenz-/Verbrauchsdaten maschinenweit (ProgramData), Fallback LocalAppData.</summary>
    private static string LicenseDir
    {
        get
        {
            foreach (var folder in new[] { Environment.SpecialFolder.CommonApplicationData, Environment.SpecialFolder.LocalApplicationData })
            {
                try
                {
                    var dir = Path.Combine(Environment.GetFolderPath(folder), "IMAP2M365");
                    Directory.CreateDirectory(dir);
                    return dir;
                }
                catch { /* nächster Kandidat */ }
            }
            return Path.Combine(Path.GetTempPath(), "IMAP2M365");
        }
    }

    public MainForm()
    {
        Text = BaseTitle;
        Width = 1040;
        Height = 720;
        StartPosition = FormStartPosition.CenterScreen;

        // Fenster-/Taskbar-Icon aus dem eingebetteten EXE-Icon (ApplicationIcon) setzen.
        try { Icon = System.Drawing.Icon.ExtractAssociatedIcon(Environment.ProcessPath!); } catch { /* egal */ }

        _license = new Licensing.LicenseManager(LicenseDir, _machine.Id);

        var tabs = new TabControl { Dock = DockStyle.Fill };

        _setupTab = new SetupTab(_config, Log);
        _migrationTab = new MigrationTab(_config, Log, _license, () => _licenseTab?.UpdateStatus());
        _licenseTab = new LicenseTab(_license, _machine, Log);

        tabs.TabPages.Add(new TabPage("Einrichtung") { Controls = { _setupTab } });
        tabs.TabPages.Add(new TabPage("Migration") { Controls = { _migrationTab } });
        tabs.TabPages.Add(new TabPage("Lizenz") { Controls = { _licenseTab } });
        _setupTab.Dock = DockStyle.Fill;
        _migrationTab.Dock = DockStyle.Fill;
        _licenseTab.Dock = DockStyle.Fill;

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

        // Profil-/Kunden-Menü zuletzt hinzufügen -> dockt oben über die Reiter.
        var menu = new MenuStrip();
        var profile = new ToolStripMenuItem("Profil/Kunde");
        profile.DropDownItems.Add("Neu", null, (_, _) => NewProfile());
        profile.DropDownItems.Add("Laden…", null, (_, _) => LoadProfile());
        profile.DropDownItems.Add(new ToolStripSeparator());
        profile.DropDownItems.Add("Speichern", null, (_, _) => SaveProfile());
        profile.DropDownItems.Add("Speichern unter…", null, (_, _) => SaveProfileAs());
        menu.Items.Add(profile);
        MainMenuStrip = menu;
        Controls.Add(menu);
    }

    // --- Profil-/Kundenverwaltung (verschlüsselte .enc je Kunde) ---

    private void NewProfile()
    {
        if (MessageBox.Show("Aktuelles Profil verwerfen und neu beginnen?", "Neues Profil",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;
        _config.CopyFrom(new AppConfig());
        _currentPath = null; _masterPassword = null;
        PushConfigToTabs();
        UpdateTitle();
        Log("Neues, leeres Profil.");
    }

    private void LoadProfile()
    {
        using var dlg = new OpenFileDialog { Filter = "IMAP2M365-Profil (*.enc)|*.enc", InitialDirectory = ProfilesDir };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        var pw = PromptPassword("Master-Passwort für dieses Profil:");
        if (pw is null) return;

        try
        {
            var loaded = ConfigStore.Load(dlg.FileName, pw);
            _config.CopyFrom(loaded);
            _currentPath = dlg.FileName; _masterPassword = pw;
            PushConfigToTabs();
            UpdateTitle();
            Log($"Profil geladen: {Path.GetFileNameWithoutExtension(dlg.FileName)} ({_config.Mailboxes.Count} Postfächer).");
        }
        catch (Exception ex)
        {
            Log("Laden fehlgeschlagen: " + ex.Message);
            MessageBox.Show(ex.Message, "Laden fehlgeschlagen", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void SaveProfile()
    {
        if (_currentPath is null) { SaveProfileAs(); return; }
        SaveTo(_currentPath, reusePassword: true);
    }

    private void SaveProfileAs()
    {
        using var dlg = new SaveFileDialog
        {
            Filter = "IMAP2M365-Profil (*.enc)|*.enc",
            InitialDirectory = ProfilesDir,
            FileName = string.IsNullOrWhiteSpace(_config.ProfileName) ? "kunde.enc" : _config.ProfileName + ".enc"
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;
        SaveTo(dlg.FileName, reusePassword: false);
    }

    private void SaveTo(string path, bool reusePassword)
    {
        var pw = reusePassword ? _masterPassword : null;
        pw ??= PromptPassword("Master-Passwort zum Verschlüsseln:");
        if (pw is null) return;

        try
        {
            FlushTabsToConfig();
            _config.ProfileName = Path.GetFileNameWithoutExtension(path);
            ConfigStore.Save(path, _config, pw);
            _currentPath = path; _masterPassword = pw;
            UpdateTitle();
            Log($"Profil gespeichert: {_config.ProfileName}.");
        }
        catch (Exception ex)
        {
            Log("Speichern fehlgeschlagen: " + ex.Message);
            MessageBox.Show(ex.Message, "Speichern fehlgeschlagen", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void FlushTabsToConfig() { _setupTab.FlushToConfig(); _migrationTab.FlushToConfig(); }
    private void PushConfigToTabs() { _setupTab.LoadFromConfig(); _migrationTab.LoadFromConfig(); }
    private void UpdateTitle() =>
        Text = string.IsNullOrWhiteSpace(_config.ProfileName) ? BaseTitle : $"{BaseTitle} – {_config.ProfileName}";

    private static string? PromptPassword(string prompt)
    {
        using var dlg = new InputDialog(prompt, password: true);
        return dlg.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(dlg.Value) ? dlg.Value : null;
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
