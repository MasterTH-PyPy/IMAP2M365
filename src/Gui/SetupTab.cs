using IMAP2M365.Sources;

namespace IMAP2M365.Gui;

/// <summary>
/// Einrichtungs-Wizard: Entra-App provisionieren, mit EXO verbinden, Postfächer laden
/// (User + Shared), zentralen Quell-IMAP-Server setzen, pro Zeile Quell-User/Kennwort
/// erfassen, IMAP je Postfach aktivieren, Rechte vergeben, optional Policy, testen.
/// Alles mit Dry-Run-Schalter (nur prüfen, nichts ändern).
/// </summary>
public sealed class SetupTab : UserControl
{
    private readonly AppConfig _config;
    private readonly Action<string> _log;

    // Tenant / zentrale Quelle
    private readonly TextBox _tenant = new() { Width = 260 };
    private readonly TextBox _srcHost = new() { Width = 200 };
    private readonly NumericUpDown _srcPort = new() { Minimum = 1, Maximum = 65535, Value = 993, Width = 70 };
    private readonly CheckBox _srcSsl = new() { Text = "SSL", Checked = true, AutoSize = true };
    private readonly CheckBox _dryRun = new() { Text = "Dry-Run (nur prüfen)", Checked = true, AutoSize = true };
    private readonly CheckBox _includeShared = new() { Text = "Freigegebene Postfächer einschließen", Checked = true, AutoSize = true };

    // App-Registrierung: nach Schritt 1 befüllt, aber auch manuell für bestehende Apps eingebbar.
    private readonly TextBox _clientId = new() { Width = 260 };
    private readonly TextBox _secret = new() { Width = 200, UseSystemPasswordChar = true };
    private readonly TextBox _spObjectId = new() { Width = 200 };

    private static Button StepButton(string text) => new()
    {
        Text = text,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        MinimumSize = new Size(0, 30),
        Padding = new Padding(12, 4, 12, 4),
        Margin = new Padding(4, 6, 4, 6)
    };

    private readonly Button _provision = StepButton("1a. Entra-App anlegen…");
    private readonly Button _readApp = StepButton("1b. App auslesen…");
    private readonly Button _connect = StepButton("2. Mit EXO verbinden");
    private readonly Button _loadMbx = StepButton("3. Postfächer laden");
    private readonly Button _apply = StepButton("4. Einrichtung ausführen");

    private readonly DataGridView _grid = new()
    {
        Dock = DockStyle.Fill, AllowUserToAddRows = false,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect
    };

    private ExoRunner? _exo;

    public SetupTab(AppConfig config, Action<string> log)
    {
        _config = config;
        _log = log;
        BuildUi();
        BuildGridColumns();
        WireEvents();
    }

    private void BuildUi()
    {
        var top = new TableLayoutPanel { Dock = DockStyle.Top, Height = 150, ColumnCount = 6, Padding = new Padding(8) };

        top.Controls.Add(new Label { Text = "Tenant (…onmicrosoft.com):", AutoSize = true }, 0, 0);
        top.Controls.Add(_tenant, 1, 0);
        top.Controls.Add(_includeShared, 2, 0);
        top.SetColumnSpan(_includeShared, 2);

        top.Controls.Add(new Label { Text = "Quell-IMAP (zentral):", AutoSize = true }, 0, 1);
        top.Controls.Add(_srcHost, 1, 1);
        top.Controls.Add(_srcPort, 2, 1);
        top.Controls.Add(_srcSsl, 3, 1);
        top.Controls.Add(_dryRun, 4, 1);

        // App-Registrierung: ClientId/Secret/ServicePrincipal-ObjectId sichtbar + editierbar.
        top.Controls.Add(new Label { Text = "App ClientId:", AutoSize = true }, 0, 2);
        top.Controls.Add(_clientId, 1, 2);
        top.Controls.Add(new Label { Text = "Secret:", AutoSize = true }, 2, 2);
        top.Controls.Add(_secret, 3, 2);
        top.Controls.Add(new Label { Text = "SP-ObjectId:", AutoSize = true }, 4, 2);
        top.Controls.Add(_spObjectId, 5, 2);

        var buttons = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 52, WrapContents = false, Padding = new Padding(8, 6, 8, 6) };
        buttons.Controls.AddRange(new Control[] { _provision, _readApp, _connect, _loadMbx, _apply });

        Controls.Add(_grid);
        Controls.Add(buttons);
        Controls.Add(top);
    }

    private void BuildGridColumns()
    {
        _grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Migrieren", DataPropertyName = "Enabled", FillWeight = 45 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Ziel-Postfach", DataPropertyName = "TargetSmtp", ReadOnly = true });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Typ", DataPropertyName = "MailboxType", ReadOnly = true, FillWeight = 60 });
        _grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "IMAP aktiv", DataPropertyName = "ImapEnabled", FillWeight = 55 });
        _grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Recht ok", DataPropertyName = "PermissionGranted", ReadOnly = true, FillWeight = 50 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Quell-User", DataPropertyName = "SourceUser" });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Quell-Kennwort", DataPropertyName = "SourcePassword" });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Quell-Server (optional)", DataPropertyName = "SourceHostOverride" });
        _grid.Columns.Add(new DataGridViewButtonColumn { HeaderText = "Test", Text = "Test", UseColumnTextForButtonValue = true, FillWeight = 45 });

        // Kennwort-Spalte maskieren (nur Anzeige; der echte Wert bleibt im gebundenen Objekt)
        _grid.CellFormatting += (s, e) =>
        {
            if (_grid.Columns[e.ColumnIndex].DataPropertyName == "SourcePassword" && e.Value is string pw && pw.Length > 0)
                e.Value = new string('•', Math.Min(pw.Length, 10));
        };

        // Klick auf „Test" prüft die Quell-IMAP-Einstellungen genau dieser Zeile.
        _grid.CellContentClick += async (s, e) =>
        {
            if (e.RowIndex < 0) return;
            if (!string.Equals(_grid.Columns[e.ColumnIndex].HeaderText, "Test", StringComparison.Ordinal)) return;
            if (_grid.Rows[e.RowIndex].DataBoundItem is MailboxRow row) await TestRowAsync(row);
        };
    }

    /// <summary>Testet Verbindung + Login der Quelle für eine Zeile (zentraler Server oder Zeilen-Override).</summary>
    private async Task TestRowAsync(MailboxRow row)
    {
        _grid.EndEdit();

        var host = (string.IsNullOrWhiteSpace(row.SourceHostOverride) ? _srcHost.Text : row.SourceHostOverride!).Trim();
        var port = row.SourcePortOverride ?? (int)_srcPort.Value;
        var ssl = _srcSsl.Checked;

        if (string.IsNullOrWhiteSpace(host)) { _log($"Test {row.TargetSmtp}: kein Quell-Server gesetzt."); return; }
        if (string.IsNullOrWhiteSpace(row.SourceUser)) { _log($"Test {row.TargetSmtp}: kein Quell-User gesetzt."); return; }

        _log($"Test {row.TargetSmtp}: verbinde {host}:{port} (SSL={ssl}) als {row.SourceUser} …");
        try
        {
            await using var src = new ImapSource(host, port, ssl, row.SourceUser, row.SourcePassword);
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            await src.ConnectAsync(cts.Token);
            var folders = await src.ListFoldersAsync(cts.Token);
            _log($"Test {row.TargetSmtp}: OK – Login erfolgreich, {folders.Count} Ordner gefunden.");
        }
        catch (Exception ex)
        {
            _log($"Test {row.TargetSmtp}: FEHLER – {ex.Message}");
        }
    }

    private void WireEvents()
    {
        _provision.Click += async (_, _) => await ProvisionAsync();
        _readApp.Click += async (_, _) => await ReadExistingAppAsync();
        _connect.Click += (_, _) => Connect();
        _loadMbx.Click += (_, _) => LoadMailboxes();
        _apply.Click += (_, _) => ApplySetup();
    }

    /// <summary>Überträgt die App-Werte aus der Konfig in die Eingabefelder.</summary>
    private void ShowAppFields()
    {
        _clientId.Text = _config.ClientId;
        _secret.Text = _config.ClientSecret;
        _spObjectId.Text = _config.SpObjectId;
    }

    /// <summary>
    /// Liest vorhandene App-Registrierungen aus dem Tenant (Graph), lässt auswählen und übernimmt
    /// ClientId + ServicePrincipal-ObjectId. Ein frisches Secret wird erzeugt (bestehende Secrets
    /// sind nicht auslesbar).
    /// </summary>
    private async Task ReadExistingAppAsync()
    {
        try
        {
            _config.TenantId = _tenant.Text.Trim();
            if (string.IsNullOrWhiteSpace(_config.TenantId)) { _log("Bitte Tenant angeben."); return; }

            var graph = new GraphProvisioner { Log = _log };
            var apps = await graph.ListApplicationsAsync(_config.TenantId, CancellationToken.None);
            if (apps.Count == 0) { _log("Keine App-Registrierungen gefunden. Bitte Schritt 1a verwenden."); return; }

            var choice = ChooserDialog.Pick("App-Registrierung auswählen:",
                apps.Select(a => $"{a.DisplayName}  —  {a.AppId}").ToList());
            if (choice < 0) return;

            var picked = apps[choice];
            var includeGraph = _config.TargetKind.Equals("Graph", StringComparison.OrdinalIgnoreCase);
            var (spObj, secret) = await graph.PrepareExistingAppAsync(
                _config.TenantId, picked.AppObjectId, picked.AppId, includeGraph, CancellationToken.None);

            _config.ClientId = picked.AppId;
            _config.SpObjectId = spObj;
            _config.ClientSecret = secret;
            ShowAppFields();
            _log($"App übernommen: {picked.DisplayName} (ClientId {picked.AppId}). Frisches Secret erzeugt.");
        }
        catch (Exception ex)
        {
            _log("App auslesen fehlgeschlagen: " + ex.Message);
            for (var inner = ex.InnerException; inner != null; inner = inner.InnerException)
                _log("  Ursache: " + inner.Message);
        }
    }

    private async Task ProvisionAsync()
    {
        try
        {
            _config.TenantId = _tenant.Text.Trim();
            if (string.IsNullOrWhiteSpace(_config.TenantId)) { _log("Bitte Tenant angeben."); return; }

            var graph = new GraphProvisioner { Log = _log };
            var includeGraph = _config.TargetKind.Equals("Graph", StringComparison.OrdinalIgnoreCase);
            var (appId, spObj, secret) = await graph.ProvisionAsync(
                _config.TenantId, "IMAP2M365 Migration", includeGraph, CancellationToken.None);

            _config.ClientId = appId;
            _config.ClientSecret = secret;
            _config.SpObjectId = spObj;
            ShowAppFields();
            _log($"Fertig: ClientId {appId}. Secret und ServicePrincipal-ObjectId liegen in der Konfiguration.");
            _log("Hinweis: ServicePrincipal in EXO wird in Schritt 4 registriert.");
        }
        catch (Exception ex)
        {
            _log("Provisionierung fehlgeschlagen: " + ex.Message);
            for (var inner = ex.InnerException; inner != null; inner = inner.InnerException)
                _log("  Ursache: " + inner.Message);
        }
    }

    private void Connect()
    {
        try
        {
            _config.TenantId = _tenant.Text.Trim();
            _exo ??= new ExoRunner { Log = _log };
            _exo.EnsureModule();
            _exo.Connect(_config.TenantId);
        }
        catch (Exception ex) { _log("EXO-Verbindung fehlgeschlagen: " + ex.Message); }
    }

    private void LoadMailboxes()
    {
        try
        {
            if (_exo is null) { _log("Erst mit EXO verbinden (Schritt 2)."); return; }
            var rows = _exo.LoadMailboxes(_includeShared.Checked);
            _config.Mailboxes = rows;
            Rebind();
        }
        catch (Exception ex) { _log("Postfächer laden fehlgeschlagen: " + ex.Message); }
    }

    /// <summary>Schreibt die aktuellen Eingabefelder (inkl. Grid) in die Konfiguration – für Speichern.</summary>
    public void FlushToConfig()
    {
        _config.TenantId = _tenant.Text.Trim();
        _config.SourceHost = _srcHost.Text.Trim();
        _config.SourcePort = (int)_srcPort.Value;
        _config.SourceUseSsl = _srcSsl.Checked;
        _config.DryRun = _dryRun.Checked;
        _config.ClientId = _clientId.Text.Trim();
        _config.ClientSecret = _secret.Text;
        _config.SpObjectId = _spObjectId.Text.Trim();
        CommitGridEdits();
    }

    /// <summary>Lädt die Konfiguration in die Eingabefelder – nach Profil-Load.</summary>
    public void LoadFromConfig()
    {
        _tenant.Text = _config.TenantId;
        _srcHost.Text = _config.SourceHost;
        _srcPort.Value = Math.Clamp(_config.SourcePort, (int)_srcPort.Minimum, (int)_srcPort.Maximum);
        _srcSsl.Checked = _config.SourceUseSsl;
        _dryRun.Checked = _config.DryRun;
        ShowAppFields();
        Rebind();
    }

    private void ApplySetup()
    {
        if (_exo is null) { _log("Erst mit EXO verbinden (Schritt 2)."); return; }

        FlushToConfig();
        var dry = _dryRun.Checked;
        _log(dry ? "Dry-Run: zeige nur, was geändert würde." : "Führe Einrichtung aus.");

        // Service-Principal einmalig registrieren
        var spObjectId = PromptMissingSpObjectId();
        if (spObjectId is null) return;

        if (!dry) _exo.RegisterServicePrincipal(_config.ClientId, spObjectId);

        foreach (var row in _config.Mailboxes)
        {
            if (!row.Enabled) { _log($"{row.TargetSmtp}: deaktiviert – übersprungen."); continue; }

            // IMAP aktivieren, wo nötig (betrifft oft Shared Mailboxes)
            if (!row.ImapEnabled)
            {
                _log($"{row.TargetSmtp}: IMAP {(dry ? "würde aktiviert" : "aktivieren")}" + (row.IsShared ? " (Shared)" : ""));
                if (!dry) { _exo.EnableImap(row.TargetSmtp); row.ImapEnabled = true; }
            }

            // Postfachrecht für den ServicePrincipal
            _log($"{row.TargetSmtp}: FullAccess {(dry ? "würde erteilt" : "erteilen")}");
            if (!dry) { _exo.GrantMailboxPermission(row.TargetSmtp, spObjectId); row.PermissionGranted = true; }
        }

        // Optional: Zugriff per Policy begrenzen
        if (!string.IsNullOrWhiteSpace(_config.ScopeGroupId) && !dry)
            _exo.SetApplicationAccessPolicy(_config.ClientId, _config.ScopeGroupId!);

        // Test
        if (!dry && _config.Mailboxes.Count > 0)
        {
            var sample = _config.Mailboxes[0].TargetSmtp;
            var ok = string.IsNullOrWhiteSpace(_config.ScopeGroupId) || _exo.TestAccess(_config.ClientId, sample);
            _log($"Zugriffstest ({sample}): {(ok ? "OK" : "VERWEIGERT – Policy/Gruppe prüfen")}");
        }

        Rebind();
        if (dry)
            _log("Dry-Run abgeschlossen – es wurde NICHTS geändert. Für die echte Einrichtung den Haken Dry-Run entfernen und erneut ausführen.");
        else
            _log("Einrichtung abgeschlossen (ServicePrincipal, IMAP, Rechte gesetzt).");
    }

    private string? PromptMissingSpObjectId()
    {
        // Normalfall: ObjectId stammt direkt aus Schritt 1 (ProvisionAsync) und steht in der Konfig.
        if (!string.IsNullOrWhiteSpace(_config.SpObjectId))
            return _config.SpObjectId;

        // Fallback: Einrichtung mit bereits vorhandener App (ohne Schritt 1) – dann manuell eingeben.
        using var dlg = new InputDialog("ServicePrincipal ObjectId (ServiceId; nur nötig, wenn Schritt 1 übersprungen):");
        if (dlg.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.Value))
        {
            _config.SpObjectId = dlg.Value.Trim();
            return _config.SpObjectId;
        }
        return null;
    }

    private void CommitGridEdits()
    {
        _grid.EndEdit();
        // DataGridView ist an _config.Mailboxes gebunden -> Änderungen sind bereits drin.
    }

    private void Rebind()
    {
        _grid.DataSource = null;
        _grid.DataSource = new BindingSource { DataSource = _config.Mailboxes };
    }
}

/// <summary>Kleiner Eingabedialog für Einzelwerte (ObjectId, Gruppen-Id …).</summary>
public sealed class InputDialog : Form
{
    private readonly TextBox _box = new() { Dock = DockStyle.Top, Width = 380 };
    public string Value => _box.Text;

    public InputDialog(string prompt, bool password = false)
    {
        if (password) _box.UseSystemPasswordChar = true;
        Text = "IMAP2M365";
        Width = 420; Height = 150; StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false;

        var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Dock = DockStyle.Right, Width = 80 };
        var cancel = new Button { Text = "Abbrechen", DialogResult = DialogResult.Cancel, Dock = DockStyle.Right, Width = 80 };
        var buttons = new Panel { Dock = DockStyle.Bottom, Height = 40 };
        buttons.Controls.Add(ok); buttons.Controls.Add(cancel);

        Controls.Add(_box);
        Controls.Add(new Label { Text = prompt, Dock = DockStyle.Top, Height = 24 });
        Controls.Add(buttons);
        AcceptButton = ok; CancelButton = cancel;

        // Nach einem Browser-Login kann ein modaler Dialog sonst hinter dem Hauptfenster landen.
        TopMost = true;
        Shown += (_, _) => { Activate(); BringToFront(); };
    }
}

/// <summary>Auswahldialog aus einer Liste (z. B. App-Registrierungen).</summary>
public sealed class ChooserDialog : Form
{
    private readonly ListBox _list = new() { Dock = DockStyle.Fill };

    private ChooserDialog(string prompt, IReadOnlyList<string> items)
    {
        Text = "IMAP2M365";
        Width = 560; Height = 380; StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable; MinimizeBox = false;

        foreach (var it in items) _list.Items.Add(it);
        if (_list.Items.Count > 0) _list.SelectedIndex = 0;

        var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Dock = DockStyle.Right, Width = 90 };
        var cancel = new Button { Text = "Abbrechen", DialogResult = DialogResult.Cancel, Dock = DockStyle.Right, Width = 90 };
        var buttons = new Panel { Dock = DockStyle.Bottom, Height = 40 };
        buttons.Controls.Add(ok); buttons.Controls.Add(cancel);

        Controls.Add(_list);
        Controls.Add(new Label { Text = prompt, Dock = DockStyle.Top, Height = 24 });
        Controls.Add(buttons);
        AcceptButton = ok; CancelButton = cancel;
    }

    /// <summary>Zeigt den Dialog im Vordergrund; gibt den gewählten Index zurück oder -1 bei Abbruch.</summary>
    public static int Pick(string prompt, IReadOnlyList<string> items)
    {
        using var dlg = new ChooserDialog(prompt, items);
        var owner = Form.ActiveForm; // sonst kann der modale Dialog hinter dem Hauptfenster landen
        dlg.StartPosition = owner is not null ? FormStartPosition.CenterParent : FormStartPosition.CenterScreen;
        dlg.TopMost = true;
        dlg.Shown += (_, _) => { dlg.Activate(); dlg.BringToFront(); };
        return dlg.ShowDialog(owner) == DialogResult.OK ? dlg._list.SelectedIndex : -1;
    }
}
