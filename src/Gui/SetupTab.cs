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

    private readonly Button _provision = new() { Text = "1. Entra-App anlegen…", Width = 170 };
    private readonly Button _connect = new() { Text = "2. Mit EXO verbinden", Width = 170 };
    private readonly Button _loadMbx = new() { Text = "3. Postfächer laden", Width = 170 };
    private readonly Button _apply = new() { Text = "4. Einrichtung ausführen", Width = 170 };

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
        var top = new TableLayoutPanel { Dock = DockStyle.Top, Height = 120, ColumnCount = 6, Padding = new Padding(8) };

        top.Controls.Add(new Label { Text = "Tenant (…onmicrosoft.com):", AutoSize = true }, 0, 0);
        top.Controls.Add(_tenant, 1, 0);
        top.Controls.Add(_includeShared, 2, 0);
        top.SetColumnSpan(_includeShared, 2);

        top.Controls.Add(new Label { Text = "Quell-IMAP (zentral):", AutoSize = true }, 0, 1);
        top.Controls.Add(_srcHost, 1, 1);
        top.Controls.Add(_srcPort, 2, 1);
        top.Controls.Add(_srcSsl, 3, 1);
        top.Controls.Add(_dryRun, 4, 1);

        var buttons = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 44, Padding = new Padding(8, 4, 8, 4) };
        buttons.Controls.AddRange(new Control[] { _provision, _connect, _loadMbx, _apply });

        Controls.Add(_grid);
        Controls.Add(buttons);
        Controls.Add(top);
    }

    private void BuildGridColumns()
    {
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Ziel-Postfach", DataPropertyName = "TargetSmtp", ReadOnly = true });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Typ", DataPropertyName = "MailboxType", ReadOnly = true, FillWeight = 60 });
        _grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "IMAP aktiv", DataPropertyName = "ImapEnabled", FillWeight = 55 });
        _grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Recht ok", DataPropertyName = "PermissionGranted", ReadOnly = true, FillWeight = 50 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Quell-User", DataPropertyName = "SourceUser" });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Quell-Kennwort", DataPropertyName = "SourcePassword" });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Quell-Server (optional)", DataPropertyName = "SourceHostOverride" });

        // Kennwort-Spalte maskieren
        _grid.CellFormatting += (s, e) =>
        {
            if (_grid.Columns[e.ColumnIndex].DataPropertyName == "SourcePassword" && e.Value is string pw && pw.Length > 0)
                e.Value = new string('•', Math.Min(pw.Length, 10));
        };
    }

    private void WireEvents()
    {
        _provision.Click += async (_, _) => await ProvisionAsync();
        _connect.Click += (_, _) => Connect();
        _loadMbx.Click += (_, _) => LoadMailboxes();
        _apply.Click += (_, _) => ApplySetup();
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
            _log($"Fertig: ClientId {appId}. Secret und ServicePrincipal-ObjectId liegen in der Konfiguration.");
            _log("Hinweis: ServicePrincipal in EXO wird in Schritt 4 registriert.");
        }
        catch (Exception ex) { _log("Provisionierung fehlgeschlagen: " + ex.Message); }
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

    private void ApplySetup()
    {
        if (_exo is null) { _log("Erst mit EXO verbinden (Schritt 2)."); return; }

        _config.SourceHost = _srcHost.Text.Trim();
        _config.SourcePort = (int)_srcPort.Value;
        _config.SourceUseSsl = _srcSsl.Checked;
        _config.DryRun = _dryRun.Checked;
        CommitGridEdits();

        var dry = _dryRun.Checked;
        _log(dry ? "Dry-Run: zeige nur, was geändert würde." : "Führe Einrichtung aus.");

        // Service-Principal einmalig registrieren
        var spObjectId = PromptMissingSpObjectId();
        if (spObjectId is null) return;

        if (!dry) _exo.RegisterServicePrincipal(_config.ClientId, spObjectId);

        foreach (var row in _config.Mailboxes)
        {
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
        _log("Einrichtung abgeschlossen.");
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

    public InputDialog(string prompt)
    {
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
    }
}
