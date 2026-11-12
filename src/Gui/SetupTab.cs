using IMAP2M365.Ews;
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

    // Quelle: IMAP-Server oder Exchange on-prem (EWS mit ApplicationImpersonation)
    private readonly ComboBox _sourceKind = new() { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly GroupBox _onPremBox = new()
    {
        Text = "Quelle: Exchange on-prem (EWS, ein Migrationskonto mit ApplicationImpersonation)",
        Dock = DockStyle.Top, Height = 150, Visible = false
    };
    private readonly TextBox _opEwsUrl = new() { Width = 300, PlaceholderText = "https://exchange01.firma.local/EWS/Exchange.asmx" };
    private readonly TextBox _opPsUri = new() { Width = 260, PlaceholderText = "http://exchange01.firma.local/PowerShell/" };
    private readonly ComboBox _opPsAuth = new() { Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly TextBox _opUser = new() { Width = 300, PlaceholderText = "FIRMA\\migadmin" };
    private readonly TextBox _opPassword = new() { Width = 260, UseSystemPasswordChar = true };
    private readonly CheckBox _opAnyCert = new() { Text = "Zertifikat nicht prüfen (interne CA)", AutoSize = true };
    private readonly Button _opCheck = StepButton("A. Voraussetzungen prüfen");
    private readonly Button _opMatch = StepButton("B. Postfächer abgleichen");
    private readonly Button _opGrant = StepButton("C. Rechte + Throttling setzen");
    private readonly Button _opRevoke = StepButton("D. Rechte wieder entfernen");

    private OnPremRunner? _onPrem;
    private string _onPremKey = "";
    private bool _loading; // true, während Felder aus der Konfiguration befüllt werden

    private bool OnPremSelected => _sourceKind.SelectedIndex == 1;

    public SetupTab(AppConfig config, Action<string> log)
    {
        _config = config;
        _log = log;
        BuildUi();
        BuildGridColumns();
        WireEvents();
        ApplySourceKind();
        Disposed += (_, _) => _onPrem?.Dispose();
    }

    private void BuildUi()
    {
        var top = new TableLayoutPanel { Dock = DockStyle.Top, Height = 150, ColumnCount = 6, Padding = new Padding(8) };

        top.Controls.Add(new Label { Text = "Tenant (…onmicrosoft.com):", AutoSize = true }, 0, 0);
        top.Controls.Add(_tenant, 1, 0);
        top.Controls.Add(_includeShared, 2, 0);
        top.SetColumnSpan(_includeShared, 2);
        _sourceKind.Items.AddRange(new object[] { "IMAP-Server", "Exchange on-prem (EWS)" });
        _sourceKind.SelectedIndex = 0;
        top.Controls.Add(new Label { Text = "Quelle:", AutoSize = true }, 4, 0);
        top.Controls.Add(_sourceKind, 5, 0);

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

        // Exchange on-prem als Quelle: Verbindungsdaten + die vier Schritte A–D.
        _opPsAuth.Items.AddRange(new object[] { "Kerberos", "Negotiate", "Basic" });
        _opPsAuth.SelectedIndex = 0;
        var opGrid = new TableLayoutPanel { Dock = DockStyle.Top, Height = 64, ColumnCount = 6, Padding = new Padding(6, 2, 6, 0) };
        opGrid.Controls.Add(new Label { Text = "EWS-URL oder Server:", AutoSize = true }, 0, 0);
        opGrid.Controls.Add(_opEwsUrl, 1, 0);
        opGrid.Controls.Add(new Label { Text = "PowerShell-URI:", AutoSize = true }, 2, 0);
        opGrid.Controls.Add(_opPsUri, 3, 0);
        opGrid.Controls.Add(new Label { Text = "Anmeldung:", AutoSize = true }, 4, 0);
        opGrid.Controls.Add(_opPsAuth, 5, 0);
        opGrid.Controls.Add(new Label { Text = "Migrationskonto:", AutoSize = true }, 0, 1);
        opGrid.Controls.Add(_opUser, 1, 1);
        opGrid.Controls.Add(new Label { Text = "Kennwort:", AutoSize = true }, 2, 1);
        opGrid.Controls.Add(_opPassword, 3, 1);
        opGrid.Controls.Add(_opAnyCert, 4, 1);
        opGrid.SetColumnSpan(_opAnyCert, 2);

        var opButtons = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 52, WrapContents = false, Padding = new Padding(4, 4, 4, 4) };
        opButtons.Controls.AddRange(new Control[] { _opCheck, _opMatch, _opGrant, _opRevoke });
        _onPremBox.Controls.Add(opButtons);
        _onPremBox.Controls.Add(opGrid);

        Controls.Add(_grid);
        Controls.Add(buttons);
        Controls.Add(_onPremBox);
        Controls.Add(top);
    }

    /// <summary>Blendet je nach Quelle die passenden Felder/Spalten ein.</summary>
    private void ApplySourceKind()
    {
        var onPrem = OnPremSelected;
        _onPremBox.Visible = onPrem;
        _srcHost.Enabled = _srcPort.Enabled = _srcSsl.Enabled = !onPrem;

        foreach (DataGridViewColumn c in _grid.Columns)
        {
            if (c.DataPropertyName is "SourcePassword" or "SourceHostOverride") c.Visible = !onPrem;
            if (c.DataPropertyName == "SourceUser") c.HeaderText = onPrem ? "Quell-Postfach (on-prem)" : "Quell-User";
        }
        if (!_loading) _config.SourceKind = onPrem ? "OnPremEws" : "Imap";
    }

    /// <summary>Hält die Verbindungsfelder laufend in der Konfiguration (die Migration liest von dort).</summary>
    private void SyncSourceFieldsToConfig()
    {
        if (_loading) return;
        _config.SourceHost = _srcHost.Text.Trim();
        _config.SourcePort = (int)_srcPort.Value;
        _config.SourceUseSsl = _srcSsl.Checked;
        _config.OnPremEwsUrl = EwsUrl();
        _config.OnPremPsUri = _opPsUri.Text.Trim();
        _config.OnPremPsAuth = _opPsAuth.Text;
        _config.OnPremAdminUser = _opUser.Text.Trim();
        _config.OnPremAdminPassword = _opPassword.Text;
        _config.OnPremAcceptAnyCert = _opAnyCert.Checked;
    }

    /// <summary>EWS-Adresse aus dem Feld; ein bloßer Servername wird zur Standard-URL ergänzt.</summary>
    private string EwsUrl()
    {
        var t = _opEwsUrl.Text.Trim();
        if (t.Length == 0) return "";
        return t.StartsWith("http", StringComparison.OrdinalIgnoreCase) ? t : EwsClient.DefaultUrl(t);
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
        if (OnPremSelected) { await TestRowOnPremAsync(row); return; }

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

        _sourceKind.SelectedIndexChanged += (_, _) => ApplySourceKind();
        _opCheck.Click += async (_, _) => await CheckOnPremAsync();
        _opMatch.Click += async (_, _) => await MatchOnPremAsync();
        _opGrant.Click += async (_, _) => await GrantOnPremAsync();
        _opRevoke.Click += async (_, _) => await RevokeOnPremAsync();

        // Aus der EWS-Adresse die übliche PowerShell-URI vorschlagen.
        _opEwsUrl.Leave += (_, _) =>
        {
            if (_opPsUri.Text.Trim().Length == 0 && Uri.TryCreate(EwsUrl(), UriKind.Absolute, out var u))
                _opPsUri.Text = $"http://{u.Host}/PowerShell/";
        };

        foreach (var box in new[] { _srcHost, _opEwsUrl, _opPsUri, _opUser, _opPassword })
            box.TextChanged += (_, _) => SyncSourceFieldsToConfig();
        _srcPort.ValueChanged += (_, _) => SyncSourceFieldsToConfig();
        _srcSsl.CheckedChanged += (_, _) => SyncSourceFieldsToConfig();
        _opAnyCert.CheckedChanged += (_, _) => SyncSourceFieldsToConfig();
        _opPsAuth.SelectedIndexChanged += (_, _) => SyncSourceFieldsToConfig();
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
        _config.SourceKind = OnPremSelected ? "OnPremEws" : "Imap";
        SyncSourceFieldsToConfig();
        CommitGridEdits();
    }

    /// <summary>Lädt die Konfiguration in die Eingabefelder – nach Profil-Load.</summary>
    public void LoadFromConfig()
    {
        _loading = true;
        try
        {
            _tenant.Text = _config.TenantId;
            _srcHost.Text = _config.SourceHost;
            _srcPort.Value = Math.Clamp(_config.SourcePort, (int)_srcPort.Minimum, (int)_srcPort.Maximum);
            _srcSsl.Checked = _config.SourceUseSsl;
            _dryRun.Checked = _config.DryRun;

            _opEwsUrl.Text = _config.OnPremEwsUrl;
            _opPsUri.Text = _config.OnPremPsUri;
            _opPsAuth.SelectedIndex = Math.Max(0, _opPsAuth.Items.IndexOf(_config.OnPremPsAuth));
            _opUser.Text = _config.OnPremAdminUser;
            _opPassword.Text = _config.OnPremAdminPassword;
            _opAnyCert.Checked = _config.OnPremAcceptAnyCert;
            _sourceKind.SelectedIndex = _config.SourceIsOnPrem ? 1 : 0;

            ShowAppFields();
            Rebind();
            ApplySourceKind();
        }
        finally { _loading = false; }
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

    // ================= Exchange on-prem als Quelle =================

    private sealed record OnPremSettings(string EwsUrl, string PsUri, string PsAuth, string User, string Password, bool AnyCert);

    /// <summary>Liest die Verbindungsfelder (auf dem UI-Thread) und meldet, was noch fehlt.</summary>
    private OnPremSettings? ReadOnPremSettings(bool needPowerShell)
    {
        var s = new OnPremSettings(EwsUrl(), _opPsUri.Text.Trim(), _opPsAuth.Text, _opUser.Text.Trim(), _opPassword.Text, _opAnyCert.Checked);

        var missing = new List<string>();
        if (s.EwsUrl.Length == 0) missing.Add("EWS-URL oder Servername");
        if (needPowerShell && s.PsUri.Length == 0) missing.Add("PowerShell-URI");
        if (s.User.Length == 0) missing.Add("Migrationskonto");
        if (s.Password.Length == 0) missing.Add("Kennwort");
        if (missing.Count == 0) return s;

        _log("Exchange on-prem: Es fehlt noch " + string.Join(", ", missing) + ".");
        return null;
    }

    /// <summary>Stellt die Remote-PowerShell-Sitzung her bzw. verwendet die bestehende (Hintergrund-Thread).</summary>
    private OnPremRunner ConnectOnPrem(OnPremSettings s)
    {
        var key = $"{s.PsUri}|{s.User}|{s.PsAuth}|{s.AnyCert}";
        if (_onPrem is { IsConnected: true } && key == _onPremKey) return _onPrem;

        _onPrem?.Dispose();
        _onPrem = new OnPremRunner { Log = _log };
        _onPrem.Connect(s.PsUri, s.User, s.Password, s.PsAuth, s.AnyCert);
        _onPremKey = key;
        return _onPrem;
    }

    /// <summary>Führt eine On-Prem-Aktion im Hintergrund aus, damit das Fenster bedienbar bleibt.</summary>
    private async Task RunOnPremAsync(string title, Action work)
    {
        var buttons = new[] { _opCheck, _opMatch, _opGrant, _opRevoke };
        foreach (var b in buttons) b.Enabled = false;
        try { await Task.Run(work); }
        catch (Exception ex) { _log($"{title} fehlgeschlagen: {ex.Message}"); }
        finally { foreach (var b in buttons) b.Enabled = true; }
    }

    /// <summary>A. Prüft, ob alles für die Migration vom Exchange on-prem vorhanden ist, und listet auf, was fehlt.</summary>
    private async Task CheckOnPremAsync()
    {
        _grid.EndEdit();
        var s = ReadOnPremSettings(needPowerShell: true);
        if (s is null) return;

        var rows = _config.Mailboxes.Where(r => r.Enabled).ToList();
        var sample = rows.Select(r => r.SourceUser?.Trim()).FirstOrDefault(u => !string.IsNullOrEmpty(u));
        var withoutSource = rows.Count(r => string.IsNullOrWhiteSpace(r.SourceUser));
        var targetReady = _tenant.Text.Trim().Length > 0 && _clientId.Text.Trim().Length > 0 && _secret.Text.Length > 0;

        await RunOnPremAsync("Prüfung", () =>
        {
            var missing = new List<string>();
            _log("Prüfe Voraussetzungen (Quelle: Exchange on-prem) …");

            // 1) Remote-PowerShell, Rolle, Throttling
            OnPremStatus? status = null;
            try
            {
                var op = ConnectOnPrem(s);
                _log("  OK      Remote-PowerShell erreichbar.");
                status = op.GetStatus(s.User);
            }
            catch (Exception ex)
            {
                _log("  FEHLT   Remote-PowerShell: " + ex.Message);
                missing.Add("Remote-PowerShell-Verbindung (URI, Anmeldeart, Konto prüfen)");
            }

            if (status is not null)
            {
                if (status.HasImpersonation)
                    _log("  OK      Rolle ApplicationImpersonation vorhanden" + (status.OwnRoleAssignment ? " (von diesem Tool gesetzt)." : " (bestehende Zuweisung)."));
                else
                {
                    _log("  FEHLT   Rolle ApplicationImpersonation.");
                    missing.Add("Rolle ApplicationImpersonation (Schritt C)");
                }

                if (status.AssignedPolicy.Length > 0)
                    _log($"  OK      Throttling-Policy zugewiesen: {status.AssignedPolicy}.");
                else
                {
                    _log("  HINWEIS Keine Throttling-Policy zugewiesen – Exchange bremst EWS mit den Standardlimits.");
                    missing.Add("Throttling-Policy (Schritt C, empfohlen, nicht zwingend)");
                }
            }

            // 2) EWS-Zugriff und ob die Impersonation tatsächlich greift
            try
            {
                using var ews = new EwsClient(s.EwsUrl, s.User, s.Password, s.AnyCert) { ImpersonatedSmtp = sample };
                ews.GetWellKnownFolderIdsAsync(CancellationToken.None).GetAwaiter().GetResult();
                _log(sample is null
                    ? "  OK      EWS erreichbar."
                    : $"  OK      EWS erreichbar, Zugriff im Namen von {sample} funktioniert.");
            }
            catch (EwsException ex) when (ex.Code == "ErrorImpersonateUserDenied")
            {
                _log($"  FEHLT   EWS erreichbar, aber der Zugriff im Namen von {sample} wird verweigert (Rolle fehlt oder ist noch nicht aktiv).");
                if (!missing.Any(m => m.StartsWith("Rolle"))) missing.Add("Wirksame Impersonation (nach Schritt C einige Minuten warten)");
            }
            catch (EwsException ex) when (ex.Code == "Unauthorized" || ex.Code.StartsWith("Http", StringComparison.Ordinal))
            {
                _log("  FEHLT   EWS: " + ex.Message);
                missing.Add("EWS-Zugriff (URL, Konto, Kennwort prüfen)");
            }
            catch (EwsException ex) when (sample is null)
            {
                // Ohne zugeordnetes Postfach antwortet der Server für das Konto selbst – erreichbar ist er damit.
                _log($"  OK      EWS erreichbar und Anmeldung angenommen (Antwort: {ex.Code}).");
            }
            catch (Exception ex)
            {
                _log($"  FEHLT   EWS-Zugriff ({sample ?? s.EwsUrl}): {ex.Message}");
                missing.Add("EWS-Zugriff (Server erreichbar? URL und Zertifikat prüfen)");
            }

            // 3) Zuordnung und Ziel
            if (rows.Count == 0)
            {
                _log("  FEHLT   Keine Zielpostfächer geladen.");
                missing.Add("Zielpostfächer (Schritt 2 und 3)");
            }
            else if (withoutSource > 0)
            {
                _log($"  FEHLT   {withoutSource} von {rows.Count} Postfächern ohne Quell-Postfach.");
                missing.Add("Zuordnung der Quell-Postfächer (Schritt B oder von Hand)");
            }
            else _log($"  OK      {rows.Count} Postfächer mit Quell-Postfach zugeordnet.");

            if (targetReady) _log("  OK      App-Registrierung für das Ziel eingetragen.");
            else
            {
                _log("  FEHLT   App-Registrierung für das Ziel (Tenant, ClientId, Secret).");
                missing.Add("App-Registrierung (Schritt 1a oder 1b) und Einrichtung (Schritt 4)");
            }

            if (missing.Count == 0) _log("Ergebnis: Alles vorhanden – die Migration kann starten.");
            else
            {
                _log("Ergebnis: Es fehlt noch:");
                foreach (var m in missing) _log("  - " + m);
            }
        });
    }

    /// <summary>B. Liest die Postfächer vom Exchange on-prem und ordnet sie den Zielpostfächern zu.</summary>
    private async Task MatchOnPremAsync()
    {
        _grid.EndEdit();
        var s = ReadOnPremSettings(needPowerShell: true);
        if (s is null) return;
        if (_config.Mailboxes.Count == 0)
        {
            _log("Erst die Zielpostfächer laden (Schritt 2 und 3), dann abgleichen.");
            return;
        }

        var includeShared = _includeShared.Checked;
        List<OnPremMailbox>? list = null;
        await RunOnPremAsync("Abgleich", () => list = ConnectOnPrem(s).LoadMailboxes(includeShared));
        if (list is null) return;

        static string Local(string smtp) => smtp.Split('@')[0];

        var bySmtp = list.GroupBy(m => m.PrimarySmtp, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        var byLocal = list.ToLookup(m => Local(m.PrimarySmtp), StringComparer.OrdinalIgnoreCase);
        var byAlias = list.ToLookup(m => m.Alias, StringComparer.OrdinalIgnoreCase);

        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int matched = 0, open = 0;

        foreach (var row in _config.Mailboxes)
        {
            var local = Local(row.TargetSmtp);
            OnPremMailbox? hit = null;
            var how = "";

            if (bySmtp.TryGetValue((row.SourceUser ?? "").Trim(), out var current)) { hit = current; how = "bereits zugeordnet"; }
            else if (bySmtp.TryGetValue(row.TargetSmtp, out var same)) { hit = same; how = "gleiche Adresse"; }
            else if (byLocal[local].Count() == 1) { hit = byLocal[local].First(); how = "gleicher Name vor dem @"; }
            else if (byAlias[local].Count() == 1) { hit = byAlias[local].First(); how = "Alias"; }

            if (hit is not null)
            {
                row.SourceUser = hit.PrimarySmtp;
                used.Add(hit.PrimarySmtp);
                matched++;
                _log($"  {row.TargetSmtp}  <-  {hit.PrimarySmtp}  ({how})");
            }
            else
            {
                open++;
                var ambiguous = byLocal[local].Count() > 1 || byAlias[local].Count() > 1;
                _log($"  {row.TargetSmtp}: " + (ambiguous
                    ? "mehrere mögliche On-Prem-Postfächer – bitte von Hand eintragen."
                    : "kein passendes On-Prem-Postfach gefunden – bitte von Hand eintragen."));
            }
        }

        var unused = list.Where(m => !used.Contains(m.PrimarySmtp)).Select(m => m.PrimarySmtp).ToList();
        _log($"Abgleich: {matched} zugeordnet, {open} offen, {unused.Count} On-Prem-Postfächer ohne Ziel.");
        if (unused.Count > 0)
            _log("  Ohne Ziel: " + string.Join(", ", unused.Take(20)) + (unused.Count > 20 ? " …" : ""));

        Rebind();
    }

    /// <summary>C. Gibt dem Migrationskonto ApplicationImpersonation und eine EWS-Throttling-Policy.</summary>
    private async Task GrantOnPremAsync()
    {
        var s = ReadOnPremSettings(needPowerShell: true);
        if (s is null) return;
        var dry = _dryRun.Checked;

        await RunOnPremAsync("Rechte setzen", () =>
        {
            var op = ConnectOnPrem(s);
            if (dry)
            {
                var st = op.GetStatus(s.User);
                _log("Dry-Run – es wird nichts geändert:");
                _log(st.HasImpersonation
                    ? "  Rolle ApplicationImpersonation ist bereits vorhanden."
                    : $"  Rolle ApplicationImpersonation würde zugewiesen (Zuweisung '{OnPremRunner.RoleAssignmentName}').");
                _log(st.AssignedPolicy.Length > 0
                    ? $"  Throttling: Konto hat bereits die Policy '{st.AssignedPolicy}'."
                    : $"  Throttling-Policy '{OnPremRunner.ThrottlingPolicyName}' würde angelegt und zugewiesen.");
                return;
            }

            op.GrantImpersonation(s.User);
            try { op.ApplyThrottling(s.User); }
            catch (Exception ex)
            {
                _log("Throttling konnte nicht gesetzt werden (die Migration läuft trotzdem, nur langsamer): " + ex.Message);
            }
            _log("Fertig. Exchange übernimmt neue Rollen oft erst nach einigen Minuten – danach mit Schritt A prüfen.");
        });
    }

    /// <summary>D. Entfernt nach der Migration wieder, was Schritt C angelegt hat.</summary>
    private async Task RevokeOnPremAsync()
    {
        var s = ReadOnPremSettings(needPowerShell: true);
        if (s is null) return;
        var dry = _dryRun.Checked;

        if (!dry && MessageBox.Show(
                "Die von diesem Tool gesetzte Rolle ApplicationImpersonation und die Throttling-Policy werden " +
                "wieder entfernt. Danach ist keine Migration vom Exchange on-prem mehr möglich. Fortfahren?",
                "Rechte entfernen", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK)
            return;

        await RunOnPremAsync("Rechte entfernen", () =>
        {
            var op = ConnectOnPrem(s);
            if (dry)
            {
                var st = op.GetStatus(s.User);
                _log("Dry-Run – es wird nichts geändert:");
                _log(st.OwnRoleAssignment
                    ? $"  Rollenzuweisung '{OnPremRunner.RoleAssignmentName}' würde entfernt."
                    : "  Keine eigene Rollenzuweisung vorhanden.");
                _log(st.OwnPolicyExists
                    ? $"  Throttling-Policy '{OnPremRunner.ThrottlingPolicyName}' würde entfernt."
                    : "  Keine eigene Throttling-Policy vorhanden.");
                return;
            }

            op.RevokeAll(s.User);
            _log("Rückbau abgeschlossen.");
        });
    }

    /// <summary>Zeilen-Test bei On-Prem-Quelle: greift per EWS im Namen des Quell-Postfachs zu.</summary>
    private async Task TestRowOnPremAsync(MailboxRow row)
    {
        var s = ReadOnPremSettings(needPowerShell: false);
        if (s is null) return;
        if (string.IsNullOrWhiteSpace(row.SourceUser))
        {
            _log($"Test {row.TargetSmtp}: kein Quell-Postfach eingetragen (Schritt B).");
            return;
        }

        var mailbox = row.SourceUser.Trim();
        _log($"Test {row.TargetSmtp}: EWS-Zugriff auf {mailbox} …");
        try
        {
            await using var src = new EwsSource(s.EwsUrl, s.User, s.Password, s.AnyCert, mailbox);
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            await src.ConnectAsync(cts.Token);
            var folders = await src.ListFoldersAsync(cts.Token);
            var messages = 0;
            foreach (var f in folders) messages += await src.GetMessageCountAsync(f, cts.Token);
            _log($"Test {row.TargetSmtp}: OK – {folders.Count} Mailordner, {messages} Nachrichten.");
        }
        catch (EwsException ex) when (ex.Code == "ErrorImpersonateUserDenied")
        {
            _log($"Test {row.TargetSmtp}: FEHLER – Zugriff im Namen von {mailbox} verweigert. Rolle setzen (Schritt C) und einige Minuten warten.");
        }
        catch (Exception ex)
        {
            _log($"Test {row.TargetSmtp}: FEHLER – {ex.Message}");
        }
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
