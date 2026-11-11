using IMAP2M365.Core;
using IMAP2M365.Licensing;
using IMAP2M365.Sources;
using IMAP2M365.Targets;

namespace IMAP2M365.Gui;

/// <summary>
/// Fährt die eigentliche Migration: pro Postfachzeile Quell-IMAP -> EXO (XOAUTH2-APPEND),
/// mit Fortschrittsbalken. Testmodus kopiert nur die ersten N Mails/Postfach (keine Lizenz),
/// Vollmigration verbraucht pro Postfach eine Lizenz.
/// </summary>
public sealed class MigrationTab : UserControl
{
    private const int TestMessageLimit = 10;

    private readonly AppConfig _config;
    private readonly Action<string> _log;
    private readonly LicenseManager _license;
    private readonly Action _onLicenseChanged;

    private static Button ActionButton(string text, bool enabled = true) => new()
    {
        Text = text,
        Enabled = enabled,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        MinimumSize = new Size(0, 30),
        Padding = new Padding(12, 4, 12, 4),
        Margin = new Padding(4, 4, 4, 4),
        Anchor = AnchorStyles.Left
    };

    private readonly Button _start = ActionButton("Migration starten");
    private readonly Button _cancel = ActionButton("Abbrechen", enabled: false);
    private readonly Label _modeLabel = new() { Text = "Modus:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 4, 2, 4) };
    private readonly ComboBox _mode = new() { Width = 260, DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left, Margin = new Padding(4) };
    private readonly CheckBox _debug = new() { Text = "Debug (Betreffs zeigen)", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 4, 4, 4) };
    private readonly Label _parallelLabel = new() { Text = "Parallel:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 4, 2, 4) };
    private readonly NumericUpDown _parallel = new() { Minimum = 1, Maximum = 4, Value = 1, Width = 50, Anchor = AnchorStyles.Left, Margin = new Padding(2, 4, 4, 4) };
    private readonly Label _backendLabel = new() { Text = "Ziel-Backend:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(16, 4, 2, 4) };
    private readonly ComboBox _backend = new() { Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left, Margin = new Padding(4) };

    private readonly object _licLock = new();
    private readonly Label _state = new() { Dock = DockStyle.Top, Height = 22, Text = "Bereit." };
    private readonly Label[] _jobLabels = new Label[4];
    private readonly ProgressBar[] _jobBars = new ProgressBar[4];
    private readonly System.Collections.Concurrent.ConcurrentStack<int> _freeSlots = new();
    private readonly string?[] _slotMailbox = new string?[4];
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, System.Text.StringBuilder> _jobLogs = new();

    private const int ModeTest = 0, ModeInitial = 1, ModeMirror = 2;

    private CancellationTokenSource? _cts;

    public MigrationTab(AppConfig config, Action<string> log, LicenseManager license, Action onLicenseChanged)
    {
        _config = config;
        _log = log;
        _license = license;
        _onLicenseChanged = onLicenseChanged;

        _mode.Items.AddRange(new object[] { $"Testmodus (erste {TestMessageLimit} Mails)", "Erstmigration (voll)", "Spiegelabgleich" });
        _mode.SelectedIndex = ModeTest;
        _backend.Items.AddRange(new object[] { "IMAP-APPEND (XOAUTH2)", "Microsoft Graph" });
        _backend.SelectedIndex = _config.TargetKind.Equals("Graph", StringComparison.OrdinalIgnoreCase) ? 1 : 0;

        // TableLayoutPanel richtet alle Elemente (Buttons/Combos/Checkbox) vertikal zentriert aus.
        var buttons = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Height = 46, ColumnCount = 9, RowCount = 1, Padding = new Padding(8, 4, 8, 4) };
        for (int i = 0; i < 9; i++) buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttons.Controls.Add(_start, 0, 0);
        buttons.Controls.Add(_cancel, 1, 0);
        buttons.Controls.Add(_modeLabel, 2, 0);
        buttons.Controls.Add(_mode, 3, 0);
        buttons.Controls.Add(_parallelLabel, 4, 0);
        buttons.Controls.Add(_parallel, 5, 0);
        buttons.Controls.Add(_debug, 6, 0);
        buttons.Controls.Add(_backendLabel, 7, 0);
        buttons.Controls.Add(_backend, 8, 0);

        // Pro Job eine eigene Zeile: Label (Postfach + Position) und Fortschrittsbalken.
        var jobsBox = new GroupBox { Text = "Jobs", Dock = DockStyle.Top, Height = 132 };
        var jobGrid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 4, Padding = new Padding(6, 2, 6, 2) };
        jobGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
        jobGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
        for (int r = 0; r < 4; r++)
        {
            jobGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 25));
            var slot = r;
            _jobLabels[r] = new Label { Dock = DockStyle.Fill, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Text = "— (frei)", Cursor = Cursors.Hand };
            _jobBars[r] = new ProgressBar { Dock = DockStyle.Fill, Style = ProgressBarStyle.Continuous, Margin = new Padding(4, 4, 4, 4), Cursor = Cursors.Hand };
            _jobLabels[r].Click += (_, _) => OpenJobLog(slot);
            _jobBars[r].Click += (_, _) => OpenJobLog(slot);
            jobGrid.Controls.Add(_jobLabels[r], 0, r);
            jobGrid.Controls.Add(_jobBars[r], 1, r);
        }
        jobsBox.Controls.Add(jobGrid);

        Controls.Add(jobsBox);
        Controls.Add(_state);
        Controls.Add(buttons);

        _start.Click += async (_, _) => await RunAsync();
        _cancel.Click += (_, _) => _cts?.Cancel();
    }

    /// <summary>Schreibt Backend-/Dry-Run-Auswahl in die Konfiguration – für Speichern.</summary>
    public void FlushToConfig()
    {
        _config.TargetKind = _backend.SelectedIndex == 1 ? "Graph" : "ImapAppend";
    }

    /// <summary>Lädt die Konfiguration in die Bedienelemente – nach Profil-Load.</summary>
    public void LoadFromConfig()
    {
        _backend.SelectedIndex = _config.TargetKind.Equals("Graph", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
    }

    private async Task RunAsync()
    {
        if (_config.Mailboxes.Count == 0) { _log("Keine Postfächer - erst im Reiter Einrichtung laden."); return; }
        if (string.IsNullOrWhiteSpace(_config.ClientId)) { _log("Keine ClientId – erst Einrichtung ausführen."); return; }

        var mode = _mode.SelectedIndex;
        var testMode = mode == ModeTest;
        var mirror = mode == ModeMirror;

        // Testmodus ist frei. Erstmigration/Spiegel brauchen eine gültige Lizenz.
        if (!testMode)
        {
            if (!_license.HasValidLicense)
            {
                _log("Keine gültige Lizenz - bitte im Reiter Lizenz aktivieren. (Testmodus bleibt frei nutzbar.)");
                return;
            }
            var needed = _config.Mailboxes.Count(m => m.Enabled && !_license.IsConsumed(m.TargetSmtp));
            _log($"Lizenz: {_license.Remaining} frei, {needed} neue aktive Postfächer in dieser Migration.");
        }

        // Spiegelabgleich löscht im Ziel -> ausdrücklich bestätigen.
        if (mirror)
        {
            var ok = MessageBox.Show(
                "Spiegelabgleich gleicht das Ziel an die Quelle an und VERSCHIEBT im Ziel Nachrichten in den " +
                "Papierkorb, die in der Quelle nicht mehr am selben Ort sind (verschoben/gelöscht). Fortfahren?",
                "Spiegelabgleich", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (ok != DialogResult.OK) return;
        }

        _cts = new CancellationTokenSource();
        _start.Enabled = false; _cancel.Enabled = true;
        var useGraph = _backend.SelectedIndex == 1;

        var cfg = new MigrationConfig
        {
            TenantId = _config.TenantId,
            ClientId = _config.ClientId,
            ClientSecret = _config.ClientSecret,
            SourceHost = _config.SourceHost,
            SourcePort = _config.SourcePort,
            SourceUseSsl = _config.SourceUseSsl,
            TestMessageLimit = testMode ? TestMessageLimit : 0,
            Mirror = mirror,
            Debug = _debug.Checked,
            ExcludedFolders = _config.ExcludedFolders,
            Target = useGraph ? TargetKind.Graph : TargetKind.ImapAppend
        };

        // Token-Provider nur für das IMAP-Backend nötig; GraphTarget baut eigene Credentials.
        var tokens = new ExchangeOnlineTokenProvider(cfg.TenantId, cfg.ClientId, cfg.ClientSecret);
        var modeName = testMode ? "Testmodus" : mirror ? "Spiegelabgleich" : "Erstmigration";
        var degree = (int)_parallel.Value;
        _log($"Ziel-Backend: {(useGraph ? "Microsoft Graph" : "IMAP-APPEND (XOAUTH2)")} · {modeName} · {degree} Job(s) parallel");

        var active = _config.Mailboxes.Where(m => m.Enabled).ToList();
        var activeTotal = active.Count;

        int activeJobs = 0, completed = 0;
        long totalCopied = 0, totalSkipped = 0, totalFailed = 0;
        var perMailboxCopied = new System.Collections.Concurrent.ConcurrentDictionary<string, long>();

        // Logger je Postfach: schreibt in den dedizierten Job-Puffer UND ins globale Protokoll.
        Action<string> Logger(string mailbox)
        {
            var sb = _jobLogs.GetOrAdd(mailbox, _ => new System.Text.StringBuilder());
            return line =>
            {
                lock (sb) sb.AppendLine($"{DateTime.Now:HH:mm:ss}  {line}");
                _log(line);
            };
        }

        void UpdateStatus()
        {
            var copiedSum = perMailboxCopied.Values.Sum();
            var failed = Interlocked.Read(ref totalFailed);
            var active = Volatile.Read(ref activeJobs);
            var done = Volatile.Read(ref completed);
            var text = $"Jobs aktiv: {active}/{degree} · Postfächer {done}/{activeTotal} · {copiedSum} kopiert" +
                       (failed > 0 ? $", {failed} fehlerhaft" : "");
            RunOnUi(() => _state.Text = text);
        }

        // Alle Job-Zeilen zurücksetzen und die freien Slots 0..3 bereitstellen (LIFO -> 0 zuerst).
        _jobLogs.Clear();
        for (int r = 0; r < 4; r++) _slotMailbox[r] = null;
        _freeSlots.Clear();
        for (int r = 3; r >= 0; r--) _freeSlots.Push(r);
        RunOnUi(() =>
        {
            for (int r = 0; r < 4; r++) { _jobLabels[r].Text = "— (frei)"; _jobBars[r].Value = 0; }
        });

        void SetSlot(int slot, string text) => RunOnUi(() => _jobLabels[slot].Text = text);
        void ClearSlot(int slot) => RunOnUi(() =>
        {
            var mb = _slotMailbox[slot];
            _jobLabels[slot].Text = mb is null ? "— (frei)" : $"✓ {mb} (fertig – Log ansehen)";
            if (mb is null) _jobBars[slot].Value = 0;
        });
        void UpdateSlot(int slot, MigrationProgress p) => RunOnUi(() =>
        {
            _jobLabels[slot].Text = $"{p.Mailbox} · Ordner {p.FolderIndex}/{p.FolderTotal} · Nachricht {p.MessageIndex}/{p.MessageTotal} · {p.MessagesCopied} kopiert";
            var bar = _jobBars[slot];
            bar.Maximum = Math.Max(1, p.MessageTotal);
            bar.Value = Math.Clamp(p.MessageIndex, 0, bar.Maximum);
        });

        async Task MigrateOneAsync(MailboxRow row, int index, CancellationToken ct)
        {
            if (!_freeSlots.TryPop(out var slot)) slot = 0; // sollte bei degree<=4 nicht passieren
            SetSlot(slot, $"{row.TargetSmtp} · verbinde …");
            try
            {
                await MigrateCoreAsync(row, index, slot, ct);
            }
            finally
            {
                ClearSlot(slot);
                _freeSlots.Push(slot);
            }
        }

        async Task MigrateCoreAsync(MailboxRow row, int index, int slot, CancellationToken ct)
        {
            _slotMailbox[slot] = row.TargetSmtp;
            var jlog = Logger(row.TargetSmtp);

            // Eigener Fortschritt pro Job -> eigener Balken/Label in dieser Slot-Zeile.
            var jobProgress = new Progress<MigrationProgress>(p =>
            {
                perMailboxCopied[p.Mailbox] = p.MessagesCopied;
                UpdateSlot(slot, p);
                UpdateStatus();
            });

            // Lizenz-Vorabprüfung (lose); endgültig unter Lock nach Connect.
            if (!testMode)
            {
                bool quotaLeft;
                lock (_licLock) quotaLeft = _license.IsConsumed(row.TargetSmtp) || _license.Remaining > 0;
                if (!quotaLeft) { jlog($"{row.TargetSmtp}: übersprungen – keine freie Lizenz mehr."); return; }
            }

            var host = (row.SourceHostOverride ?? cfg.SourceHost ?? "").Trim();
            var port = row.SourcePortOverride ?? cfg.SourcePort;
            if (string.IsNullOrWhiteSpace(host)) { jlog($"{row.TargetSmtp}: kein Quell-IMAP-Server gesetzt – übersprungen."); return; }
            if (string.IsNullOrWhiteSpace(row.SourceUser)) { jlog($"{row.TargetSmtp}: kein Quell-User gesetzt – übersprungen."); return; }

            await using var source = new ImapSource(host, port, cfg.SourceUseSsl, row.SourceUser, row.SourcePassword);
            jlog($"Verbinde Quelle {host}:{port} (SSL={cfg.SourceUseSsl}) als {row.SourceUser} …");
            try { await source.ConnectAsync(WithTimeout(ct, 30)); }
            catch (Exception ex) when (!ct.IsCancellationRequested)
            {
                jlog($"{row.TargetSmtp}: Quell-IMAP {host}:{port} nicht erreichbar oder Login fehlgeschlagen – {ex.Message}");
                return;
            }

            IMailTarget target = useGraph
                ? new GraphTarget(cfg.TenantId, cfg.ClientId, cfg.ClientSecret)
                : new ImapAppendTarget(tokens);
            await using var _ = target;
            jlog("Verbinde Ziel (Exchange Online) …");
            try { await target.ConnectAsync(row.TargetSmtp, WithTimeout(ct, 30)); }
            catch (Exception ex) when (!ct.IsCancellationRequested)
            {
                jlog($"{row.TargetSmtp}: Ziel (Exchange Online) nicht erreichbar oder Auth fehlgeschlagen – {ex.Message}");
                jlog("  Prüfen: Einrichtung OHNE Dry-Run ausgeführt? IMAP am Postfach aktiv, ServicePrincipal registriert, FullAccess erteilt? (Änderungen können einige Minuten brauchen.)");
                return;
            }

            // Lizenz endgültig und atomar ziehen (verhindert Doppelvergabe bei paralleler Ausführung).
            if (!testMode)
            {
                bool got;
                lock (_licLock) got = _license.TryConsume(row.TargetSmtp);
                RunOnUi(() => _onLicenseChanged());
                if (!got) { jlog($"{row.TargetSmtp}: übersprungen – keine freie Lizenz mehr."); return; }
            }

            var engine = new MigrationEngine(cfg, mapFolder: MapFolder, progress: jobProgress, log: jlog);
            var (copied, skipped) = await engine.RunAsync(row.TargetSmtp, source, target, ct, index, activeTotal);
            perMailboxCopied[row.TargetSmtp] = copied;
            Interlocked.Add(ref totalCopied, copied);
            Interlocked.Add(ref totalSkipped, skipped);
            jlog($"{row.TargetSmtp}: {copied} kopiert, {skipped} übersprungen" + (testMode ? " (Testmodus)" : mirror ? " (Spiegel)" : ""));
        }

        _start.Enabled = false; _cancel.Enabled = true;
        try
        {
            var indexed = active.Select((row, idx) => (row, idx: idx + 1)).ToList();
            await Parallel.ForEachAsync(indexed,
                new ParallelOptions { MaxDegreeOfParallelism = degree, CancellationToken = _cts.Token },
                async (item, ct) =>
                {
                    Interlocked.Increment(ref activeJobs); UpdateStatus();
                    try { await MigrateOneAsync(item.row, item.idx, ct); }
                    catch (Exception ex) when (!_cts.IsCancellationRequested)
                    {
                        Interlocked.Increment(ref totalFailed);
                        _log($"{item.row.TargetSmtp}: Postfach-Fehler – {ex.Message}");
                    }
                    finally
                    {
                        Interlocked.Decrement(ref activeJobs);
                        Interlocked.Increment(ref completed);
                        UpdateStatus();
                    }
                });

            SetState($"Fertig. {Interlocked.Read(ref totalCopied)} kopiert, {Interlocked.Read(ref totalSkipped)} übersprungen" +
                     (Interlocked.Read(ref totalFailed) > 0 ? $", {Interlocked.Read(ref totalFailed)} fehlerhaft" : ""));
        }
        catch (OperationCanceledException) { SetState("Abgebrochen."); _log("Migration abgebrochen."); }
        catch (Exception ex) { SetState("Fehler."); _log("Migration fehlgeschlagen: " + ex.Message); }
        finally { _start.Enabled = true; _cancel.Enabled = false; _cts?.Dispose(); _cts = null; }
    }

    /// <summary>Verknüpft den Abbruch-Token mit einem Timeout, damit Verbindungsversuche nicht ewig hängen.</summary>
    private static CancellationToken WithTimeout(CancellationToken outer, int seconds)
    {
        var cts = CancellationTokenSource.CreateLinkedTokenSource(outer);
        cts.CancelAfter(TimeSpan.FromSeconds(seconds));
        return cts.Token;
    }

    /// <summary>
    /// Ordner-Mapping: Sonderordner der Quelle auf EXO-Wellknown-Namen abbilden,
    /// damit „Gesendet"/„Sent" nicht als Extra-Ordner neben „Sent Items" landet.
    /// </summary>
    private static IReadOnlyList<string> MapFolder(IReadOnlyList<string> source)
    {
        if (source.Count == 1)
        {
            var name = source[0];
            var mapped = name.ToLowerInvariant() switch
            {
                "inbox" or "posteingang" => "INBOX",
                "sent" or "sent items" or "gesendet" or "gesendete elemente" => "Sent Items",
                "drafts" or "entwürfe" => "Drafts",
                "deleted items" or "gelöschte elemente" or "papierkorb" => "Deleted Items",
                "junk" or "junk-e-mail" or "spam" => "Junk Email",
                _ => name
            };
            return new[] { mapped };
        }
        return source;
    }

    private void SetState(string text)
    {
        if (InvokeRequired) { BeginInvoke(() => SetState(text)); return; }
        _state.Text = text;
    }

    /// <summary>Führt eine UI-Aktion sicher auf dem UI-Thread aus (für Aufrufe aus parallelen Jobs).</summary>
    private void RunOnUi(Action action)
    {
        if (IsDisposed) return;
        if (InvokeRequired) BeginInvoke(action);
        else action();
    }

    /// <summary>Öffnet das dedizierte Log-Fenster für das Postfach dieser Job-Zeile.</summary>
    private void OpenJobLog(int slot)
    {
        var mb = _slotMailbox[slot];
        if (string.IsNullOrEmpty(mb)) return;
        var sb = _jobLogs.GetOrAdd(mb, _ => new System.Text.StringBuilder());
        new JobLogForm(mb, sb) { Owner = FindForm() }.Show();
    }
}

/// <summary>Live-Log-Fenster für ein einzelnes Postfach (aktualisiert sich selbst).</summary>
public sealed class JobLogForm : Form
{
    private readonly System.Text.StringBuilder _sb;
    private readonly TextBox _box = new()
    {
        Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false,
        Dock = DockStyle.Fill, Font = new Font("Consolas", 8.5f)
    };
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 600 };
    private int _lastLen = -1;

    public JobLogForm(string title, System.Text.StringBuilder sb)
    {
        _sb = sb;
        Text = $"Log: {title}";
        Width = 780; Height = 480; StartPosition = FormStartPosition.CenterParent;
        Controls.Add(_box);
        _timer.Tick += (_, _) => RefreshLog();
        _timer.Start();
        RefreshLog();
        FormClosed += (_, _) => _timer.Dispose();
    }

    private void RefreshLog()
    {
        string text;
        lock (_sb) text = _sb.ToString();
        if (text.Length == _lastLen) return;
        _lastLen = text.Length;

        var atBottom = _box.SelectionStart >= _box.TextLength - 2;
        _box.Text = text;
        if (atBottom) { _box.SelectionStart = _box.TextLength; _box.ScrollToCaret(); }
    }
}
