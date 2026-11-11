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
    private readonly ProgressBar _bar = new() { Dock = DockStyle.Top, Height = 22 };
    private readonly Label _state = new() { Dock = DockStyle.Top, Height = 22, Text = "Bereit." };

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

        Controls.Add(_state);
        Controls.Add(_bar);
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
        MigrationProgress? last = null;

        void UpdateStatus()
        {
            var copiedSum = perMailboxCopied.Values.Sum();
            var failed = Interlocked.Read(ref totalFailed);
            var active = Volatile.Read(ref activeJobs);
            var done = Volatile.Read(ref completed);
            var p = last;
            var detail = p is null ? "" : $" · akt.: {p.Mailbox} {p.FolderIndex}/{p.FolderTotal} · Nachricht {p.MessageIndex}/{p.MessageTotal}";
            var text = $"Jobs aktiv: {active}/{degree} · Postfächer {done}/{activeTotal} · {copiedSum} kopiert" +
                       (failed > 0 ? $", {failed} fehlerhaft" : "") + detail;
            RunOnUi(() =>
            {
                _state.Text = text;
                _bar.Style = ProgressBarStyle.Continuous;
                _bar.Maximum = Math.Max(1, activeTotal);
                _bar.Value = Math.Min(done, activeTotal);
            });
        }

        // Fortschritt je Nachricht: kopiert-Zähler des Postfachs + letzte Detailposition aktualisieren.
        var progress = new Progress<MigrationProgress>(p =>
        {
            perMailboxCopied[p.Mailbox] = p.MessagesCopied;
            last = p;
            UpdateStatus();
        });

        async Task MigrateOneAsync(MailboxRow row, int index, CancellationToken ct)
        {
            // Lizenz-Vorabprüfung (lose); endgültig unter Lock nach Connect.
            if (!testMode)
            {
                bool quotaLeft;
                lock (_licLock) quotaLeft = _license.IsConsumed(row.TargetSmtp) || _license.Remaining > 0;
                if (!quotaLeft) { _log($"{row.TargetSmtp}: übersprungen – keine freie Lizenz mehr."); return; }
            }

            var host = (row.SourceHostOverride ?? cfg.SourceHost ?? "").Trim();
            var port = row.SourcePortOverride ?? cfg.SourcePort;
            if (string.IsNullOrWhiteSpace(host)) { _log($"{row.TargetSmtp}: kein Quell-IMAP-Server gesetzt – übersprungen."); return; }
            if (string.IsNullOrWhiteSpace(row.SourceUser)) { _log($"{row.TargetSmtp}: kein Quell-User gesetzt – übersprungen."); return; }

            await using var source = new ImapSource(host, port, cfg.SourceUseSsl, row.SourceUser, row.SourcePassword);
            _log($"[{row.TargetSmtp}] Verbinde Quelle {host}:{port} (SSL={cfg.SourceUseSsl}) als {row.SourceUser} …");
            try { await source.ConnectAsync(WithTimeout(ct, 30)); }
            catch (Exception ex) when (!ct.IsCancellationRequested)
            {
                _log($"{row.TargetSmtp}: Quell-IMAP {host}:{port} nicht erreichbar oder Login fehlgeschlagen – {ex.Message}");
                return;
            }

            IMailTarget target = useGraph
                ? new GraphTarget(cfg.TenantId, cfg.ClientId, cfg.ClientSecret)
                : new ImapAppendTarget(tokens);
            await using var _ = target;
            _log($"[{row.TargetSmtp}] Verbinde Ziel (Exchange Online) …");
            try { await target.ConnectAsync(row.TargetSmtp, WithTimeout(ct, 30)); }
            catch (Exception ex) when (!ct.IsCancellationRequested)
            {
                _log($"{row.TargetSmtp}: Ziel (Exchange Online) nicht erreichbar oder Auth fehlgeschlagen – {ex.Message}");
                _log("  Prüfen: Einrichtung OHNE Dry-Run ausgeführt? IMAP am Postfach aktiv, ServicePrincipal registriert, FullAccess erteilt? (Änderungen können einige Minuten brauchen.)");
                return;
            }

            // Lizenz endgültig und atomar ziehen (verhindert Doppelvergabe bei paralleler Ausführung).
            if (!testMode)
            {
                bool got;
                lock (_licLock) got = _license.TryConsume(row.TargetSmtp);
                RunOnUi(() => _onLicenseChanged());
                if (!got) { _log($"{row.TargetSmtp}: übersprungen – keine freie Lizenz mehr."); return; }
            }

            var engine = new MigrationEngine(cfg, mapFolder: MapFolder, progress: progress, log: _log);
            var (copied, skipped) = await engine.RunAsync(row.TargetSmtp, source, target, ct, index, activeTotal);
            perMailboxCopied[row.TargetSmtp] = copied;
            Interlocked.Add(ref totalCopied, copied);
            Interlocked.Add(ref totalSkipped, skipped);
            _log($"{row.TargetSmtp}: {copied} kopiert, {skipped} übersprungen" + (testMode ? " (Testmodus)" : mirror ? " (Spiegel)" : ""));
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
}
