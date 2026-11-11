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
    private readonly CheckBox _testMode = new() { Text = $"Testmodus (erste {TestMessageLimit} Mails)", Checked = true, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 4, 4, 4) };
    private readonly Label _backendLabel = new() { Text = "Ziel-Backend:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(16, 4, 2, 4) };
    private readonly ComboBox _backend = new() { Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Left, Margin = new Padding(4) };
    private readonly ProgressBar _bar = new() { Dock = DockStyle.Top, Height = 22 };
    private readonly Label _state = new() { Dock = DockStyle.Top, Height = 22, Text = "Bereit." };

    private CancellationTokenSource? _cts;

    public MigrationTab(AppConfig config, Action<string> log, LicenseManager license, Action onLicenseChanged)
    {
        _config = config;
        _log = log;
        _license = license;
        _onLicenseChanged = onLicenseChanged;

        _backend.Items.AddRange(new object[] { "IMAP-APPEND (XOAUTH2)", "Microsoft Graph" });
        _backend.SelectedIndex = _config.TargetKind.Equals("Graph", StringComparison.OrdinalIgnoreCase) ? 1 : 0;

        // TableLayoutPanel richtet alle Elemente (Buttons/Checkbox/Label/Combo) vertikal zentriert aus.
        var buttons = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Height = 46, ColumnCount = 5, RowCount = 1, Padding = new Padding(8, 4, 8, 4) };
        for (int i = 0; i < 5; i++) buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttons.Controls.Add(_start, 0, 0);
        buttons.Controls.Add(_cancel, 1, 0);
        buttons.Controls.Add(_testMode, 2, 0);
        buttons.Controls.Add(_backendLabel, 3, 0);
        buttons.Controls.Add(_backend, 4, 0);

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

        var testMode = _testMode.Checked;

        // Lizenzprüfung nur für die echte Vollmigration; Testmodus ist frei.
        if (!testMode)
        {
            if (!_license.HasValidLicense)
            {
                _log("Keine gültige Lizenz - bitte im Reiter Lizenz aktivieren. (Testmodus bleibt nutzbar.)");
                return;
            }
            var needed = _config.Mailboxes.Count(m => m.Enabled && !_license.IsConsumed(m.TargetSmtp));
            _log($"Lizenz: {_license.Remaining} frei, {needed} neue aktive Postfächer in dieser Migration.");
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
            ExcludedFolders = _config.ExcludedFolders,
            Target = useGraph ? TargetKind.Graph : TargetKind.ImapAppend
        };

        // Token-Provider nur für das IMAP-Backend nötig; GraphTarget baut eigene Credentials.
        var tokens = new ExchangeOnlineTokenProvider(cfg.TenantId, cfg.ClientId, cfg.ClientSecret);
        _log($"Ziel-Backend: {(useGraph ? "Microsoft Graph" : "IMAP-APPEND (XOAUTH2)")}" + (testMode ? " · Testmodus" : " · Vollmigration"));

        try
        {
            int i = 0;
            foreach (var row in _config.Mailboxes)
            {
                _cts.Token.ThrowIfCancellationRequested();

                if (!row.Enabled) { _log($"{row.TargetSmtp}: deaktiviert – übersprungen."); continue; }

                SetState($"Postfach {++i}/{_config.Mailboxes.Count}: {row.TargetSmtp}");

                // Vollmigration: früh prüfen, ob noch Kontingent frei ist (verbraucht wird erst nach Connect).
                if (!testMode && !_license.IsConsumed(row.TargetSmtp) && _license.Remaining <= 0)
                {
                    _log($"{row.TargetSmtp}: übersprungen – keine freie Lizenz mehr ({_license.ConsumedCount}/{_license.MaxAccounts} verbraucht).");
                    continue;
                }

                var host = (row.SourceHostOverride ?? cfg.SourceHost ?? "").Trim();
                var port = row.SourcePortOverride ?? cfg.SourcePort;
                if (string.IsNullOrWhiteSpace(host))
                {
                    _log($"{row.TargetSmtp}: kein Quell-IMAP-Server gesetzt (zentrales Feld oder Override leer) – übersprungen.");
                    continue;
                }
                if (string.IsNullOrWhiteSpace(row.SourceUser))
                {
                    _log($"{row.TargetSmtp}: kein Quell-User gesetzt – übersprungen.");
                    continue;
                }

                // Quelle verbinden (mit 30-s-Timeout, damit es nicht ewig hängt)
                await using var source = new ImapSource(host, port, cfg.SourceUseSsl, row.SourceUser, row.SourcePassword);
                _log($"Verbinde Quelle {host}:{port} (SSL={cfg.SourceUseSsl}) als {row.SourceUser} …");
                try
                {
                    await source.ConnectAsync(WithTimeout(_cts.Token, 30));
                }
                catch (Exception ex) when (!_cts.IsCancellationRequested)
                {
                    _log($"{row.TargetSmtp}: Quell-IMAP {host}:{port} nicht erreichbar oder Login fehlgeschlagen – {ex.Message}");
                    continue;
                }

                // Ziel verbinden
                IMailTarget target = useGraph
                    ? new GraphTarget(cfg.TenantId, cfg.ClientId, cfg.ClientSecret)
                    : new ImapAppendTarget(tokens);
                await using var _ = target;
                _log("Verbinde Ziel (Exchange Online) …");
                try
                {
                    await target.ConnectAsync(row.TargetSmtp, WithTimeout(_cts.Token, 30));
                }
                catch (Exception ex) when (!_cts.IsCancellationRequested)
                {
                    _log($"{row.TargetSmtp}: Ziel (Exchange Online) nicht erreichbar oder Auth fehlgeschlagen – {ex.Message}");
                    _log("  Prüfen: Einrichtung OHNE Dry-Run ausgeführt? IMAP am Postfach aktiv, ServicePrincipal registriert, FullAccess erteilt? (Änderungen können einige Minuten brauchen.)");
                    continue;
                }

                // Jetzt erst Lizenz ziehen (nach erfolgreichem Connect), idempotent.
                if (!testMode) _license.TryConsume(row.TargetSmtp);

                var engine = new MigrationEngine(
                    cfg,
                    mapFolder: MapFolder,
                    progress: new Progress<MigrationProgress>(p =>
                    {
                        _bar.Style = ProgressBarStyle.Continuous;
                        if (p.FolderTotal > 0) { _bar.Maximum = p.FolderTotal; _bar.Value = Math.Min(p.FolderDone, p.FolderTotal); }
                        SetState($"{p.Mailbox} · {p.Folder} · {p.MessagesCopied} kopiert, {p.MessagesSkipped} übersprungen");
                    }),
                    log: _log);

                var (copied, skipped) = await engine.RunAsync(row.TargetSmtp, source, target, _cts.Token);
                _log($"{row.TargetSmtp}: {copied} kopiert, {skipped} übersprungen{(testMode ? " (Testmodus)" : "")}");
                _onLicenseChanged();
            }
            SetState("Fertig.");
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
}
