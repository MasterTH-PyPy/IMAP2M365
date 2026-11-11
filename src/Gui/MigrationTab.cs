using IMAP2M365.Core;
using IMAP2M365.Sources;
using IMAP2M365.Targets;

namespace IMAP2M365.Gui;

/// <summary>
/// Fährt die eigentliche Migration: pro Postfachzeile Quell-IMAP -> EXO (XOAUTH2-APPEND),
/// mit Fortschrittsbalken und Dry-Run. Nutzt den zentralen Quell-Server bzw. den Override pro Zeile.
/// </summary>
public sealed class MigrationTab : UserControl
{
    private readonly AppConfig _config;
    private readonly Action<string> _log;

    private readonly Button _start = new() { Text = "Migration starten", Width = 160 };
    private readonly Button _cancel = new() { Text = "Abbrechen", Width = 100, Enabled = false };
    private readonly CheckBox _dryRun = new() { Text = "Dry-Run", Checked = true, AutoSize = true };
    private readonly Label _backendLabel = new() { Text = "Ziel-Backend:", AutoSize = true, Padding = new Padding(12, 6, 2, 0) };
    private readonly ComboBox _backend = new() { Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ProgressBar _bar = new() { Dock = DockStyle.Top, Height = 22 };
    private readonly Label _state = new() { Dock = DockStyle.Top, Height = 22, Text = "Bereit." };

    private CancellationTokenSource? _cts;

    public MigrationTab(AppConfig config, Action<string> log)
    {
        _config = config;
        _log = log;

        _backend.Items.AddRange(new object[] { "IMAP-APPEND (XOAUTH2)", "Microsoft Graph" });
        _backend.SelectedIndex = _config.TargetKind.Equals("Graph", StringComparison.OrdinalIgnoreCase) ? 1 : 0;

        var buttons = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 44, Padding = new Padding(8, 6, 8, 6) };
        buttons.Controls.AddRange(new Control[] { _start, _cancel, _dryRun, _backendLabel, _backend });

        Controls.Add(_state);
        Controls.Add(_bar);
        Controls.Add(buttons);

        _start.Click += async (_, _) => await RunAsync();
        _cancel.Click += (_, _) => _cts?.Cancel();
    }

    private async Task RunAsync()
    {
        if (_config.Mailboxes.Count == 0) { _log("Keine Postfächer – erst im Reiter „Einrichtung" laden."); return; }
        if (string.IsNullOrWhiteSpace(_config.ClientId)) { _log("Keine ClientId – erst Einrichtung ausführen."); return; }

        _cts = new CancellationTokenSource();
        _start.Enabled = false; _cancel.Enabled = true;
        var dry = _dryRun.Checked;
        var useGraph = _backend.SelectedIndex == 1;

        var cfg = new MigrationConfig
        {
            TenantId = _config.TenantId,
            ClientId = _config.ClientId,
            ClientSecret = _config.ClientSecret,
            SourceHost = _config.SourceHost,
            SourcePort = _config.SourcePort,
            SourceUseSsl = _config.SourceUseSsl,
            DryRun = dry,
            ExcludedFolders = _config.ExcludedFolders,
            Target = useGraph ? TargetKind.Graph : TargetKind.ImapAppend
        };

        // Token-Provider nur für das IMAP-Backend nötig; GraphTarget baut eigene Credentials.
        var tokens = new ExchangeOnlineTokenProvider(cfg.TenantId, cfg.ClientId, cfg.ClientSecret);
        _log($"Ziel-Backend: {(useGraph ? "Microsoft Graph" : "IMAP-APPEND (XOAUTH2)")}");

        try
        {
            int i = 0;
            foreach (var row in _config.Mailboxes)
            {
                _cts.Token.ThrowIfCancellationRequested();
                SetState($"Postfach {++i}/{_config.Mailboxes.Count}: {row.TargetSmtp}");

                var host = row.SourceHostOverride ?? cfg.SourceHost;
                var port = row.SourcePortOverride ?? cfg.SourcePort;

                await using var source = new ImapSource(host, port, cfg.SourceUseSsl, row.SourceUser, row.SourcePassword);
                await source.ConnectAsync(_cts.Token);

                IMailTarget target = useGraph
                    ? new GraphTarget(cfg.TenantId, cfg.ClientId, cfg.ClientSecret)
                    : new ImapAppendTarget(tokens);
                await using var _ = target;
                if (!dry) await target.ConnectAsync(row.TargetSmtp, _cts.Token);

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
                _log($"{row.TargetSmtp}: {copied} kopiert, {skipped} übersprungen{(dry ? " (dry-run)" : "")}");
            }
            SetState("Fertig.");
        }
        catch (OperationCanceledException) { SetState("Abgebrochen."); _log("Migration abgebrochen."); }
        catch (Exception ex) { SetState("Fehler."); _log("Migration fehlgeschlagen: " + ex.Message); }
        finally { _start.Enabled = true; _cancel.Enabled = false; _cts?.Dispose(); _cts = null; }
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
