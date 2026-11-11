using IMAP2M365.Licensing;

namespace IMAP2M365.LicenseGen;

/// <summary>
/// GUI-Generator für den Lizenzgeber. Hält den privaten Schlüssel.
/// Ablauf: privaten Key wählen (oder Schlüsselpaar erzeugen) -> Aktivierungsanfrage entschlüsseln
/// (zeigt Kundendaten + Computer-ID + VM) -> Kontenzahl setzen -> Lizenz erstellen.
/// </summary>
public sealed class GenForm : Form
{
    private readonly TextBox _privatePath = new() { Width = 420 };
    private readonly TextBox _request = new() { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, Font = new Font("Consolas", 8.5f) };
    private readonly TextBox _licenseOut = new() { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, Font = new Font("Consolas", 8.5f) };
    private readonly Label _info = new() { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 9f) };
    private readonly NumericUpDown _count = new() { Minimum = 1, Maximum = 100000, Value = 5, Width = 90 };
    private readonly Label _status = new() { Dock = DockStyle.Bottom, Height = 22, Text = "Bereit." };

    private ActivationRequest? _decoded;

    public GenForm()
    {
        Text = "IMAP2M365 – Lizenz-Generator";
        Width = 920; Height = 740; StartPosition = FormStartPosition.CenterScreen;
        try { Icon = System.Drawing.Icon.ExtractAssociatedIcon(Environment.ProcessPath!); } catch { }

        if (File.Exists("private.pem")) _privatePath.Text = Path.GetFullPath("private.pem");

        BuildUi();
    }

    private void BuildUi()
    {
        // 1) Privater Schlüssel
        var keyBox = new GroupBox { Text = "Privater Schlüssel (geheim halten)", Dock = DockStyle.Top, Height = 70 };
        var keyRow = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(8, 14, 8, 8) };
        var browse = Btn("Durchsuchen…", BrowsePrivate);
        var genKey = Btn("Schlüsselpaar erzeugen…", GenerateKeyPair);
        keyRow.Controls.AddRange(new Control[] { new Label { Text = "private.pem:", AutoSize = true, Margin = new Padding(2, 8, 4, 0) }, _privatePath, browse, genKey });
        keyBox.Controls.Add(keyRow);

        // 2) Aktivierungsanfrage
        var reqButtons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, Padding = new Padding(6, 4, 6, 4) };
        reqButtons.Controls.AddRange(new Control[] { Btn("Aus Datei…", LoadRequestFile), Btn("Entschlüsseln", DecodeRequest) });
        var reqBox = new GroupBox { Text = "1) Aktivierungsanfrage des Kunden einfügen", Dock = DockStyle.Fill };
        reqBox.Controls.Add(_request);
        reqBox.Controls.Add(reqButtons);

        // 3) Entschlüsselte Infos
        var infoBox = new GroupBox { Text = "2) Kundendaten / Computer-ID", Dock = DockStyle.Fill };
        infoBox.Controls.Add(_info);

        // 4) Lizenz erstellen
        var makeRow = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(6, 6, 6, 6) };
        makeRow.Controls.AddRange(new Control[]
        {
            new Label { Text = "Anzahl Postfächer:", AutoSize = true, Margin = new Padding(2, 8, 4, 0) },
            _count,
            Btn("Lizenz erstellen", CreateLicense),
            Btn("Schlüssel kopieren", () => { if (_licenseOut.TextLength > 0) Clipboard.SetText(_licenseOut.Text); }),
            Btn("Speichern…", SaveLicense)
        });
        var outBox = new GroupBox { Text = "3) Lizenzschlüssel (an den Kunden senden)", Dock = DockStyle.Fill };
        outBox.Controls.Add(_licenseOut);
        outBox.Controls.Add(makeRow);

        // Mittlerer Split: oben Anfrage+Info, unten Ausgabe
        var topSplit = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterDistance = 440 };
        topSplit.Panel1.Controls.Add(reqBox);
        topSplit.Panel2.Controls.Add(infoBox);

        var mainSplit = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterDistance = 380 };
        mainSplit.Panel1.Controls.Add(topSplit);
        mainSplit.Panel2.Controls.Add(outBox);

        Controls.Add(mainSplit);
        Controls.Add(keyBox);
        Controls.Add(_status);

        ShowInfo();
    }

    private static Button Btn(string text, Action onClick)
    {
        var b = new Button { Text = text, AutoSize = true, Padding = new Padding(10, 4, 10, 4), Margin = new Padding(4) };
        b.Click += (_, _) => onClick();
        return b;
    }

    private void SetStatus(string s) => _status.Text = s;

    private void BrowsePrivate()
    {
        using var dlg = new OpenFileDialog { Filter = "PEM (*.pem)|*.pem|Alle Dateien|*.*" };
        if (dlg.ShowDialog(this) == DialogResult.OK) _privatePath.Text = dlg.FileName;
    }

    private void GenerateKeyPair()
    {
        using var dlg = new FolderBrowserDialog { Description = "Zielordner für private.pem / public.pem" };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        var privPath = Path.Combine(dlg.SelectedPath, "private.pem");
        if (File.Exists(privPath) &&
            MessageBox.Show("private.pem existiert bereits. Überschreiben?", "Schlüsselpaar",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK)
            return;

        var (pub, priv) = LicenseCrypto.CreateKeyPair();
        File.WriteAllText(privPath, priv);
        File.WriteAllText(Path.Combine(dlg.SelectedPath, "public.pem"), pub);
        _privatePath.Text = privPath;
        _licenseOut.Text = pub;
        SetStatus("Schlüsselpaar erzeugt. public.pem in EmbeddedPublicKey.cs des Tools eintragen.");
        MessageBox.Show("Erzeugt:\n  private.pem (GEHEIM halten!)\n  public.pem (ins Tool einbetten)",
            "Schlüsselpaar", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void LoadRequestFile()
    {
        using var dlg = new OpenFileDialog { Filter = "Aktivierungsanfrage (*.req)|*.req|Alle Dateien|*.*" };
        if (dlg.ShowDialog(this) == DialogResult.OK) _request.Text = File.ReadAllText(dlg.FileName);
    }

    private string? ReadPrivate()
    {
        var p = _privatePath.Text.Trim();
        if (string.IsNullOrWhiteSpace(p) || !File.Exists(p))
        {
            MessageBox.Show("Bitte zuerst eine gültige private.pem wählen.", "Privater Schlüssel fehlt",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }
        return File.ReadAllText(p);
    }

    private void DecodeRequest()
    {
        var priv = ReadPrivate(); if (priv is null) return;
        try
        {
            _decoded = LicenseCrypto.DecryptRequest(_request.Text, priv);
            ShowInfo();
            SetStatus("Anfrage entschlüsselt.");
        }
        catch (Exception ex)
        {
            _decoded = null; ShowInfo();
            MessageBox.Show("Entschlüsselung fehlgeschlagen:\n" + ex.Message, "Fehler",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ShowInfo()
    {
        if (_decoded is null) { _info.Text = "(noch keine Anfrage entschlüsselt)"; return; }
        var r = _decoded;
        _info.Text =
            $"Firma:            {r.Company}\r\n" +
            $"Ansprechpartner:  {r.Contact}\r\n" +
            $"Adresse:          {r.Address}\r\n" +
            $"Mailadresse:      {r.Email}\r\n\r\n" +
            $"Computer-ID:      {r.MachineId}\r\n" +
            $"Rechner:          {r.MachineName}\r\n" +
            $"Virtuelle Masch.: {(r.IsVirtualMachine ? "JA  ->  " + r.VmHint : "nein")}\r\n" +
            $"Tool-Version:     {r.ToolVersion}\r\n" +
            $"Erstellt (UTC):   {r.CreatedUtc:u}";
        _info.ForeColor = r.IsVirtualMachine ? Color.Firebrick : SystemColors.ControlText;
    }

    private void CreateLicense()
    {
        if (_decoded is null)
        {
            MessageBox.Show("Bitte zuerst die Aktivierungsanfrage entschlüsseln.", "Keine Anfrage",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        var priv = ReadPrivate(); if (priv is null) return;

        if (_decoded.IsVirtualMachine &&
            MessageBox.Show("Achtung: Virtuelle Maschine erkannt. Trotzdem lizenzieren?", "VM erkannt",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK)
            return;

        var payload = new LicensePayload
        {
            MachineId = _decoded.MachineId,
            MaxAccounts = (int)_count.Value,
            Company = _decoded.Company,
            Contact = _decoded.Contact,
            Email = _decoded.Email
        };
        _licenseOut.Text = LicenseCrypto.SignLicense(payload, priv);
        SetStatus($"Lizenz erstellt: {_count.Value} Postfächer für {_decoded.Company}, gebunden an {_decoded.MachineId}.");
    }

    private void SaveLicense()
    {
        if (_licenseOut.TextLength == 0) return;
        using var dlg = new SaveFileDialog
        {
            Filter = "Lizenzschlüssel (*.key)|*.key",
            FileName = $"lizenz_{new string((_decoded?.Company ?? "kunde").Where(char.IsLetterOrDigit).ToArray())}_{_count.Value}.key"
        };
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            File.WriteAllText(dlg.FileName, _licenseOut.Text);
            SetStatus("Lizenz gespeichert: " + dlg.FileName);
        }
    }
}
