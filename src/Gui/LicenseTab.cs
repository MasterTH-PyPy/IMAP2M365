using System.Reflection;
using IMAP2M365.Licensing;

namespace IMAP2M365.Gui;

/// <summary>
/// Lizenz-Reiter: zeigt Computer-ID/Rechnername/VM, nimmt Firmendaten auf, erzeugt die
/// verschlüsselte Aktivierungsanfrage für den Lizenzgeber und aktiviert den zurückgelieferten
/// Lizenzschlüssel. Zeigt Kontingent (gesamt/verbraucht/frei).
/// </summary>
public sealed class LicenseTab : UserControl
{
    private readonly LicenseManager _license;
    private readonly MachineId.Info _machine;
    private readonly Action<string> _log;

    private readonly TextBox _company = new() { Width = 300 };
    private readonly TextBox _contact = new() { Width = 300 };
    private readonly TextBox _address = new() { Width = 300 };
    private readonly TextBox _email = new() { Width = 300 };

    private readonly TextBox _request = new() { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, Font = new Font("Consolas", 8.5f) };
    private readonly TextBox _licenseKey = new() { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, Font = new Font("Consolas", 8.5f) };
    private readonly Label _status = new() { AutoSize = true, Font = new Font(DefaultFont, FontStyle.Bold) };
    private readonly Label _machineLabel = new() { AutoSize = true };

    public LicenseTab(LicenseManager license, MachineId.Info machine, Action<string> log)
    {
        _license = license;
        _machine = machine;
        _log = log;
        BuildUi();
        RefreshStatus();
    }

    private void BuildUi()
    {
        var top = new TableLayoutPanel { Dock = DockStyle.Top, Height = 210, ColumnCount = 4, Padding = new Padding(10) };

        _machineLabel.Text = $"Computer-ID:  {_machine.Id}        Rechner:  {_machine.MachineName}" +
                             (_machine.IsVirtual ? $"        ⚠ VM erkannt: {_machine.VmHint}" : "");
        top.Controls.Add(_machineLabel, 0, 0);
        top.SetColumnSpan(_machineLabel, 4);

        top.Controls.Add(_status, 0, 1);
        top.SetColumnSpan(_status, 4);

        top.Controls.Add(new Label { Text = "Firma:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
        top.Controls.Add(_company, 1, 2);
        top.Controls.Add(new Label { Text = "Ansprechpartner:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 2);
        top.Controls.Add(_contact, 3, 2);

        top.Controls.Add(new Label { Text = "Adresse:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
        top.Controls.Add(_address, 1, 3);
        top.Controls.Add(new Label { Text = "Mailadresse:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 3);
        top.Controls.Add(_email, 3, 3);

        var btnRequest = new Button { Text = "Aktivierungsanfrage erzeugen", AutoSize = true, Padding = new Padding(10, 4, 10, 4), Margin = new Padding(4) };
        var btnCopyReq = new Button { Text = "Kopieren", AutoSize = true, Padding = new Padding(10, 4, 10, 4), Margin = new Padding(4) };
        var btnSaveReq = new Button { Text = "Speichern…", AutoSize = true, Padding = new Padding(10, 4, 10, 4), Margin = new Padding(4) };
        var btnActivate = new Button { Text = "Lizenz aktivieren", AutoSize = true, Padding = new Padding(10, 4, 10, 4), Margin = new Padding(4) };
        btnRequest.Click += (_, _) => GenerateRequest();
        btnCopyReq.Click += (_, _) => { if (_request.TextLength > 0) Clipboard.SetText(_request.Text); };
        btnSaveReq.Click += (_, _) => SaveRequest();
        btnActivate.Click += (_, _) => Activate();

        var reqButtons = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 44, Padding = new Padding(8, 4, 8, 4) };
        reqButtons.Controls.AddRange(new Control[] { btnRequest, btnCopyReq, btnSaveReq });

        var reqBox = new GroupBox { Text = "1) Aktivierungsanfrage (an den Lizenzgeber senden)", Dock = DockStyle.Fill };
        reqBox.Controls.Add(_request);

        var keyBox = new GroupBox { Text = "2) Erhaltenen Lizenzschlüssel einfügen und aktivieren", Dock = DockStyle.Fill };
        var keyInner = new Panel { Dock = DockStyle.Fill };
        keyInner.Controls.Add(_licenseKey);
        var actButtons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(8, 4, 8, 4) };
        actButtons.Controls.Add(btnActivate);
        keyInner.Controls.Add(actButtons);
        keyBox.Controls.Add(keyInner);

        var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical };
        split.Panel1.Controls.Add(reqBox);
        split.Panel1.Controls.Add(reqButtons);
        split.Panel2.Controls.Add(keyBox);

        Controls.Add(split);
        Controls.Add(top);
    }

    private void GenerateRequest()
    {
        try
        {
            var req = new ActivationRequest
            {
                MachineId = _machine.Id,
                MachineName = _machine.MachineName,
                IsVirtualMachine = _machine.IsVirtual,
                VmHint = _machine.VmHint,
                Company = _company.Text.Trim(),
                Contact = _contact.Text.Trim(),
                Address = _address.Text.Trim(),
                Email = _email.Text.Trim(),
                ToolVersion = Assembly.GetExecutingAssembly()
                    .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "dev"
            };
            _request.Text = LicenseCrypto.EncryptRequest(req, EmbeddedPublicKey.Pem);
            _log("Aktivierungsanfrage erzeugt – an den Lizenzgeber senden.");
        }
        catch (Exception ex) { _log("Anfrage erzeugen fehlgeschlagen: " + ex.Message); }
    }

    private void SaveRequest()
    {
        if (_request.TextLength == 0) { GenerateRequest(); if (_request.TextLength == 0) return; }
        using var dlg = new SaveFileDialog { Filter = "Aktivierungsanfrage (*.req)|*.req", FileName = "aktivierung.req" };
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            File.WriteAllText(dlg.FileName, _request.Text);
            _log("Aktivierungsanfrage gespeichert: " + dlg.FileName);
        }
    }

    private void Activate()
    {
        var (ok, msg) = _license.Activate(_licenseKey.Text);
        _log(msg);
        if (ok) RefreshStatus();
    }

    private void RefreshStatus()
    {
        if (_license.HasValidLicense)
            _status.Text = $"Lizenz aktiv – {_license.Company}:  {_license.MaxAccounts} Postfächer gesamt, " +
                           $"{_license.ConsumedCount} verbraucht, {_license.Remaining} frei.";
        else
            _status.Text = "Keine gültige Lizenz auf diesem Rechner. Bitte Anfrage erzeugen und Lizenzschlüssel aktivieren.";
    }

    /// <summary>Von außen aufrufbar, damit der Reiter nach einer Migration das Kontingent aktualisiert.</summary>
    public void UpdateStatus() => RefreshStatus();
}
