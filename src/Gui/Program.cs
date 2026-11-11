using System.Runtime.InteropServices;

namespace IMAP2M365.Gui;

internal static class Program
{
    [DllImport("kernel32.dll")] private static extern bool AllocConsole();
    [DllImport("kernel32.dll")] private static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    private const int SW_HIDE = 0;

    [STAThread]
    private static void Main()
    {
        // ExchangeOnlineManagement 3.x nutzt den WAM-Login-Broker, der ein Fenster-Handle
        // als Parent braucht. In einer WinForms-App ohne Konsole liefert GetConsoleWindow()
        // kein Handle ("A window handle must be configured"). Daher eine Konsole anlegen und
        // sofort verstecken – das Handle existiert dann für den Broker.
        if (GetConsoleWindow() == IntPtr.Zero && AllocConsole())
            ShowWindow(GetConsoleWindow(), SW_HIDE);

        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}
