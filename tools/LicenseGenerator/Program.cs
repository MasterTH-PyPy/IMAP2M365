namespace IMAP2M365.LicenseGen;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new GenForm());
    }
}
