using LMM.UI;
using FormsApplication = System.Windows.Forms.Application;

namespace LMM;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        // Better DPI behavior across 100%/150% and multi-monitor
        FormsApplication.SetHighDpiMode(HighDpiMode.PerMonitorV2);

        FormsApplication.EnableVisualStyles();
        FormsApplication.SetCompatibleTextRenderingDefault(false);

        FormsApplication.Run(new MainForm());
    }
}