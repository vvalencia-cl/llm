using LMM.UI;
using static System.Windows.Forms.Application;

namespace LMM;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Run(new MainForm());
    }
}