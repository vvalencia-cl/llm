using LMM.UI;
using static System.Windows.Forms.Application;

namespace LMM;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Run(new MainForm());
    }
}