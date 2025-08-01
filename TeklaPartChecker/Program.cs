using System;
using System.Configuration.Assemblies;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace TeklaPartChecker
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (var splash = new SplashForm())
            {
                splash.ShowDialog();
            }

            Application.Run(new MainForm());
        }

    }
}
