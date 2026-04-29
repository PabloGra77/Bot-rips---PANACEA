using System;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using PanaceaIEWrapper.Bot;

namespace PanaceaIEWrapper
{
    internal static class Program
    {
        private static Mutex _singleInstanceMutex;

        [STAThread]
        private static void Main()
        {
            bool createdNew;
            _singleInstanceMutex = new Mutex(true, "PANACEA_IE_WRAPPER_SINGLE_INSTANCE", out createdNew);
            if (!createdNew)
            {
                return;
            }

            string[] args = Environment.GetCommandLineArgs().Skip(1).ToArray();
            if (args.Any(a => string.Equals(a, "--bot", StringComparison.OrdinalIgnoreCase)))
            {
                int exitCode = BotHost.RunAsync(args).GetAwaiter().GetResult();
                Environment.Exit(exitCode);
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            SelfInstaller.EnsureShortcut();
            AutoUpdater.CheckAndApply();
            Application.Run(new MainForm());
        }
    }
}
