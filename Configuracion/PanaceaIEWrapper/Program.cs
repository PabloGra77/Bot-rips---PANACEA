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
                MessageBox.Show(
                    "Ya hay una instancia de Panacea RIPS abierta o bloqueada en segundo plano.\n\n" +
                    "Cierre PanaceaIEWrapper.exe desde el Administrador de tareas y vuelva a abrirlo. Si no aparece, reinicie Windows.",
                    "Panacea RIPS",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
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

            // Pantalla de configuración inicial
            using (var startup = new StartupForm())
            {
                if (startup.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return; // El usuario canceló

                Application.Run(new MainForm(
                    startup.PanaceaUsername,
                    startup.PanaceaPassword,
                    startup.SelectedExcelPath));
            }
        }
    }
}
