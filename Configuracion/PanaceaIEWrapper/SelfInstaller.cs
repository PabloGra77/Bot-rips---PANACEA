using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace PanaceaIEWrapper
{
    internal static class SelfInstaller
    {
        private static readonly string ShortcutPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                         "Panacea RIPS.lnk");

        /// <summary>
        /// Crea el acceso directo en el escritorio la primera vez que se ejecuta la app.
        /// </summary>
        public static void EnsureShortcut()
        {
            try
            {
                if (File.Exists(ShortcutPath)) return;

                string exePath = Assembly.GetExecutingAssembly().Location;
                CreateDesktopShortcut(exePath);
            }
            catch { }
        }

        private static void CreateDesktopShortcut(string exePath)
        {
            try
            {
                Type shellType = Type.GetTypeFromProgID("WScript.Shell");
                if (shellType == null) return;

                dynamic shell = Activator.CreateInstance(shellType);
                dynamic shortcut = shell.CreateShortcut(ShortcutPath);
                shortcut.TargetPath = exePath;
                shortcut.WorkingDirectory = Path.GetDirectoryName(exePath);
                shortcut.IconLocation = exePath + ",0";
                shortcut.Description = "Panacea RIPS Bot";
                shortcut.Save();
            }
            catch { }
        }
    }
}
