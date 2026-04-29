using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace PanaceaIEWrapper
{
    internal static class SelfInstaller
    {
        // Carpeta de instalacion en AppData del usuario
        private static readonly string InstallDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "PanaceaRIPS");

        private static readonly string ShortcutPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                         "Panacea RIPS.lnk");

        private static readonly string InstalledExe =
            Path.Combine(InstallDir, "PanaceaIEWrapper.exe");

        /// <summary>
        /// Si la aplicacion ya esta instalada en AppData, asegura el acceso directo y retorna.
        /// Si no, copia todos los archivos a AppData y relanza desde ahi.
        /// </summary>
        public static void EnsureInstalled()
        {
            try
            {
                string currentExe = Assembly.GetExecutingAssembly().Location;

                // Ya corremos desde la carpeta de instalacion
                if (string.Equals(Path.GetFullPath(currentExe),
                                  Path.GetFullPath(InstalledExe),
                                  StringComparison.OrdinalIgnoreCase))
                {
                    EnsureShortcut(currentExe);
                    return;
                }

                // Copiar toda la distribucion a AppData
                Directory.CreateDirectory(InstallDir);
                string sourceDir = Path.GetDirectoryName(currentExe) ?? AppDomain.CurrentDomain.BaseDirectory;
                CopyDirectoryRecursive(sourceDir, InstallDir);

                // Crear acceso directo al EXE instalado
                EnsureShortcut(InstalledExe);

                // Relanzar desde la carpeta de instalacion
                System.Diagnostics.Process.Start(InstalledExe);
                Application.Exit();
            }
            catch { }
        }

        // Mantener compatibilidad con Program.cs que llama EnsureShortcut()
        public static void EnsureShortcut()
        {
            try
            {
                string exePath = Assembly.GetExecutingAssembly().Location;
                EnsureShortcut(exePath);
            }
            catch { }
        }

        private static void EnsureShortcut(string exePath)
        {
            try
            {
                if (!File.Exists(ShortcutPath))
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

        private static void CopyDirectoryRecursive(string source, string dest)
        {
            Directory.CreateDirectory(dest);

            foreach (string file in Directory.GetFiles(source))
            {
                string fileName = Path.GetFileName(file);
                string destFile = Path.Combine(dest, fileName);
                try { File.Copy(file, destFile, overwrite: true); }
                catch { }
            }

            foreach (string subDir in Directory.GetDirectories(source))
            {
                string dirName = Path.GetFileName(subDir);
                // No copiar carpetas obj, .git, etc.
                if (dirName.StartsWith(".") || dirName.Equals("obj", StringComparison.OrdinalIgnoreCase))
                    continue;
                CopyDirectoryRecursive(subDir, Path.Combine(dest, dirName));
            }
        }
    }
}
