using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PanaceaIEWrapper
{
    internal static class AutoUpdater
    {
        private const string GitHubOwner = "PabloGra77";
        private const string GitHubRepo  = "Bot-rips---PANACEA";

        private static readonly string ApiUrl =
            "https://api.github.com/repos/" + GitHubOwner + "/" + GitHubRepo + "/releases/latest";

        public static void CheckAndApply()
        {
            try
            {
                CheckAndApplyAsync().GetAwaiter().GetResult();
            }
            catch
            {
                // No interrumpir la app si falla la verificacion de actualizacion
            }
        }

        private static async Task CheckAndApplyAsync()
        {
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("User-Agent", "PanaceaAutoUpdater/1.0");
                client.Timeout = TimeSpan.FromSeconds(15);

                string json;
                try
                {
                    json = await client.GetStringAsync(ApiUrl).ConfigureAwait(false);
                }
                catch
                {
                    // Sin internet o repositorio no configurado: continuar normalmente
                    return;
                }

                // Extraer tag_name  ej: "v1.2.0"
                Match tagMatch = Regex.Match(json, "\"tag_name\"\\s*:\\s*\"v?([^\"]+)\"");
                if (!tagMatch.Success) return;

                if (!Version.TryParse(tagMatch.Groups[1].Value, out Version remoteVersion)) return;

                Version localVersion = Assembly.GetExecutingAssembly().GetName().Version;
                if (remoteVersion <= localVersion) return;

                // Extraer URL de descarga del primer asset .exe en el release
                Match urlMatch = Regex.Match(json,
                    "\"browser_download_url\"\\s*:\\s*\"([^\"]+\\.exe)\"",
                    RegexOptions.IgnoreCase);
                if (!urlMatch.Success) return;

                string downloadUrl = urlMatch.Groups[1].Value;

                DialogResult dr = MessageBox.Show(
                    "Hay una nueva version disponible: v" + remoteVersion + "\n\n" +
                    "La aplicacion se actualizara y reiniciara automaticamente.\n" +
                    "Version actual: v" + localVersion.ToString(3),
                    "Actualizacion disponible",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Information);

                if (dr != DialogResult.OK) return;

                string exePath = Assembly.GetExecutingAssembly().Location;
                string tempPath = exePath + ".new";

                byte[] data = await client.GetByteArrayAsync(downloadUrl).ConfigureAwait(false);
                File.WriteAllBytes(tempPath, data);

                // Script que espera a que cierre el proceso, reemplaza el exe y reinicia
                string batPath = Path.Combine(Path.GetTempPath(), "panacea_update.bat");
                File.WriteAllText(batPath,
                    "@echo off\r\n" +
                    "timeout /t 2 /nobreak >nul\r\n" +
                    "move /y \"" + tempPath + "\" \"" + exePath + "\"\r\n" +
                    "start \"\" \"" + exePath + "\"\r\n" +
                    "del \"%~f0\"\r\n");

                Process.Start(new ProcessStartInfo("cmd.exe", "/c \"" + batPath + "\"")
                {
                    WindowStyle = ProcessWindowStyle.Hidden,
                    CreateNoWindow = true
                });

                Environment.Exit(0);
            }
        }
    }
}
