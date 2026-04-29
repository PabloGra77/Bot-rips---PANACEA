using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PanaceaIEWrapper
{
    internal static class AutoUpdater
    {
        private const string GitHubOwner = "PabloGra77";
        private const string GitHubRepo  = "Bot-rips---PANACEA";

        private static readonly string ApiUrl =
            "https://api.github.com/repos/" + GitHubOwner + "/" + GitHubRepo + "/releases/latest";

        private sealed class UpdateInfo
        {
            public string NewVersion     { get; set; }
            public string CurrentVersion { get; set; }
            public string DownloadUrl    { get; set; }
        }

        public static void CheckAndApply()
        {
            try
            {
                UpdateInfo info = FetchUpdateInfo();
                if (info == null) return;

                using (var form = new UpdateForm(info.CurrentVersion, info.NewVersion))
                {
                    form.ShowDialog();
                    if (!form.ShouldUpdate) return;
                }

                DownloadAndApply(info);
            }
            catch { }
        }

        private static UpdateInfo FetchUpdateInfo()
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "PanaceaAutoUpdater/1.0");
                    client.Timeout = TimeSpan.FromSeconds(15);

                    string json = client.GetStringAsync(ApiUrl).GetAwaiter().GetResult();

                    Match tagMatch = Regex.Match(json, "\"tag_name\"\\s*:\\s*\"v?([^\"]+)\"");
                    if (!tagMatch.Success) return null;

                    if (!Version.TryParse(tagMatch.Groups[1].Value, out Version remoteVersion)) return null;

                    Version localVersion = Assembly.GetExecutingAssembly().GetName().Version;
                    if (remoteVersion <= localVersion) return null;

                    Match urlMatch = Regex.Match(json,
                        "\"browser_download_url\"\\s*:\\s*\"([^\"]+\\.exe)\"",
                        RegexOptions.IgnoreCase);
                    if (!urlMatch.Success) return null;

                    return new UpdateInfo
                    {
                        CurrentVersion = localVersion.ToString(3),
                        NewVersion     = remoteVersion.ToString(3),
                        DownloadUrl    = urlMatch.Groups[1].Value
                    };
                }
            }
            catch
            {
                return null;
            }
        }

        private static void DownloadAndApply(UpdateInfo info)
        {
            byte[]    data       = null;
            Exception downloadEx = null;

            using (var dlgProgress = new DownloadProgressForm())
            {
                var worker = new BackgroundWorker();
                worker.DoWork += (s, e) =>
                {
                    try
                    {
                        using (var client = new HttpClient())
                        {
                            client.DefaultRequestHeaders.Add("User-Agent", "PanaceaAutoUpdater/1.0");
                            client.Timeout = TimeSpan.FromMinutes(5);
                            data = client.GetByteArrayAsync(info.DownloadUrl).GetAwaiter().GetResult();
                        }
                    }
                    catch (Exception ex) { downloadEx = ex; }
                };
                worker.RunWorkerCompleted += (s, e) => dlgProgress.Close();
                worker.RunWorkerAsync();

                dlgProgress.ShowDialog();
            }

            if (downloadEx != null || data == null)
            {
                MessageBox.Show(
                    "No se pudo descargar la actualizacion.\n\n" +
                    (downloadEx?.Message ?? "Respuesta vacia del servidor."),
                    "Panacea RIPS",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string exePath  = Assembly.GetExecutingAssembly().Location;
            string tempPath = exePath + ".new";
            File.WriteAllBytes(tempPath, data);

            string batPath = Path.Combine(Path.GetTempPath(), "panacea_update.bat");
            File.WriteAllText(batPath,
                "@echo off\r\n" +
                "timeout /t 2 /nobreak >nul\r\n" +
                "move /y \"" + tempPath + "\" \"" + exePath + "\"\r\n" +
                "start \"\" \"" + exePath + "\"\r\n" +
                "del \"%~f0\"\r\n");

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("cmd.exe", "/c \"" + batPath + "\"")
            {
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden,
                CreateNoWindow = true
            });

            Environment.Exit(0);
        }
    }
}
