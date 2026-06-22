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
            public bool   IsZip          { get; set; }
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

                    // Buscar primero el ZIP correcto de distribución (RoBRips).
                    // IMPORTANTE: no usar ZIP de código fuente como asset del release, porque
                    // el actualizador extrae el ZIP sobre la carpeta instalada.
                    Match urlMatch = Regex.Match(json,
                        "\"browser_download_url\"\\s*:\\s*\"([^\"]*RoBRips[^\"]*\\.zip)\"",
                        RegexOptions.IgnoreCase);
                    bool isZip = urlMatch.Success;

                    if (!isZip)
                    {
                        urlMatch = Regex.Match(json,
                            "\"browser_download_url\"\\s*:\\s*\"([^\"]*Panacea[^\"]*\\.zip)\"",
                            RegexOptions.IgnoreCase);
                        isZip = urlMatch.Success;
                    }

                    if (!isZip)
                    {
                        urlMatch = Regex.Match(json,
                            "\"browser_download_url\"\\s*:\\s*\"([^\"]+\\.exe)\"",
                            RegexOptions.IgnoreCase);
                    }

                    if (!urlMatch.Success) return null;

                    // Normalizar version: si solo tiene 2 componentes (ej: "2.0"),
                    // ToString(3) falla. Usamos ToString() sin parametro.
                    string localStr  = localVersion.Major  + "." + localVersion.Minor  + "." + (localVersion.Build  >= 0 ? localVersion.Build  : 0);
                    string remoteStr = remoteVersion.Major + "." + remoteVersion.Minor + "." + (remoteVersion.Build >= 0 ? remoteVersion.Build : 0);

                    return new UpdateInfo
                    {
                        CurrentVersion = localStr,
                        NewVersion     = remoteStr,
                        DownloadUrl    = urlMatch.Groups[1].Value,
                        IsZip          = isZip
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

            string exePath = Assembly.GetExecutingAssembly().Location;
            string appDir  = Path.GetDirectoryName(exePath);
            string batPath = Path.Combine(Path.GetTempPath(), "panacea_update.bat");

            if (info.IsZip)
            {
                // ZIP: contiene el exe + todas las DLLs — extraer sobre el directorio actual
                string zipPath = Path.Combine(Path.GetTempPath(), "panacea_update.zip");
                File.WriteAllBytes(zipPath, data);

                // Usar PowerShell Expand-Archive (disponible en Windows 10+/PS5)
                // Las comillas internas se manejan con variables de PS para soportar espacios en rutas
                // Backup user config files antes de extraer para no sobreescribirlos
                string appDirPs = appDir.Replace("'", "''");
                string psCmd = string.Format(
                    "$z='{0}'; $d='{1}'; " +
                    "$cfg='$d\\bot-config.json'; $bak=if(Test-Path $cfg){{Get-Content -Raw $cfg}}else{{$null}}; " +
                    "$xcfg='$d\\PanaceaIEWrapper.exe.config'; $xbak=if(Test-Path $xcfg){{Get-Content -Raw $xcfg}}else{{$null}}; " +
                    "Expand-Archive -Path $z -DestinationPath $d -Force; " +
                    "if($bak){{Set-Content -Path $cfg -Value $bak -NoNewline}}; " +
                    "if($xbak){{Set-Content -Path $xcfg -Value $xbak -NoNewline}}; " +
                    "Remove-Item $z",
                    zipPath.Replace("'", "''"),
                    appDirPs);

                File.WriteAllText(batPath,
                    "@echo off\r\n" +
                    "timeout /t 2 /nobreak >nul\r\n" +
                    "powershell -NoProfile -NonInteractive -Command \"" + psCmd.Replace("\"", "\\\"") + "\"\r\n" +
                    "start \"\" \"" + exePath + "\"\r\n" +
                    "del \"%~f0\"\r\n");
            }
            else
            {
                // Fallback legacy: solo reemplazar el exe
                string tempPath = exePath + ".new";
                File.WriteAllBytes(tempPath, data);

                File.WriteAllText(batPath,
                    "@echo off\r\n" +
                    "timeout /t 2 /nobreak >nul\r\n" +
                    "move /y \"" + tempPath + "\" \"" + exePath + "\"\r\n" +
                    "start \"\" \"" + exePath + "\"\r\n" +
                    "del \"%~f0\"\r\n");
            }

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("cmd.exe", "/c \"" + batPath + "\"")
            {
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden,
                CreateNoWindow = true
            });

            Environment.Exit(0);
        }
    }
}
