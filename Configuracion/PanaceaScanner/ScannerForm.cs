using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PanaceaScanner
{
    public class ScannerForm : Form
    {
        private readonly WebBrowser _wb;
        private readonly Label _status;
        private readonly string _logPath;
        private readonly string _username;
        private readonly string _password;
        private readonly System.Windows.Forms.Timer _rescanTimer;
        private int _rescanCount;
        private bool _scanDone;

        private const string LoginUrl = "http://181.51.196.194/Panacea/LogOnForm.aspx?ReturnUrl=%2fPanacea";

        public ScannerForm()
        {
            _logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Panacea", "scanner.log");

            _username = Environment.GetEnvironmentVariable("PANACEA_USER") ?? "1016018747";
            _password = Environment.GetEnvironmentVariable("PANACEA_PASS") ?? "Gole2026";

            // Limpiar log anterior
            try
            {
                string dir = Path.GetDirectoryName(_logPath);
                if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
                File.WriteAllText(_logPath, "[" + DateTime.Now + "] === NUEVO ESCANEO ===" + Environment.NewLine);
            }
            catch { }

            Text = "Panacea Scanner  |  Log: " + _logPath;
            WindowState = FormWindowState.Maximized;

            _status = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 32,
                Font = new System.Drawing.Font("Segoe UI", 9.5f, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.LightYellow,
                ForeColor = System.Drawing.Color.DarkBlue,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Text = "  Iniciando..."
            };

            _wb = new WebBrowser
            {
                Dock = DockStyle.Fill,
                ScriptErrorsSuppressed = true
            };
            _wb.DocumentCompleted += OnDocumentCompleted;
            // Interceptar ventanas nuevas (window.open) y redirigir en el mismo browser
            _wb.NewWindow += OnNewWindow;

            // Timer para re-escanear iframes que cargan tarde (modales dinamicos)
            _rescanTimer = new System.Windows.Forms.Timer { Interval = 1500 };
            _rescanTimer.Tick += (s, e) =>
            {
                _rescanCount++;
                if (_rescanCount > 20) { _rescanTimer.Stop(); return; } // max 30s
                var doc = _wb.Document;
                if (doc == null) return;
                string url = _wb.Url?.ToString() ?? string.Empty;
                if (url.IndexOf("LogOnForm", StringComparison.OrdinalIgnoreCase) >= 0) return;
                if (url.IndexOf("Whoami", StringComparison.OrdinalIgnoreCase) >= 0) return;
                // Escanear iframes buscando el dialogo de sede
                ScanAllFrames(doc, url);
            };

            Controls.Add(_wb);
            Controls.Add(_status);

            Load += (s, e) =>
            {
                try { ConfigureBrowserEmulation(); } catch { }
                _wb.Navigate(LoginUrl);
                SetStatus("Cargando pagina de login...");
            };
        }

        private static void ConfigureBrowserEmulation()
        {
            const int Ie11Mode = 11001;
            string exeName = Path.GetFileName(Application.ExecutablePath);
            using (var key = Registry.CurrentUser.CreateSubKey(
                @"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"))
            {
                if (key != null)
                    key.SetValue(exeName, Ie11Mode, RegistryValueKind.DWord);
            }
        }

        private void OnDocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string url = e?.Url?.ToString() ?? string.Empty;
            // Solo procesar la URL final (DocumentCompleted se dispara tambien para subrecursos)
            if (url != (_wb.Url?.ToString() ?? string.Empty)) return;

            Log("PAGE: " + url);

            HtmlDocument doc = _wb.Document;
            if (doc == null) return;

            // SIEMPRE inyectar override de window.open para que los popups se abran aqui
            InjectWindowOpenOverride(doc);

            // Whoami = pagina intermedia SSO, ignorar
            if (url.IndexOf("Whoami", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                SetStatus("Redireccion SSO (Whoami)... esperando pagina principal...");
                return;
            }

            // Pagina de login (LogOnForm o la pagina raiz SSO con usuario+contraseña)
            bool hasUser = GetEl(doc, "UserNameTextBox") != null || GetEl(doc, "UserNameTextBox_I") != null;
            bool hasPass = GetEl(doc, "PasswordTextBox") != null || GetEl(doc, "PasswordTextBox_I") != null;

            // Pagina raiz SSO: buscar campos genericos si no tiene los Infragistics
            HtmlElement userEl = GetEl(doc, "UserNameTextBox_I")
                              ?? GetEl(doc, "UserNameTextBox")
                              ?? FindByType(doc, "text");
            HtmlElement passEl = GetEl(doc, "PasswordTextBox_I")
                              ?? GetEl(doc, "PasswordTextBox")
                              ?? FindByType(doc, "password");

            bool isLoginPage = url.IndexOf("LogOnForm", StringComparison.OrdinalIgnoreCase) >= 0
                            || (passEl != null)
                            || url == "http://181.51.196.194/"
                            || url.StartsWith("http://181.51.196.194/?", StringComparison.OrdinalIgnoreCase);

            if (isLoginPage)
            {
                _scanDone = false;
                Log("LOGIN PAGE detectada. user=" + (userEl?.Id ?? "null") + " pass=" + (passEl?.Id ?? "null"));
                SetStatus("Login detectado. Intentando auto-login...");
                TryAutoLogin(doc, userEl, passEl);
                return;
            }

            // Pagina "elige usuario" (solo username, sin password): solo click en la flecha
            bool onlyUsername = userEl != null && passEl == null;
            if (onlyUsername)
            {
                _scanDone = false;
                Log("PASO 1 LOGIN: solo campo usuario. Clickeando flecha...");
                SetStatus("Pagina de seleccion de usuario. Clickeando flecha...");
                // En esta pagina el usuario ya esta rellenado, solo click en la flecha de avance
                HtmlElement arrow = GetEl(doc, "ImageButton1")
                                 ?? GetEl(doc, "InicioRedV2Button")
                                 ?? FindInputByType(doc, "image")
                                 ?? FindInputByType(doc, "submit");
                if (arrow != null)
                {
                    arrow.InvokeMember("click");
                    Log("PASO 1 LOGIN: click en flecha id='" + (arrow.Id ?? "") + "'");
                }
                return;
            }

            // Cualquier otra pagina: escanear si aun no lo hicimos
            if (!_scanDone)
            {
                _scanDone = true;
                SetStatus("Pagina cargada. Escaneando DOM e iframes...");
                ScanAndAttach(doc, url);
                // Arrancar re-escaneo periodico para capturar iframes/modales que cargan tarde
                _rescanCount = 0;
                _rescanTimer.Start();
                SetStatus("  LISTO. HAZ CLIC en:  1) fila SEDE  2) dropdown Contingencias  3) boton Aceptar  |  Log: " + _logPath);
            }
        }

        // Inyecta override de window.open para que los popups naveguen en el mismo browser
        private void InjectWindowOpenOverride(HtmlDocument doc)
        {
            try
            {
                doc.InvokeScript("eval", new object[]
                {
                    "if(!window.__popupOverriden){window.__popupOverriden=true;" +
                    "window.open=function(url,n,f){if(url&&url!='')window.location.href=url;return null;};}"
                });
            }
            catch { }
        }

        // Cuando Panacea intenta abrir una nueva ventana (window.open), redirigir aqui
        private void OnNewWindow(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true; // Cancelar apertura en IE externo
            Log("NewWindow interceptado - redirigiendo en mismo browser");
        }

        // -----------------------------------------------------------------------
        // AUTO-LOGIN
        // -----------------------------------------------------------------------
        private void TryAutoLogin(HtmlDocument doc, HtmlElement userEl, HtmlElement passEl)
        {
            string escapedUser = EscapeJs(_username);
            string escapedPass = EscapeJs(_password);

            // Intentar via Infragistics JS API primero
            try
            {
                string js = "(function(){"
                    + "function igSet(b,v){"
                    + "  var inp=document.getElementById(b+'_I');if(inp)inp.value=v;"
                    + "  var all=document.getElementsByTagName('input');"
                    + "  for(var i=0;i<all.length;i++){var n=all[i].name||'';"
                    + "    if(n===b||n===b+'$CVS')all[i].value=v;}"
                    + "}"
                    + "igSet('UserNameTextBox','" + escapedUser + "');"
                    + "igSet('PasswordTextBox','" + escapedPass + "');"
                    + "var a=document.getElementById('Automatico');if(a)a.value='1';"
                    + "var f=document.getElementById('FlagDominio');if(f)f.value='0';"
                    + "})();";
                doc.InvokeScript("eval", new object[] { js });
            }
            catch { }

            // Fallback DOM directo para campos genericos (pagina SSO raiz)
            if (userEl != null) try { userEl.SetAttribute("value", _username); } catch { }
            if (passEl != null) try { passEl.SetAttribute("value", _password); } catch { }

            Log("AUTO-LOGIN: user=" + (userEl?.Id ?? "null") + " pass=" + (passEl?.Id ?? "null"));

            // Click en boton de ingreso (buscar varios posibles IDs)
            HtmlElement btn = GetEl(doc, "ImageButton1")
                           ?? GetEl(doc, "InicioRedV2Button")
                           ?? FindInputByType(doc, "image")
                           ?? FindInputByType(doc, "submit");

            if (btn != null)
            {
                btn.InvokeMember("click");
                Log("AUTO-LOGIN: click en boton id='" + (btn.Id ?? string.Empty) + "' name='" + (btn.GetAttribute("name") ?? string.Empty) + "'");
                SetStatus("Credenciales enviadas. Esperando redireccion...");
            }
            else
            {
                // Ultimo recurso: submit del primer form
                try
                {
                    HtmlElementCollection forms = doc.GetElementsByTagName("form");
                    foreach (HtmlElement form in forms)
                    {
                        form.InvokeMember("submit");
                        Log("AUTO-LOGIN: form.submit()");
                        break;
                    }
                }
                catch { }
                Log("AUTO-LOGIN: boton no encontrado, intenta manualmente.");
                SetStatus("  Boton no encontrado. INGRESA MANUALMENTE en la ventana.");
            }
        }

        // -----------------------------------------------------------------------
        // ESCANEO + CAPTURA DE CLICKS
        // -----------------------------------------------------------------------
        private void ScanAndAttach(HtmlDocument doc, string url)
        {
            Log("=== SCAN BEGIN  url=" + url + " ===");
            int attached = AttachToDoc(doc, "MAIN");
            ScanAllFrames(doc, url);
            Log("=== SCAN END  main_handlers=" + attached + " ===");
        }

        // Escanea todos los iframes (llamado tanto en DocumentCompleted como por el timer)
        private void ScanAllFrames(HtmlDocument doc, string url)
        {
            if (doc == null) return;
            try
            {
                // Metodo 1: via HtmlWindow.Frames
                HtmlWindow win = doc.Window;
                if (win != null && win.Frames != null)
                {
                    int count = win.Frames.Count;
                    for (int fi = 0; fi < count; fi++)
                    {
                        try
                        {
                            HtmlWindow frame = win.Frames[fi];
                            HtmlDocument fd = frame?.Document;
                            if (fd == null) continue;
                            string furl = fd.Url?.ToString() ?? string.Empty;
                            int n = AttachToDoc(fd, "FRAME" + fi + "[" + furl + "]");
                            if (n > 0)
                                SetStatus("  IFRAME" + fi + " escaneado (" + n + " handlers). HAZ LOS 3 CLICS.");
                        }
                        catch { }
                    }
                }
            }
            catch { }

            // Metodo 2: via getElementsByTagName("iframe") - puede acceder a frames distintos
            try
            {
                HtmlElementCollection iframes = doc.GetElementsByTagName("iframe");
                int fi2 = 0;
                foreach (HtmlElement iframe in iframes)
                {
                    try
                    {
                        // Acceder al contentDocument via el objeto COM subyacente
                        dynamic domEl = iframe.DomElement;
                        dynamic cDoc = domEl.contentDocument;
                        if (cDoc == null) { fi2++; continue; }
                        string furl = string.Empty;
                        try { furl = (string)cDoc.location.href; } catch { }
                        Log("IFRAME[" + fi2 + "] url=" + furl);
                    }
                    catch { }
                    fi2++;
                }
            }
            catch { }
        }

        // Adjunta handlers de click a todos los elementos interactivos de un documento
        private int AttachToDoc(HtmlDocument doc, string context)
        {
            if (doc == null) return 0;
            int attached = 0;
            try
            {
                foreach (HtmlElement el in doc.All)
                {
                    string tag = (el.TagName ?? string.Empty).ToUpperInvariant();
                    if (tag != "INPUT" && tag != "SELECT" && tag != "OPTION"
                        && tag != "BUTTON" && tag != "A"
                        && tag != "TD" && tag != "TR" && tag != "SPAN" && tag != "DIV"
                        && tag != "LI" && tag != "UL")
                        continue;

                    string id   = el.Id ?? string.Empty;
                    string name = el.GetAttribute("name") ?? string.Empty;
                    string type = el.GetAttribute("type") ?? string.Empty;
                    string val  = el.GetAttribute("value") ?? string.Empty;
                    string cls  = el.GetAttribute("className") ?? string.Empty;
                    string txt  = (el.InnerText ?? string.Empty).Trim();
                    if (txt.Length > 120) txt = txt.Substring(0, 120);

                    bool hasId  = id.Length > 0 || name.Length > 0;
                    bool hasTxt = txt.Length > 0 && txt.Length < 100;
                    if (hasId || (hasTxt && (tag == "TD" || tag == "BUTTON" || tag == "A" || tag == "LI")))
                    {
                        Log(string.Format("  [{0}][{1}] id='{2}' name='{3}' type='{4}' val='{5}' class='{6}' txt='{7}'",
                            context, tag, id, name, type, val, cls, txt));
                    }

                    HtmlElement captured = el;
                    string ctx = context;
                    el.AttachEventHandler("onclick", (s2, e2) => OnClick(captured, ctx));
                    attached++;
                }
            }
            catch (Exception ex) { Log("[" + context + "] AttachToDoc error: " + ex.Message); }
            return attached;
        }

        // -----------------------------------------------------------------------
        // CAPTURA DE CLICK DEL USUARIO
        // -----------------------------------------------------------------------
        private void OnClick(HtmlElement el, string context)
        {
            string tag  = (el.TagName ?? string.Empty).ToUpperInvariant();
            string id   = el.Id ?? string.Empty;
            string name = el.GetAttribute("name") ?? string.Empty;
            string type = el.GetAttribute("type") ?? string.Empty;
            string val  = el.GetAttribute("value") ?? string.Empty;
            string cls  = el.GetAttribute("className") ?? string.Empty;
            string txt  = (el.InnerText ?? string.Empty).Trim();
            if (txt.Length > 120) txt = txt.Substring(0, 120);

            string parentInfo = string.Empty;
            try
            {
                HtmlElement p = el.Parent;
                if (p != null)
                    parentInfo = " parent=[" + p.TagName + " id='" + (p.Id ?? "") + "' name='" + (p.GetAttribute("name") ?? "") + "']";
            }
            catch { }

            Log(string.Format("*** CLICK [{0}][{1}] id='{2}' name='{3}' type='{4}' val='{5}' class='{6}' txt='{7}'{8} ***",
                context, tag, id, name, type, val, cls, txt, parentInfo));

            SetStatus("  CLICK: [" + tag + "] id='" + id + "' txt='" + txt.Substring(0, Math.Min(60, txt.Length)) + "'");
        }

        // -----------------------------------------------------------------------
        // HELPERS
        // -----------------------------------------------------------------------
        private static HtmlElement GetEl(HtmlDocument doc, string idOrName)
        {
            if (doc == null) return null;
            HtmlElement el = doc.GetElementById(idOrName);
            if (el != null) return el;
            foreach (HtmlElement e in doc.All)
            {
                if (string.Equals(e.GetAttribute("name"), idOrName, StringComparison.OrdinalIgnoreCase))
                    return e;
            }
            return null;
        }

        // Busca el primer <input> de un tipo dado (text, password, image, submit)
        private static HtmlElement FindByType(HtmlDocument doc, string inputType)
        {
            if (doc == null) return null;
            foreach (HtmlElement el in doc.GetElementsByTagName("input"))
            {
                if (string.Equals(el.GetAttribute("type"), inputType, StringComparison.OrdinalIgnoreCase))
                    return el;
            }
            return null;
        }

        private static HtmlElement FindInputByType(HtmlDocument doc, string inputType)
            => FindByType(doc, inputType);

        private void SetStatus(string msg)
        {
            if (InvokeRequired) { Invoke(new Action(() => SetStatus(msg))); return; }
            _status.Text = "  " + msg;
        }

        private void Log(string msg)
        {
            try
            {
                File.AppendAllText(_logPath,
                    "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + Environment.NewLine);
            }
            catch { }
        }

        private static string EscapeJs(string v)
        {
            if (string.IsNullOrEmpty(v)) return string.Empty;
            return v.Replace("\\", "\\\\").Replace("'", "\\'")
                    .Replace("\r", "").Replace("\n", "");
        }
    }
}
