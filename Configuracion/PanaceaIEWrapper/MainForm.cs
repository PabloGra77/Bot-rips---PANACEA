using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using OfficeOpenXml;
using PanaceaIEWrapper.Bot;

namespace PanaceaIEWrapper
{
    public partial class MainForm : Form
    {
        private const string PanaceaUrl = "http://181.51.196.194/Panacea";
        private const string LoginPath = "/Panacea/LogOnForm.aspx?ReturnUrl=%2fPanacea";
        private const string AppExeName = "PanaceaIEWrapper.exe";
        private bool _autoLoginAttempted;
        private readonly System.Windows.Forms.Timer _autoLoginTimer;
        private readonly System.Windows.Forms.Timer _authOutcomeTimer;
        private int _autoLoginTries;
        private int _authOutcomeChecks;
        private bool _autoLoginFailureNotified;
        private bool _missingCredentialsLogged;
        private bool _loginSubmitDetected;
        private bool _authOutcomeResolved;
        private bool _alternateStrategyTried;
        private DateTime _loginSubmitUtc;
        private string _username;
        private string _password;
        private readonly string _uiLogPath;

        // RIPS flow
        private bool _ripsNavigated;
        private bool _ripsFlowStarted;
        private bool _ripsFlowDone;
        private bool _waitingForUserConvenio;  // true = pausado esperando que usuario seleccione convenio
        private bool _awaitingBuscarResult;    // true = Buscar fue clickeado, esperando que la pagina de resultados cargue
        private bool _awaitingCompletarForm;    // true = click Completar hecho, esperando que cargue el formulario
        private string _completarFormUrl;       // URL exacta del frame del formulario Completar
        private int _ripsStep;
        private int _completarStep;
        private int _ripsCcRetries;
        private int _convenioRetries;
        private int _buscarSinResultadoRetries;
        private int _gridEmptyPolls;   // re-chequeos del grid vacio antes de re-clickar Buscar
        private int _ripsGeneration;  // se incrementa cada vez que comienza un nuevo ciclo CC; cancela timers huerfanos
        private System.Windows.Forms.Timer _ripsTimer;
        private System.Windows.Forms.Timer _completarTimer;
        private string _cedulaPaciente;
        private string _ripsFechaInicio;
        private string _ripsFechaFin;
        private List<RipsRecord> _ripsRecords;
        private int _ripsRecordIndex;

        // Nuevos campos para la UI mejorada
        private string _overrideExcelPath;
        private bool _botPaused;

        [DllImport("wininet.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool InternetSetCookie(string lpszUrl, string lpszCookieName, string lpszCookieData);

        public MainForm() : this(null, null, null) { }

        public MainForm(string username, string password, string excelPath)
        {
            InitializeComponent();
            WindowState = FormWindowState.Maximized;
            _autoLoginTimer = new System.Windows.Forms.Timer { Interval = 1200 };
            _autoLoginTimer.Tick += AutoLoginTimer_Tick;
            _authOutcomeTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            _authOutcomeTimer.Tick += AuthOutcomeTimer_Tick;
            _uiLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Panacea", "ui-autologin.log");

            // Credenciales y archivo pre-cargados desde StartupForm
            if (!string.IsNullOrWhiteSpace(username)) _username = username;
            if (!string.IsNullOrWhiteSpace(password)) _password = password;
            if (!string.IsNullOrWhiteSpace(excelPath)) _overrideExcelPath = excelPath;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                ConfigureBrowserEmulation();
            }
            catch
            {
            }

            webBrowser1.ScriptErrorsSuppressed = true;
            webBrowser1.NewWindow += (s, ne) => ne.Cancel = true;
            // Bloquear navegacion a FinSesion.htm antes de que cargue y llame window.close()
            webBrowser1.Navigating += (s, ne) =>
            {
                string nav = ne.Url?.ToString() ?? string.Empty;
                if (nav.IndexOf("FinSesion", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ne.Cancel = true;
                    WriteUiLog("Navegacion a FinSesion bloqueada. Redirigiendo a login...");
                    // Navegar a la pagina de login directamente
                    try { webBrowser1.Navigate("http://181.51.196.194/Panacea/LogOnForm.aspx"); } catch { }
                }
            };
            // Solo cargar credenciales desde config si no llegaron del StartupForm
            if (string.IsNullOrWhiteSpace(_username) || string.IsNullOrWhiteSpace(_password))
                LoadCredentials();

            // Cargar Excel si viene pre-seleccionado desde StartupForm
            if (!string.IsNullOrWhiteSpace(_overrideExcelPath))
            {
                txtExcelPath.Text = System.IO.Path.GetFileName(_overrideExcelPath);
                try { LoadRipsExcel(); UpdateBotProgress(); }
                catch (Exception exXls) { WriteUiLog("LoadRipsExcel al iniciar (DLL faltante?): " + exXls.Message); }
            }

            // IE COM deshabilitado para evitar doble ventana - todo via WebBrowser embebido

            bool preLoginOk = TryHttpPreLogin();
            if (!preLoginOk)
            {
                _autoLoginTimer.Start();
            }
            WriteUiLog("MainForm_Load: autologin timer iniciado.");

            try
            {
                webBrowser1.Navigate(PanaceaUrl);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "No se pudo abrir Panacea.\n\nDetalle: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static void ConfigureBrowserEmulation()
        {
            const int Ie11Mode = 11001;
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(
                       @"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"))
            {
                if (key == null)
                    return;

                key.SetValue(AppExeName, Ie11Mode, RegistryValueKind.DWord);
            }
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string url = e?.Url?.ToString() ?? string.Empty;
            WriteUiLog("DocumentCompleted: " + url);

            // Inyectar override window.open en CADA pagina para evitar ventanas emergentes
            InjectWindowOpenOverrideManaged(webBrowser1.Document);
            // Inyectar tambien en todos los frames (FinSesion.htm llama window.close desde un frame)
            try
            {
                if (webBrowser1.Document?.Window?.Frames != null)
                {
                    foreach (HtmlWindow frame in webBrowser1.Document.Window.Frames)
                    {
                        try { InjectWindowOpenOverrideManaged(frame.Document); } catch { }
                    }
                }
            }
            catch { }

            EvaluateAuthOutcome(e?.Url, webBrowser1.Document);
            TryAutoLogin();

            // Cuando el frame ValidacionesUsuario termina de cargar, iniciar seleccion de sede
            if (!_sedeSelectionDone && !_sedeSelectionStarted
                && url.IndexOf("ValidacionesUsuario", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                _sedeSelectionStarted = true;
                WriteUiLog("Frame ValidacionesUsuario cargado. Iniciando seleccion de sede en 1.5s...");
                var t = new System.Windows.Forms.Timer { Interval = 1500 };
                t.Tick += (ts, te) => { ((System.Windows.Forms.Timer)ts).Stop(); StartSedeSelectionFlow(); };
                t.Start();
            }

            ScanSedesPage(webBrowser1.Document, url);

            // Cuando carga DefaultForm, navegar al modulo RIPS
            // Cuando carga DefaultForm (directamente o via redirect de ValidacionesUsuario),
            // navegar al modulo RIPS. Se detecta el redirect de ValidacionesUsuario porque
            // el servidor lo devuelve ANTES de que DefaultForm y LeftSide terminen de cargar,
            // dando tiempo suficiente con el timer de 4s.
            bool isDefaultFormReady =
                // Señal temprana: ValidacionesUsuario redirigiendo a DefaultForm (funciona mejor)
                (url.IndexOf("ValidacionesUsuario", StringComparison.OrdinalIgnoreCase) >= 0
                 && url.IndexOf("DefaultForm", StringComparison.OrdinalIgnoreCase) >= 0)
                // Señal tardía: DefaultForm.aspx cargado como URL principal
                || url.EndsWith("DefaultForm.aspx", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(url, @"DefaultForm\.aspx\?", RegexOptions.IgnoreCase);

            if (_sedeSelectionDone && !_ripsNavigated && isDefaultFormReady)
            {
                _ripsNavigated = true;
                WriteUiLog("Señal DefaultForm detectada en: " + url + " — Navegando a RIPS en 4s...");
                var tNav = new System.Windows.Forms.Timer { Interval = 4000 };
                tNav.Tick += (ts, te) =>
                {
                    ((System.Windows.Forms.Timer)ts).Stop();
                    string ripsUrl = "http://181.51.196.194/Panacea/FacturacionPaciente/Facturacion_Pacientes/ConsultarEstadoRipsForm.aspx?IdOpcion=828";
                    WriteUiLog("Navegando a RIPS via JS (con Referer): " + ripsUrl);
                    // Navegar via JS desde dentro de la pagina para que el navegador
                    // envie el Referer automaticamente — el servidor lo requiere para validar
                    // que la navegacion viene de dentro del sistema y no es directa.
                    bool jsOk = false;
                    try
                    {
                        // Intentar navegar desde el frame DefaultForm.aspx (FRAME2)
                        HtmlDocument mainDoc = webBrowser1.Document;
                        if (mainDoc != null)
                        {
                            foreach (HtmlWindow frame in mainDoc.Window.Frames)
                            {
                                try
                                {
                                    string furl = frame.Document?.Url?.ToString() ?? string.Empty;
                                    if (furl.IndexOf("DefaultForm", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        frame.Document.InvokeScript("eval", new object[]
                                            { "window.location.href='" + ripsUrl + "';" });
                                        WriteUiLog("Navegacion JS desde FRAME DefaultForm OK");
                                        jsOk = true;
                                        break;
                                    }
                                }
                                catch { }
                            }
                            if (!jsOk)
                            {
                                // Navegar desde el documento principal (Mainframe)
                                mainDoc.InvokeScript("eval", new object[]
                                    { "window.location.href='" + ripsUrl + "';" });
                                WriteUiLog("Navegacion JS desde documento principal OK");
                                jsOk = true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteUiLog("JS navigate fallback a webBrowser1.Navigate: " + ex.Message);
                    }
                    if (!jsOk)
                    {
                        webBrowser1.Navigate(ripsUrl);
                    }
                };
                tNav.Start();
            }

            // Mientras el bot espera que el usuario seleccione el convenio,
            // NO hacer nada en DocumentCompleted — el usuario presionara "Continuar bot"
            if (_waitingForUserConvenio)
            {
                WriteUiLog("DocumentCompleted ignorado (esperando convenio manual): " + url);
                return;
            }

            // Bot en pausa — no procesar eventos de navegacion
            if (_botPaused)
            {
                WriteUiLog("DocumentCompleted ignorado (bot pausado): " + url);
                return;
            }

            // Cuando carga el formulario Completar RIPS
            if (_awaitingCompletarForm
                && !url.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase)
                && url.Length > 10
                && url.IndexOf("ConsultarEstadoRipsForm", StringComparison.OrdinalIgnoreCase) < 0
                && (url.IndexOf("Panacea", StringComparison.OrdinalIgnoreCase) >= 0
                    || url.IndexOf(".aspx", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                _awaitingCompletarForm = false;
                _completarFormUrl = url;  // guardar URL exacta para que JsInRips encuentre el frame correcto
                WriteUiLog("Formulario Completar cargado: " + url);
                var tComp = new System.Windows.Forms.Timer { Interval = 1500 };
                tComp.Tick += (tsC, teC) => { ((System.Windows.Forms.Timer)tsC).Stop(); StartCompletarFlow(); };
                tComp.Start();
                return;
            }

            // Cuando la pagina de resultados carga tras el Buscar automatico, procesar el grid
            if (_awaitingBuscarResult
                && url.IndexOf("ConsultarEstadoRipsForm", StringComparison.OrdinalIgnoreCase) >= 0
                && !url.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase))
            {
                _awaitingBuscarResult = false;
                WriteUiLog("Pagina de resultados cargada. Procesando grid RIPS...");
                var tGrid = new System.Windows.Forms.Timer { Interval = 2000 };
                tGrid.Tick += (tsG, teG) => { ((System.Windows.Forms.Timer)tsG).Stop(); StartRipsGridProcessing(); };
                tGrid.Start();
                return;
            }

            // Cuando carga la pagina RIPS, iniciar el flujo de consulta
            if (_sedeSelectionDone && !_ripsFlowStarted && !_ripsFlowDone
                && url.IndexOf("ConsultarEstadoRipsForm", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                _ripsFlowStarted = true;
                _ripsStep = 0;
                _ripsCcRetries = 0;
                WriteUiLog("Pagina RIPS cargada. Iniciando flujo RIPS en 2s...");
                var tRips = new System.Windows.Forms.Timer { Interval = 2000 };
                tRips.Tick += (ts, te) => { ((System.Windows.Forms.Timer)ts).Stop(); StartRipsFlow(); };
                tRips.Start();
            }

            // Si el timer RIPS estaba corriendo y navegamos fuera de la pagina RIPS, detenerlo
            if (_ripsFlowStarted && !_ripsFlowDone
                && url.IndexOf("ConsultarEstadoRipsForm", StringComparison.OrdinalIgnoreCase) < 0
                && !url.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase)
                && !url.StartsWith("about:", StringComparison.OrdinalIgnoreCase)
                && url.IndexOf("cargando", StringComparison.OrdinalIgnoreCase) < 0
                && url.Length > 10)
            {
                if (_ripsTimer != null && _ripsTimer.Enabled)
                {
                    _ripsTimer.Stop();
                    WriteUiLog("Timer RIPS detenido por navegacion a: " + url);
                }
                // Si navegamos al login, reiniciar flag para reintento tras re-login
                if (url.IndexOf("LogOnForm", StringComparison.OrdinalIgnoreCase) >= 0
                    || url.IndexOf("FinSesion", StringComparison.OrdinalIgnoreCase) >= 0
                    || url.IndexOf("Whoami", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    _ripsFlowStarted = false;
                    _ripsNavigated = false;
                    WriteUiLog("Detectado redirect a login/FinSesion. Se reintentara el flujo RIPS tras re-login.");
                }
            }
        }

        private void InjectWindowOpenOverrideManaged(HtmlDocument doc)
        {
            if (doc == null) return;
            try
            {
                doc.InvokeScript("eval", new object[]
                {
                    "(function(){if(window.__popupOverriden)return;window.__popupOverriden=true;" +
                    "window.open=function(url,n,f){if(url&&url!=''){window.location.href=url;}return null;};" +
                    "window.close=function(){};" +
                    "})()"
                });
            }
            catch { }
        }

        private void AutoLoginTimer_Tick(object sender, EventArgs e)
        {
            TryAutoLogin();
            if (_autoLoginAttempted || _autoLoginTries >= 15)
            {
                _autoLoginTimer.Stop();
                WriteUiLog(_autoLoginAttempted
                    ? "AutoLoginTimer detenido: login intentado."
                    : "AutoLoginTimer detenido: maximo de intentos alcanzado.");

                if (!_autoLoginAttempted && !_autoLoginFailureNotified)
                {
                    _autoLoginFailureNotified = true;
                    MessageBox.Show(
                        "No se pudo ejecutar el autologin visual.\n\n" +
                        "Revisa el log: %LOCALAPPDATA%\\Panacea\\ui-autologin.log",
                        "Panacea - Diagnostico Autologin",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
        }

        private void AuthOutcomeTimer_Tick(object sender, EventArgs e)
        {
            if (_authOutcomeResolved)
            {
                _authOutcomeTimer.Stop();
                return;
            }

            _authOutcomeChecks++;
            EvaluateAuthOutcome(webBrowser1.Url, webBrowser1.Document);
            if (_authOutcomeResolved || _authOutcomeChecks >= 10)
            {
                _authOutcomeTimer.Stop();
                if (!_authOutcomeResolved)
                {
                    WriteUiLog("No se pudo confirmar el resultado del login visual en el tiempo esperado.");
                }
            }
        }

        private void LoadCredentials()
        {
            _username = Environment.GetEnvironmentVariable("PANACEA_USER");
            _password = Environment.GetEnvironmentVariable("PANACEA_PASS");

            if (!string.IsNullOrWhiteSpace(_username) && !string.IsNullOrWhiteSpace(_password))
            {
                WriteUiLog("Credenciales cargadas desde variables de entorno.");
                return;
            }

            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bot-config.json");
            if (!File.Exists(configPath))
            {
                WriteUiLog("No se encontro bot-config.json en: " + configPath);
                return;
            }

            try
            {
                BotDefinition definition = BotConfigLoader.Load(configPath);
                if (string.IsNullOrWhiteSpace(_username) && definition.Variables.TryGetValue("username", out string user))
                {
                    _username = user;
                }

                if (string.IsNullOrWhiteSpace(_password) && definition.Variables.TryGetValue("password", out string pass))
                {
                    _password = pass;
                }

                // Variables RIPS
                if (definition.Variables.TryGetValue("rips_cedula", out string ced))
                    _cedulaPaciente = ced;
                if (definition.Variables.TryGetValue("rips_fecha_inicio", out string fi))
                    _ripsFechaInicio = fi;
                if (definition.Variables.TryGetValue("rips_fecha_fin", out string ff))
                    _ripsFechaFin = ff;

                WriteUiLog(string.Format(
                    "Credenciales desde config: user={0}, pass={1}",
                    string.IsNullOrWhiteSpace(_username) ? "NO" : "SI",
                    string.IsNullOrWhiteSpace(_password) ? "NO" : "SI"));
            }
            catch
            {
                WriteUiLog("Error al leer bot-config.json.");
            }
        }

        private bool TryHttpPreLogin()
        {
            if (string.IsNullOrWhiteSpace(_username) || string.IsNullOrWhiteSpace(_password))
            {
                return false;
            }

            try
            {
                var baseUri = new Uri("http://181.51.196.194", UriKind.Absolute);
                var loginUri = new Uri(baseUri, LoginPath);
                var cookieContainer = new CookieContainer();
                var handler = new HttpClientHandler
                {
                    UseCookies = true,
                    CookieContainer = cookieContainer,
                    AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
                };

                using (handler)
                using (var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(25) })
                {
                    string firstHtml = client.GetStringAsync(loginUri).GetAwaiter().GetResult();

                    string body = BuildLoginBody(firstHtml, _username, _password);
                    using (var request = new HttpRequestMessage(HttpMethod.Post, loginUri))
                    {
                        request.Headers.Referrer = loginUri;
                        request.Content = new StringContent(body, Encoding.UTF8, "application/x-www-form-urlencoded");
                        using (HttpResponseMessage response = client.SendAsync(request).GetAwaiter().GetResult())
                        {
                            string html = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                            bool loginFormReturned = ContainsToken(html, "UserNameTextBox") && ContainsToken(html, "LogOnForm.aspx");
                            if (!response.IsSuccessStatusCode || loginFormReturned)
                            {
                                WriteUiLog("Prelogin HTTP no autentico sesion. Status=" + (int)response.StatusCode);
                                return false;
                            }
                        }
                    }
                }

                ApplyCookiesToIe(baseUri, cookieContainer);
                _authOutcomeResolved = true;
                _autoLoginAttempted = true;
                WriteUiLog("Prelogin HTTP exitoso. Cookies de sesion aplicadas a IE.");
                return true;
            }
            catch (Exception ex)
            {
                WriteUiLog("Prelogin HTTP fallo: " + ex.Message);
                return false;
            }
        }

        private static string BuildLoginBody(string html, string username, string password)
        {
            string[] viewstateKeys =
            {
                "__VIEWSTATE", "__VIEWSTATE1", "__VIEWSTATE2", "__VIEWSTATE3", "__VIEWSTATE4", "__VIEWSTATE5",
                "__VIEWSTATE6", "__VIEWSTATE7", "__VIEWSTATE8", "__VIEWSTATE9", "__VIEWSTATE10", "__VIEWSTATE11",
                "__VIEWSTATE12", "__VIEWSTATE13", "__VIEWSTATE14", "__VIEWSTATE15", "__VIEWSTATE16"
            };

            var sb = new StringBuilder();
            AppendForm(sb, "__EVENTTARGET", string.Empty);
            AppendForm(sb, "__EVENTARGUMENT", string.Empty);
            AppendForm(sb, "__VIEWSTATEFIELDCOUNT", "17");

            foreach (string key in viewstateKeys)
            {
                AppendForm(sb, key, GetInputValue(html, key));
            }

            AppendForm(sb, "__VIEWSTATEGENERATOR", GetInputValue(html, "__VIEWSTATEGENERATOR"));
            AppendForm(sb, "__SCROLLPOSITIONX", "0");
            AppendForm(sb, "__SCROLLPOSITIONY", "0");
            AppendForm(sb, "__EVENTVALIDATION", GetInputValue(html, "__EVENTVALIDATION"));
            AppendForm(sb, "Automatico", "1");
            AppendForm(sb, "UserNameTextBox", username);
            AppendForm(sb, "FlagDominio", "0");
            AppendForm(sb, "UserAlterno", string.Empty);
            AppendForm(sb, "Url", string.Empty);
            AppendForm(sb, "PasswordTextBox", password);
            AppendForm(sb, "ImageButton1.x", "21");
            AppendForm(sb, "ImageButton1.y", "10");

            return sb.ToString();
        }

        private static void AppendForm(StringBuilder sb, string key, string value)
        {
            if (sb.Length > 0)
            {
                sb.Append('&');
            }

            sb.Append(Uri.EscapeDataString(key ?? string.Empty));
            sb.Append('=');
            sb.Append(Uri.EscapeDataString(value ?? string.Empty));
        }

        private static string GetInputValue(string html, string id)
        {
            if (string.IsNullOrEmpty(html) || string.IsNullOrWhiteSpace(id))
            {
                return string.Empty;
            }

            string pattern = "id=\\\"" + Regex.Escape(id) + "\\\"[^>]*value=\\\"([^\\\"]*)\\\"";
            Match m = Regex.Match(html, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return m.Success ? WebUtility.HtmlDecode(m.Groups[1].Value) : string.Empty;
        }

        private static bool ContainsToken(string html, string token)
        {
            return !string.IsNullOrWhiteSpace(html)
                   && !string.IsNullOrWhiteSpace(token)
                   && html.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void ApplyCookiesToIe(Uri baseUri, CookieContainer cookieContainer)
        {
            if (baseUri == null || cookieContainer == null)
            {
                return;
            }

            CookieCollection cookies = cookieContainer.GetCookies(baseUri);
            foreach (Cookie cookie in cookies)
            {
                if (cookie == null || string.IsNullOrWhiteSpace(cookie.Name))
                {
                    continue;
                }

                string url = baseUri.Scheme + "://" + baseUri.Host + (cookie.Path ?? "/");
                string cookieData = cookie.Name + "=" + cookie.Value + "; path=" + (cookie.Path ?? "/");
                if (!cookie.Expires.Equals(DateTime.MinValue))
                {
                    cookieData += "; expires=" + cookie.Expires.ToUniversalTime().ToString("R");
                }

                bool ok = InternetSetCookie(url, null, cookieData);
                WriteUiLog(ok
                    ? "Cookie aplicada a IE: " + cookie.Name
                    : "No se pudo aplicar cookie a IE: " + cookie.Name);
            }
        }

        private void TryAutoLogin()
        {
            if (_autoLoginAttempted || string.IsNullOrWhiteSpace(_username) || string.IsNullOrWhiteSpace(_password))
            {
                if (string.IsNullOrWhiteSpace(_username) || string.IsNullOrWhiteSpace(_password))
                {
                    if (!_missingCredentialsLogged)
                    {
                        _missingCredentialsLogged = true;
                        WriteUiLog("No hay credenciales para autologin. Se detiene el intento automatico.");
                    }
                    _autoLoginTimer.Stop();
                }
                return;
            }

            _autoLoginTries++;
            WriteUiLog("TryAutoLogin intento #" + _autoLoginTries);

            if (webBrowser1.Document == null)
            {
                WriteUiLog("Documento nulo.");
                return;
            }

            if (TryAutoLoginInDocument(webBrowser1.Document))
            {
                return;
            }

            try
            {
                if (webBrowser1.Document.Window == null || webBrowser1.Document.Window.Frames == null)
                {
                    return;
                }

                foreach (HtmlWindow frame in webBrowser1.Document.Window.Frames)
                {
                    HtmlDocument doc = frame.Document;
                    if (doc != null && TryAutoLoginInDocument(doc))
                    {
                        return;
                    }
                }
            }
            catch
            {
                WriteUiLog("No se pudo inspeccionar frames (restriccion del navegador o DOM). ");
            }
        }

        private bool TryAutoLoginInDocument(HtmlDocument document)
        {
            if (document == null) return false;

            HtmlElement userInput = FindElementByIdOrName(document, "UserNameTextBox");
            HtmlElement passInput = FindElementByIdOrName(document, "PasswordTextBox");

            // Fallback: buscar campo type=password para pagina SSO con IDs distintos
            if (passInput == null)
            {
                foreach (HtmlElement inp in document.GetElementsByTagName("input"))
                {
                    if (string.Equals(inp.GetAttribute("type"), "password", StringComparison.OrdinalIgnoreCase))
                    { passInput = inp; break; }
                }
            }

            // Paso 1: LogOnForm con solo usuario - inyectar override + click flecha
            if (userInput != null && passInput == null && !_step1Done)
            {
                _step1Done = true;
                WriteUiLog("AutoLogin Paso 1: LogOnForm usuario-solo. Inyectando override y clickando flecha.");
                InjectWindowOpenOverrideManaged(document);
                userInput.SetAttribute("value", _username);
                try
                {
                    string eu = EscapeJs(_username);
                    document.InvokeScript("eval", new object[]
                    {
                        "(function(){" +
                        "function setIg(id,v){" +
                        "  var c=null;" +
                        "  if(window.igtxt_getById){try{c=igtxt_getById(id);}catch(e){}}" +
                        "  if(!c&&window[id]&&typeof window[id].SetText==='function'){c=window[id];}" +
                        "  if(c&&c.SetText){c.SetText(v);return;}" +
                        "  var el=document.getElementById(id+'_I')||document.getElementById(id);" +
                        "  if(el){el.value=v;}}" +
                        "setIg('UserNameTextBox','" + eu + "');" +
                        "var a=document.getElementById('Automatico'); if(a){a.value='1';}" +
                        "})()"
                    });
                }
                catch { }
                HtmlElement btn = document.GetElementById("ImageButton1");
                if (btn != null) { btn.InvokeMember("click"); return true; }
                return false;
            }

            if (userInput == null || passInput == null) return false;

            // Paso 2: Pagina SSO completa con usuario + contrasena
            WriteUiLog("AutoLogin Paso 2: pagina SSO con user+pass.");
            FillAndSubmit(document);
            _autoLoginAttempted = true;
            _autoLoginTimer.Stop();
            _loginSubmitDetected = true;
            _loginSubmitUtc = DateTime.UtcNow;
            _authOutcomeChecks = 0;
            _authOutcomeTimer.Start();
            return true;
        }

        private void EvaluateAuthOutcome(Uri currentUrl, HtmlDocument document)
        {
            if (!_loginSubmitDetected || _authOutcomeResolved)
            {
                return;
            }

            // Whoami es un paso intermedio valido del SSO; resetear reloj y esperar
            if (currentUrl != null && currentUrl.AbsoluteUri.IndexOf("Whoami", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                WriteUiLog("Whoami intermedio detectado. Reseteando reloj de autenticacion.");
                _loginSubmitUtc = DateTime.UtcNow;
                _alternateStrategyTried = false;
                return;
            }

            bool isLoginPage = IsLoginPage(currentUrl, document);
            if (!isLoginPage)
            {
                // Cualquier pagina que no sea login y no sea LogOnForm cuenta como exito
                _authOutcomeResolved = true;
                WriteUiLog("Autenticacion visual completada: " + (currentUrl?.AbsoluteUri ?? "(url nula)"));
                return;
            }

            if ((DateTime.UtcNow - _loginSubmitUtc).TotalSeconds < 3.0)
            {
                return;
            }

            if (!_alternateStrategyTried && document != null)
            {
                _alternateStrategyTried = true;
                WriteUiLog("Rebote a login detectado. Reintentando con estrategia alterna (CambiarLink/FlagDominio). ");
                TryAlternateLoginStrategy(document);
                _loginSubmitUtc = DateTime.UtcNow;
                _authOutcomeResolved = false;
                return;
            }

            _authOutcomeResolved = true;
            WriteUiLog("Autenticacion visual fallida: el sistema regreso a LogOnForm.");
            MessageBox.Show(
                "El autologin envio credenciales, pero Panacea devolvio de nuevo la pantalla de inicio de sesion.\n\n" +
                "Verifica usuario/clave y politicas de autenticacion del servidor.",
                "Panacea - Login fallido",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private static bool IsLoginPage(Uri url, HtmlDocument document)
        {
            if (url != null && url.AbsoluteUri.IndexOf("LogOnForm.aspx", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (document == null)
            {
                return false;
            }

            bool hasUser = FindElementByIdOrName(document, "UserNameTextBox") != null;
            bool hasPass = FindElementByIdOrName(document, "PasswordTextBox") != null;
            return hasUser && hasPass;
        }

        private void FillAndSubmit(HtmlDocument document)
        {
            string js = BuildAutoLoginScript(_username, _password);
            try
            {
                object result = document.InvokeScript("eval", new object[] { js });
                string resultText = result == null ? string.Empty : result.ToString();
                WriteUiLog("Resultado script autologin: " + (string.IsNullOrWhiteSpace(resultText) ? "(null)" : resultText));

                if (!string.Equals(resultText, "CLICKED", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(resultText, "FORM_SUBMIT", StringComparison.OrdinalIgnoreCase))
                {
                    WriteUiLog("Script sin confirmacion de envio. Ejecutando fallback DOM.");
                    FillAndSubmitDomFallback(document);
                }
            }
            catch
            {
                WriteUiLog("Fallo InvokeScript(eval). Intentando fallback con DOM nativo.");
                FillAndSubmitDomFallback(document);
            }
        }

        private void FillAndSubmitDomFallback(HtmlDocument document)
        {
            HtmlElement userInput = FindElementByIdOrName(document, "UserNameTextBox");
            HtmlElement passInput = FindElementByIdOrName(document, "PasswordTextBox");
            if (userInput != null)
            {
                userInput.SetAttribute("value", _username);
            }

            if (passInput != null)
            {
                passInput.SetAttribute("value", _password);
            }

            HtmlElement automatico = FindElementByIdOrName(document, "Automatico");
            if (automatico != null)
            {
                automatico.SetAttribute("value", "1");
            }

            HtmlElement submit = FindElementByIdOrName(document, "ImageButton1");
            if (submit == null)
            {
                submit = FindElementByIdOrName(document, "InicioRedV2Button");
            }
            if (submit != null)
            {
                submit.InvokeMember("click");
                WriteUiLog("Fallback: click en boton de ingreso.");
                return;
            }

            foreach (HtmlElement form in document.GetElementsByTagName("form"))
            {
                form.InvokeMember("submit");
                WriteUiLog("Fallback: submit() del primer form.");
                return;
            }

            WriteUiLog("Fallback DOM: no se encontro boton ni form para submit.");
        }

        private void TryAlternateLoginStrategy(HtmlDocument document)
        {
            try
            {
                string js = BuildAlternateLoginScript(_username, _password);
                object result = document.InvokeScript("eval", new object[] { js });
                WriteUiLog("Resultado estrategia alterna: " + (result == null ? "(null)" : result.ToString()));
            }
            catch (Exception ex)
            {
                WriteUiLog("Fallo estrategia alterna: " + ex.Message);
                FillAndSubmitDomFallback(document);
            }
        }

        private static HtmlElement FindElementByIdOrName(HtmlDocument document, string value)
        {
            if (document == null || string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            HtmlElement byId = document.GetElementById(value);
            if (byId != null)
            {
                return byId;
            }

            foreach (HtmlElement element in document.All)
            {
                string name = element.GetAttribute("name");
                if (string.Equals(name, value, StringComparison.OrdinalIgnoreCase))
                {
                    return element;
                }
            }

            return null;
        }

        private static string BuildAutoLoginScript(string username, string password)
        {
            string escapedUser = EscapeJs(username);
            string escapedPass = EscapeJs(password);

            // Infragistics WebTextEditor: visible input = baseId_I, JS control via igtxt_getById/window[baseId]
            return "(function(){"
                + "function findInput(baseId){"
                + "  var el=document.getElementById(baseId+'_I');"
                + "  if(el&&el.tagName==='INPUT')return el;"
                + "  el=document.getElementById(baseId);"
                + "  if(el&&el.tagName==='INPUT')return el;"
                + "  var byName=document.getElementsByName(baseId);"
                + "  for(var i=0;i<byName.length;i++){if(byName[i].tagName==='INPUT')return byName[i];}"
                + "  return null;}"
                + "function igSet(baseId,val){"
                + "  var ctrl=null;"
                + "  if(window.igtxt_getById){try{ctrl=igtxt_getById(baseId);}catch(e){}}"
                + "  if(!ctrl&&window.$find){try{ctrl=$find(baseId);}catch(e){}}"
                + "  if(!ctrl&&window[baseId]&&typeof window[baseId].SetText==='function'){ctrl=window[baseId];}"
                + "  if(ctrl&&typeof ctrl.SetText==='function'){ctrl.SetText(val);}"
                + "  var inp=findInput(baseId);"
                + "  if(inp){inp.value=val;"
                + "    try{inp.dispatchEvent(new Event('input',{bubbles:true}));inp.dispatchEvent(new Event('change',{bubbles:true}));}catch(e){}}}"
                + "var u=findInput('UserNameTextBox'); var p=findInput('PasswordTextBox');"
                + "if(!u||!p){return 'MISSING_FIELDS';}"
                + "igSet('UserNameTextBox','" + escapedUser + "');"
                + "igSet('PasswordTextBox','" + escapedPass + "');"
                + "var a=document.getElementById('Automatico'); if(a){a.value='1';}"
                + "var flg=document.getElementById('FlagDominio'); if(flg){flg.value='0';}"
                + "var btn=document.getElementById('ImageButton1')||document.getElementById('InicioRedV2Button')||document.querySelector('input[type=\\\"image\\\"],input[type=\\\"submit\\\"],button[type=\\\"submit\\\"]');"
                + "if(btn&&btn.click){btn.click(); return 'CLICKED';}"
                + "var f=document.forms&&document.forms.length?document.forms[0]:null; if(f&&f.submit){f.submit(); return 'FORM_SUBMIT';}"
                + "return 'NO_SUBMIT_TARGET';"
                + "})();";
        }

            private static string BuildAlternateLoginScript(string username, string password)
            {
                string escapedUser = EscapeJs(username);
                string escapedPass = EscapeJs(password);

                return "(function(){"
                + "function by(id){return document.getElementById(id)||document.querySelector('[name=\\\"'+id+'\\\"]');}"
                + "var cambiar=by('CambiarLink'); if(cambiar&&cambiar.click){cambiar.click();}"
                + "var flag=by('FlagDominio'); if(flag){flag.value='0';}"
                + "var alt=by('UserAlterno'); if(alt){alt.value='';}"
                + "var aut=by('Automatico'); if(aut){aut.value='1';}"
                + "if(window.UserNameTextBox&&UserNameTextBox.SetText){UserNameTextBox.SetText('" + escapedUser + "');}"
                + "if(window.PasswordTextBox&&PasswordTextBox.SetText){PasswordTextBox.SetText('" + escapedPass + "');}"
                + "var u=by('UserNameTextBox'); var p=by('PasswordTextBox');"
                + "if(u){u.value='" + escapedUser + "';} if(p){p.value='" + escapedPass + "';}"
                + "var btn=by('ImageButton1')||by('InicioRedV2Button')||document.querySelector('input[type=\\\"image\\\"],input[type=\\\"submit\\\"],button[type=\\\"submit\\\"]');"
                + "if(btn&&btn.click){btn.click(); return 'ALT_CLICKED';}"
                + "var f=document.forms&&document.forms.length?document.forms[0]:null; if(f&&f.submit){f.submit(); return 'ALT_FORM_SUBMIT';}"
                + "return 'ALT_NO_TARGET';"
                + "})();";
            }

        private static string EscapeJs(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value
                .Replace("\\", "\\\\")
                .Replace("'", "\\'")
                .Replace("\r", "")
                .Replace("\n", "");
        }

        private bool _sedesScanned;
        private bool _step1Done;
        private bool _sedeSelectionStarted;
        private bool _sedeSelectionDone;
        private int _sedeStep;
        private HtmlDocument _sedeFrameDoc;

        // Escanea el DOM de la pagina de sedes usando HtmlDocument (API managed, sin eval/COM)
        // Se activa cuando detecta texto "SEDE" en el body del WebBrowser embebido
        private void ScanSedesPage(HtmlDocument document, string urlStr)
        {
            if (_sedesScanned || document == null) return;
            if (urlStr.IndexOf("LogOnForm", StringComparison.OrdinalIgnoreCase) >= 0) return;

            // Verificar si hay contenido de sedes en el body
            string bodyText = string.Empty;
            try { bodyText = document.Body?.InnerText ?? string.Empty; } catch { return; }
            if (bodyText.IndexOf("SEDE", StringComparison.OrdinalIgnoreCase) < 0) return;

            _sedesScanned = true;
            WriteUiLog("=== ESCANEO DOM SEDES (HtmlDocument) url=" + urlStr + " ===");

            // 1. Todas las celdas <td> con texto (filas de la tabla de sedes)
            try
            {
                HtmlElementCollection tds = document.GetElementsByTagName("td");
                var sbTd = new StringBuilder();
                foreach (HtmlElement td in tds)
                {
                    string t = (td.InnerText ?? string.Empty).Trim();
                    if (t.Length > 0 && t.Length < 100)
                        sbTd.Append("[id=").Append(td.Id ?? "").Append(" txt=").Append(t).Append("] ");
                }
                WriteUiLog("TDs: " + sbTd);
            }
            catch (Exception ex) { WriteUiLog("TDs error: " + ex.Message); }

            // 2. Todos los <select> con sus opciones
            try
            {
                HtmlElementCollection sels = document.GetElementsByTagName("select");
                foreach (HtmlElement sel in sels)
                {
                    string sid = sel.Id ?? string.Empty;
                    string sname = sel.GetAttribute("name") ?? string.Empty;
                    var sbOpts = new StringBuilder();
                    HtmlElementCollection opts = sel.GetElementsByTagName("option");
                    foreach (HtmlElement opt in opts)
                    {
                        string ov = opt.GetAttribute("value") ?? string.Empty;
                        string ot = (opt.InnerText ?? string.Empty).Trim();
                        sbOpts.Append("[v=").Append(ov).Append(",t=").Append(ot).Append("] ");
                    }
                    WriteUiLog("SELECT id='" + sid + "' name='" + sname + "' opciones: " + sbOpts);
                }
            }
            catch (Exception ex) { WriteUiLog("SELECTs error: " + ex.Message); }

            // 3. Todos los <input> visibles
            try
            {
                HtmlElementCollection inputs = document.GetElementsByTagName("input");
                var sbInp = new StringBuilder();
                foreach (HtmlElement inp in inputs)
                {
                    string itype = inp.GetAttribute("type") ?? string.Empty;
                    if (string.Equals(itype, "hidden", StringComparison.OrdinalIgnoreCase)) continue;
                    string iid = inp.Id ?? string.Empty;
                    string iname = inp.GetAttribute("name") ?? string.Empty;
                    string ival = inp.GetAttribute("value") ?? string.Empty;
                    sbInp.Append("[type=").Append(itype).Append(" id=").Append(iid)
                         .Append(" name=").Append(iname).Append(" val=").Append(ival).Append("] ");
                }
                WriteUiLog("INPUTs: " + sbInp);
            }
            catch (Exception ex) { WriteUiLog("INPUTs error: " + ex.Message); }

            // 4. Todos los <button>
            try
            {
                HtmlElementCollection btns = document.GetElementsByTagName("button");
                var sbBtns = new StringBuilder();
                foreach (HtmlElement btn in btns)
                    sbBtns.Append("[id=").Append(btn.Id ?? "").Append(" txt=").Append((btn.InnerText ?? "").Trim()).Append("] ");
                WriteUiLog("BUTTONs: " + sbBtns);
            }
            catch (Exception ex) { WriteUiLog("BUTTONs error: " + ex.Message); }

            // 5. Todos los <a> con texto corto (posibles botones de accion)
            try
            {
                HtmlElementCollection links = document.GetElementsByTagName("a");
                var sbLinks = new StringBuilder();
                foreach (HtmlElement lnk in links)
                {
                    string lt = (lnk.InnerText ?? string.Empty).Trim();
                    if (lt.Length > 0 && lt.Length < 40)
                        sbLinks.Append("[id=").Append(lnk.Id ?? "").Append(" href=")
                               .Append(lnk.GetAttribute("href") ?? "").Append(" txt=").Append(lt).Append("] ");
                }
                WriteUiLog("LINKs: " + sbLinks);
            }
            catch (Exception ex) { WriteUiLog("LINKs error: " + ex.Message); }

            WriteUiLog("=== FIN ESCANEO DOM SEDES ===");
        }

        private void WriteUiLog(string message)
        {
            try
            {
                string dir = Path.GetDirectoryName(_uiLogPath);
                if (!string.IsNullOrWhiteSpace(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var sb = new StringBuilder();
                sb.Append("[").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("] ").Append(message);
                File.AppendAllText(_uiLogPath, sb.ToString() + Environment.NewLine);
            }
            catch
            {
            }
            AppendSidebarLog(message);
        }

        private bool TryClassicIeAutoLogin()
        {
            if (string.IsNullOrWhiteSpace(_username) || string.IsNullOrWhiteSpace(_password))
            {
                WriteUiLog("IE clasico: no hay credenciales para intentar login.");
                return false;
            }

            try
            {
                Type ieType = Type.GetTypeFromProgID("InternetExplorer.Application");
                if (ieType == null)
                {
                    WriteUiLog("IE clasico: ProgID InternetExplorer.Application no disponible.");
                    return false;
                }

                dynamic ie = Activator.CreateInstance(ieType);
                ie.Visible = true;
                ie.MenuBar = false;
                ie.ToolBar = false;
                ie.AddressBar = false;
                ie.Navigate("http://181.51.196.194" + LoginPath);

                if (!WaitIeReady(ie, 45000))
                {
                    WriteUiLog("IE clasico: timeout esperando carga de login.");
                    return false;
                }

                bool submitted = SubmitLoginInClassicIe(ie);
                WriteUiLog(submitted
                    ? "IE clasico: credenciales enviadas." : "IE clasico: no se pudieron enviar credenciales.");
                if (!submitted)
                    return false;

                // Give IE a moment to start navigating after submit click
                Thread.Sleep(1500);
                // Wait for post-login navigation to Mainframe (up to 40s)
                WriteUiLog("IE clasico: esperando pagina post-login...");
                SelectSedeAndContingencia(ie);
                try
                {
                    string finalUrl = (string)ie.LocationURL ?? string.Empty;
                    bool stillOnLogin = finalUrl.IndexOf("LogOnForm.aspx", StringComparison.OrdinalIgnoreCase) >= 0;
                    if (stillOnLogin)
                    {
                        WriteUiLog("IE clasico: rebote detectado a LogOnForm. Se habilita fallback.");
                        return false;
                    }
                }
                catch
                {
                }

                return true;
            }
            catch (Exception ex)
            {
                WriteUiLog("IE clasico fallo: " + ex.Message);
                return false;
            }
        }

        // Igual que el scanner: inyecta override de window.open en el doc COM
        private static void InjectWindowOpenOverrideCom(dynamic doc)
        {
            try
            {
                dynamic win = doc.parentWindow;
                if (win != null)
                {
                    string js = "(function(){"
                        + "if(window.__popupOverriden)return;"
                        + "window.__popupOverriden=true;"
                        + "window.open=function(url,n,f){if(url&&url!=''){window.location.href=url;}return null;};"
                        + "})();";
                    win.execScript(js, "JavaScript");
                }
            }
            catch { }
        }

        private bool SubmitLoginInClassicIe(dynamic ie)
        {
            // Panacea tiene un flujo SSO de 2 pasos:
            //   Paso 1 - LogOnForm.aspx: solo campo usuario, boton ImageButton1 llama window.open() -> nueva ventana SSO
            //   Paso 2 - SSO page: tiene usuario + contrasena
            // Solucion: inyectar override de window.open para redirigir en la misma ventana,
            //           luego detectar cada paso y actuar.

            DateTime until = DateTime.Now.AddSeconds(60);
            bool step1Done = false;

            while (DateTime.Now < until)
            {
                try
                {
                    bool busy = (bool)ie.Busy;
                    int ready = (int)ie.ReadyState;
                    if (busy || ready != 4) { Thread.Sleep(400); continue; }

                    dynamic doc = ie.Document;
                    if (doc == null) { Thread.Sleep(400); continue; }

                    string url = (string)ie.LocationURL ?? string.Empty;
                    WriteUiLog("SubmitLogin URL: " + url);

                    // Inyectar override de window.open en cada pagina para evitar segunda ventana
                    InjectWindowOpenOverrideCom(doc);

                    // Detectar campos
                    dynamic userEl = FindDynamicElement(doc, new[] { "UserNameTextBox_I", "UserNameTextBox" });
                    dynamic passEl = FindDynamicElement(doc, new[] { "PasswordTextBox_I", "PasswordTextBox" });

                    // Fallback para campos tipo password (SSO puede tener IDs distintos)
                    if (passEl == null)
                    {
                        try
                        {
                            dynamic inputs = doc.getElementsByTagName("input");
                            int ilen = (int)inputs.length;
                            for (int ii = 0; ii < ilen; ii++)
                            {
                                dynamic inp = inputs.item(ii);
                                string itype = string.Empty;
                                try { itype = (string)inp.type ?? string.Empty; } catch { }
                                if (string.Equals(itype, "password", StringComparison.OrdinalIgnoreCase))
                                { passEl = inp; break; }
                            }
                        }
                        catch { }
                    }

                    bool hasUser = userEl != null;
                    bool hasPass = passEl != null;
                    bool isLogonForm = url.IndexOf("LogOnForm", StringComparison.OrdinalIgnoreCase) >= 0;

                    // --- PASO 1: LogOnForm solo con username ---
                    if (hasUser && !hasPass && isLogonForm && !step1Done)
                    {
                        WriteUiLog("Paso 1: LogOnForm - llenando username y haciendo click en flecha");
                        string eu = EscapeJs(_username);
                        try
                        {
                            dynamic win = doc.parentWindow;
                            if (win != null)
                            {
                                string jsFill = "(function(){"
                                    + "function setIg(id,v){"
                                    + "  var c=null;"
                                    + "  if(window.igtxt_getById){try{c=igtxt_getById(id);}catch(e){}}"
                                    + "  if(!c&&window[id]&&typeof window[id].SetText==='function'){c=window[id];}"
                                    + "  if(c&&c.SetText){c.SetText(v);return;}"
                                    + "  var el=document.getElementById(id+'_I')||document.getElementById(id);"
                                    + "  if(el){el.value=v;}"
                                    + "}"
                                    + "setIg('UserNameTextBox','" + eu + "');"
                                    + "var a=document.getElementById('Automatico'); if(a){a.value='1';}"
                                    + "})();";
                                win.execScript(jsFill, "JavaScript");
                            }
                        }
                        catch { }
                        try { userEl.value = _username; } catch { }

                        dynamic btn1 = doc.getElementById("ImageButton1");
                        if (btn1 != null)
                        {
                            btn1.click();
                            step1Done = true;
                            WriteUiLog("Paso 1: click en ImageButton1 ejecutado. Esperando pagina SSO...");
                            Thread.Sleep(1200);
                            continue;
                        }
                    }

                    // --- PASO 2: Pagina SSO con usuario + contrasena ---
                    if (hasUser && hasPass)
                    {
                        WriteUiLog("Paso 2: pagina SSO con user+pass - url=" + url);
                        string eu = EscapeJs(_username);
                        string ep = EscapeJs(_password);

                        try
                        {
                            dynamic win = doc.parentWindow;
                            if (win != null)
                            {
                                string jsFill2 = "(function(){"
                                    + "function setIg(id,v){"
                                    + "  var c=null;"
                                    + "  if(window.igtxt_getById){try{c=igtxt_getById(id);}catch(e){}}"
                                    + "  if(!c&&window[id]&&typeof window[id].SetText==='function'){c=window[id];}"
                                    + "  if(c&&c.SetText){c.SetText(v);return;}"
                                    + "  var el=document.getElementById(id+'_I')||document.getElementById(id);"
                                    + "  if(el){el.value=v;}"
                                    + "  var all=document.getElementsByTagName('input');"
                                    + "  for(var i=0;i<all.length;i++){var n=all[i].name||'';"
                                    + "    if(n.indexOf(id)===0&&(n===id||n.indexOf('CVS')>0)){all[i].value=v;}}"
                                    + "}"
                                    + "setIg('UserNameTextBox','" + eu + "');"
                                    + "setIg('PasswordTextBox','" + ep + "');"
                                    + "var a=document.getElementById('Automatico'); if(a){a.value='1';}"
                                    + "var f=document.getElementById('FlagDominio'); if(f){f.value='0';}"
                                    + "})();";
                                win.execScript(jsFill2, "JavaScript");
                            }
                        }
                        catch { }

                        try { userEl.value = _username; } catch { }
                        try { passEl.value = _password; } catch { }

                        // CVS hidden fields
                        try
                        {
                            dynamic allInputs = doc.getElementsByTagName("input");
                            int alen = (int)allInputs.length;
                            for (int ai = 0; ai < alen; ai++)
                            {
                                dynamic el = allInputs.item(ai);
                                string en = string.Empty;
                                try { en = (string)el.name ?? string.Empty; } catch { }
                                if (en == "UserNameTextBox$CVS") try { el.value = _username; } catch { }
                                else if (en == "PasswordTextBox$CVS") try { el.value = _password; } catch { }
                            }
                        }
                        catch { }

                        // Agregar coordenadas ImageButton1 y hacer submit
                        try
                        {
                            dynamic win2 = doc.parentWindow;
                            if (win2 != null)
                            {
                                string jsCoords = "(function(){"
                                    + "var f=document.forms[0];if(!f)return;"
                                    + "function addH(n,v){var e=document.getElementsByName(n);"
                                    + "  if(e.length>0){e[0].value=v;return;}"
                                    + "  var h=document.createElement('input');h.type='hidden';h.name=n;h.value=v;f.appendChild(h);}"
                                    + "addH('ImageButton1.x','21');addH('ImageButton1.y','10');"
                                    + "})();";
                                win2.execScript(jsCoords, "JavaScript");
                            }
                        }
                        catch { }

                        dynamic submit = doc.getElementById("ImageButton1");
                        if (submit == null) submit = doc.getElementById("InicioRedV2Button");
                        if (submit == null) submit = doc.getElementById("btnIngresar");
                        if (submit == null) submit = doc.getElementById("LoginButton");

                        if (submit != null)
                        {
                            WriteUiLog("Paso 2: submit click en id=" + SafeGetId(submit));
                            submit.click();
                            return true;
                        }

                        // Fallback form submit
                        try
                        {
                            dynamic forms = doc.forms;
                            if (forms != null && forms.length > 0)
                            {
                                forms.item(0).submit();
                                WriteUiLog("Paso 2: form.submit() ejecutado.");
                                return true;
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                Thread.Sleep(500);
            }

            WriteUiLog("SubmitLoginInClassicIe: timeout sin enviar credenciales.");
            return false;
        }

        private static dynamic FindDynamicElement(dynamic doc, string[] ids)
        {
            foreach (string id in ids)
            {
                try
                {
                    dynamic el = doc.getElementById(id);
                    if (el != null) return el;
                }
                catch { }
                try
                {
                    dynamic els = doc.getElementsByName(id);
                    if (els != null && els.length > 0) return els.item(0);
                }
                catch { }
            }
            return null;
        }

        private static string SafeGetId(dynamic el)
        {
            try { return (string)el.id ?? string.Empty; } catch { return string.Empty; }
        }

        private string DumpFormFieldsDynamic(dynamic doc)
        {
            var sb = new StringBuilder();
            try
            {
                dynamic inputs = doc.getElementsByTagName("input");
                if (inputs != null)
                {
                    int len = inputs.length;
                    for (int i = 0; i < len && i < 30; i++)
                    {
                        dynamic el = inputs.item(i);
                        string id = string.Empty;
                        string name = string.Empty;
                        string type = string.Empty;
                        try { id = (string)el.id ?? string.Empty; } catch { }
                        try { name = (string)el.name ?? string.Empty; } catch { }
                        try { type = (string)el.type ?? string.Empty; } catch { }
                        sb.Append("[").Append(type).Append(":").Append(id).Append("/").Append(name).Append("]");
                    }
                }
            }
            catch (Exception ex)
            {
                sb.Append("ERR:" + ex.Message);
            }
            return sb.ToString();
        }

        // Ejecuta JS via execScript (compatible con IE COM) y retorna valor via variable global __pbot_result
        private static string IeExecJs(dynamic doc, string jsExpression)
        {
            try
            {
                dynamic win = doc.parentWindow;
                if (win == null) return string.Empty;
                // Escribe resultado en variable global para poder leerlo luego
                string wrapped = "window.__pbot_result=(" + jsExpression + ");";
                win.execScript(wrapped, "JavaScript");
                object val = doc.parentWindow.__pbot_result;
                return val == null ? string.Empty : val.ToString();
            }
            catch { return string.Empty; }
        }

        private void SelectSedeAndContingencia(dynamic ie)
        {
            // Esperar inicio de navegacion post-login
            DateTime busyWait = DateTime.Now.AddSeconds(5);
            while (DateTime.Now < busyWait)
            {
                try { if ((bool)ie.Busy) break; } catch { }
                Thread.Sleep(200);
            }

            // Esperar el frame ValidacionesUsuario.aspx (dialogo de sedes)
            // La pagina pasa primero por Whoami y luego por Mainframe.aspx que tiene el iframe
            DateTime mainframeWait = DateTime.Now.AddSeconds(90);
            dynamic validacionesDoc = null;
            while (DateTime.Now < mainframeWait)
            {
                try
                {
                    bool busy = (bool)ie.Busy;
                    int state = (int)ie.ReadyState;
                    string url = (string)ie.LocationURL ?? string.Empty;
                    WriteUiLog("Esperando pagina sedes... busy=" + busy + " state=" + state + " url=" + url);

                    if (!busy && state == 4)
                    {
                        bool isLogin = url.IndexOf("LogOnForm.aspx", StringComparison.OrdinalIgnoreCase) >= 0;
                        bool isWhoami = url.IndexOf("Whoami", StringComparison.OrdinalIgnoreCase) >= 0;

                        if (isLogin || isWhoami)
                        {
                            WriteUiLog("Pagina intermedia: " + url);
                            Thread.Sleep(800);
                            continue;
                        }

                        dynamic found = FindValidacionesFrame(ie.Document);
                        if (found != null)
                        {
                            validacionesDoc = found;
                            WriteUiLog("Frame ValidacionesUsuario.aspx encontrado.");
                            break;
                        }
                    }
                }
                catch { }
                Thread.Sleep(500);
            }

            if (validacionesDoc == null)
            {
                WriteUiLog("SelectSedeAndContingencia: no se encontro el frame de sedes en el tiempo esperado.");
                return;
            }

            Thread.Sleep(1500); // Dejar que Infragistics termine de renderizar

            // === PASO 1: CLICK EN FILA SEDE ADMINISTRATIVA CALLE 74 ===
            WriteUiLog("=== PASO 1: Seleccionar sede ===");
            try
            {
                string r = IeExecJs(validacionesDoc,
                    "(function(){" +
                    "  var trs=document.getElementsByTagName('tr');" +
                    "  for(var i=0;i<trs.length;i++){" +
                    "    var t=(trs[i].innerText||'').toUpperCase();" +
                    "    if(t.indexOf('SEDE ADMINISTRATIVA CALLE 74')>=0){trs[i].click();return 'OK:'+trs[i].id;}" +
                    "  }" +
                    "  var r0=document.getElementById('SedesDatagrid_r_0');" +
                    "  if(r0){r0.click();return 'OK-fallback:SedesDatagrid_r_0';}" +
                    "  return 'NOT_FOUND';" +
                    "})()");
                WriteUiLog("Sede click: " + r);
                if (!r.StartsWith("OK")) { WriteUiLog("ERROR: sede no encontrada."); return; }
            }
            catch (Exception ex) { WriteUiLog("ERROR sede: " + ex.Message); return; }

            Thread.Sleep(1200);

            // === PASO 2: SELECCIONAR CONTINGENCIA FACTURACION (DevExpress ComboBox) ===
            WriteUiLog("=== PASO 2: Seleccionar contingencia FACTURACION ===");
            try
            {
                // Intentar API cliente DevExpress primero
                string r1 = IeExecJs(validacionesDoc,
                    "(function(){" +
                    "  try{var cb=window['ContingenciasComboBox'];if(cb&&cb.SetText){cb.SetText('FACTURACION');return 'DX-API';}}" +
                    "  catch(e){}" +
                    "  var btn=document.getElementById('ContingenciasComboBox_B-1');" +
                    "  if(btn){btn.click();return 'DROPDOWN_OPEN';}" +
                    "  return 'BTN_NOT_FOUND';" +
                    "})()");
                WriteUiLog("Contingencia paso1: " + r1);

                if (r1 != "DX-API")
                {
                    Thread.Sleep(800);
                    // Click en el item FACTURACION (LBI1T0 = index 1 de la lista)
                    string r2 = IeExecJs(validacionesDoc,
                        "(function(){" +
                        "  var item=document.getElementById('ContingenciasComboBox_DDD_L_LBI1T0');" +
                        "  if(item){item.click();return 'ITEM_CLICKED';}" +
                        "  var tds=document.getElementsByTagName('td');" +
                        "  for(var i=0;i<tds.length;i++){" +
                        "    if((tds[i].innerText||'').toUpperCase().indexOf('FACTURACION')>=0)" +
                        "      {tds[i].click();return 'TEXT_CLICKED:'+tds[i].id;}" +
                        "  }" +
                        "  return 'ITEM_NOT_FOUND';" +
                        "})()");
                    WriteUiLog("Contingencia paso2: " + r2);
                }
            }
            catch (Exception ex) { WriteUiLog("ERROR contingencia: " + ex.Message); }

            Thread.Sleep(800);

            // === PASO 3: CLICK ACEPTAR ===
            WriteUiLog("=== PASO 3: Click Aceptar ===");
            try
            {
                string r = IeExecJs(validacionesDoc,
                    "(function(){" +
                    "  var inputs=document.getElementsByTagName('input');" +
                    "  for(var i=0;i<inputs.length;i++){" +
                    "    if(inputs[i].name==='AceptarButton'){inputs[i].click();return 'SUBMIT_CLICKED';}" +
                    "  }" +
                    "  var td=document.getElementById('AceptarButton_B');" +
                    "  if(td){td.click();return 'TD_CLICKED';}" +
                    "  var tds=document.getElementsByTagName('td');" +
                    "  for(var i=0;i<tds.length;i++){" +
                    "    if((tds[i].innerText||'').toUpperCase().trim()==='ACEPTAR'){tds[i].click();return 'TEXT_CLICKED:'+tds[i].id;}" +
                    "  }" +
                    "  return 'NOT_FOUND';" +
                    "})()");
                WriteUiLog("Aceptar resultado: " + r);
            }
            catch (Exception ex) { WriteUiLog("ERROR aceptar: " + ex.Message); }

            WriteUiLog("SelectSedeAndContingencia: secuencia completada.");
        }

        private dynamic FindValidacionesFrame(dynamic doc)
        {
            try
            {
                int frameCount = 0;
                try { frameCount = (int)doc.parentWindow.frames.length; } catch { return null; }
                for (int fi = 0; fi < frameCount; fi++)
                {
                    try
                    {
                        dynamic frameWin = doc.parentWindow.frames.item(fi);
                        dynamic frameDoc = frameWin.document;
                        if (frameDoc == null) continue;
                        string furl = string.Empty;
                        try { furl = (string)frameDoc.location.href ?? string.Empty; } catch { }
                        if (furl.IndexOf("ValidacionesUsuario", StringComparison.OrdinalIgnoreCase) >= 0)
                            return frameDoc;
                        // Buscar en sub-frames recursivamente
                        dynamic nested = FindValidacionesFrame(frameDoc);
                        if (nested != null) return nested;
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        private void DomScan(dynamic doc, string context)
        {
            WriteUiLog("--- DomScan [" + context + "] ---");

            // 1. URL de este documento
            try
            {
                string docUrl = string.Empty;
                try { docUrl = (string)doc.location.href ?? string.Empty; } catch { }
                WriteUiLog("[" + context + "] URL: " + docUrl);
            }
            catch { }

            // 2. Todas las celdas <td> con texto no vacio (para encontrar filas de sede)
            try
            {
                dynamic cells = doc.getElementsByTagName("td");
                int len = (int)cells.length;
                var sb2 = new StringBuilder();
                for (int i = 0; i < len; i++)
                {
                    dynamic cell = cells.item(i);
                    string txt = string.Empty;
                    try { txt = ((string)cell.innerText ?? string.Empty).Trim(); } catch { }
                    if (txt.Length > 0 && txt.Length < 120)
                        sb2.Append("[td:").Append(txt).Append("] ");
                }
                WriteUiLog("[" + context + "] TDs: " + sb2);
            }
            catch (Exception ex) { WriteUiLog("[" + context + "] TDs error: " + ex.Message); }

            // 3. Todos los <select> con sus opciones
            try
            {
                dynamic selects = doc.getElementsByTagName("select");
                int len = (int)selects.length;
                for (int i = 0; i < len; i++)
                {
                    dynamic sel = selects.item(i);
                    string sid = string.Empty; string sname = string.Empty;
                    try { sid = (string)sel.id ?? string.Empty; } catch { }
                    try { sname = (string)sel.name ?? string.Empty; } catch { }
                    var sbOpts = new StringBuilder();
                    try
                    {
                        int optLen = (int)sel.options.length;
                        for (int j = 0; j < optLen; j++)
                        {
                            dynamic opt = sel.options.item(j);
                            string oval = string.Empty; string otxt = string.Empty;
                            try { oval = (string)opt.value ?? string.Empty; } catch { }
                            try { otxt = (string)opt.text ?? string.Empty; } catch { }
                            sbOpts.Append("[v=").Append(oval).Append(",t=").Append(otxt).Append("] ");
                        }
                    }
                    catch { }
                    WriteUiLog("[" + context + "] SELECT id='" + sid + "' name='" + sname + "' options: " + sbOpts);
                }
            }
            catch (Exception ex) { WriteUiLog("[" + context + "] SELECTs error: " + ex.Message); }

            // 4. Todos los <input> visibles (type != hidden) con id, name, value
            try
            {
                dynamic inputs = doc.getElementsByTagName("input");
                int len = (int)inputs.length;
                var sbInp = new StringBuilder();
                for (int i = 0; i < len; i++)
                {
                    dynamic inp = inputs.item(i);
                    string itype = string.Empty; string iid = string.Empty;
                    string iname = string.Empty; string ival = string.Empty;
                    try { itype = (string)inp.type ?? string.Empty; } catch { }
                    try { iid = (string)inp.id ?? string.Empty; } catch { }
                    try { iname = (string)inp.name ?? string.Empty; } catch { }
                    try { ival = (string)inp.value ?? string.Empty; } catch { }
                    if (!string.Equals(itype, "hidden", StringComparison.OrdinalIgnoreCase))
                        sbInp.Append("[type=").Append(itype).Append(" id=").Append(iid).Append(" name=").Append(iname).Append(" val=").Append(ival).Append("] ");
                }
                WriteUiLog("[" + context + "] INPUTs: " + sbInp);
            }
            catch (Exception ex) { WriteUiLog("[" + context + "] INPUTs error: " + ex.Message); }

            // 5. Todos los <button> y <a> con texto que parezcan botones
            try
            {
                dynamic buttons = doc.getElementsByTagName("button");
                int len = (int)buttons.length;
                var sbBtns = new StringBuilder();
                for (int i = 0; i < len; i++)
                {
                    dynamic btn = buttons.item(i);
                    string bid = string.Empty; string btxt = string.Empty;
                    try { bid = (string)btn.id ?? string.Empty; } catch { }
                    try { btxt = ((string)btn.innerText ?? string.Empty).Trim(); } catch { }
                    sbBtns.Append("[id=").Append(bid).Append(" txt=").Append(btxt).Append("] ");
                }
                WriteUiLog("[" + context + "] BUTTONs: " + sbBtns);
            }
            catch (Exception ex) { WriteUiLog("[" + context + "] BUTTONs error: " + ex.Message); }
        }

        // Busca la fila de sede en el DOM (C# nativo + frames) y hace click
        private bool ClickSedeRow(dynamic ie, dynamic doc, string targetSede)
        {
            // Intentar en documento principal
            if (ClickSedeInDoc(doc, targetSede)) return true;

            // Intentar en frames
            try
            {
                int frameCount = 0;
                try { frameCount = (int)doc.parentWindow.frames.length; } catch { }
                for (int fi = 0; fi < frameCount; fi++)
                {
                    try
                    {
                        dynamic frameDoc = doc.parentWindow.frames.item(fi).document;
                        if (frameDoc != null && ClickSedeInDoc(frameDoc, targetSede)) return true;
                    }
                    catch { }
                }
            }
            catch { }

            WriteUiLog("Sede no encontrada en DOM todavia.");
            return false;
        }

        private bool ClickSedeInDoc(dynamic doc, string targetSede)
        {
            try
            {
                dynamic cells = doc.getElementsByTagName("td");
                if (cells == null) return false;
                int len = (int)cells.length;
                for (int i = 0; i < len; i++)
                {
                    dynamic cell = cells.item(i);
                    string text = string.Empty;
                    try { text = ((string)cell.innerText ?? string.Empty).Trim().ToUpperInvariant(); } catch { }
                    if (text == targetSede || text.IndexOf(targetSede, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        try { cell.click(); } catch { }
                        WriteUiLog("Sede seleccionada: " + targetSede);
                        return true;
                    }
                }
            }
            catch { }
            return false;
        }

        private static bool WaitIeReady(dynamic ie, int timeoutMs)
        {
            DateTime end = DateTime.Now.AddMilliseconds(timeoutMs);
            while (DateTime.Now < end)
            {
                try
                {
                    bool busy = ie.Busy;
                    int readyState = ie.ReadyState;
                    if (!busy && readyState == 4)
                    {
                        return true;
                    }
                }
                catch
                {
                }

                Thread.Sleep(250);
            }

            return false;
        }

        // -----------------------------------------------------------------------
        // SELECCION AUTOMATICA DE SEDE / CONTINGENCIA / ACEPTAR
        // -----------------------------------------------------------------------

        private void StartSedeSelectionFlow()
        {
            if (_sedeSelectionDone) return;
            _sedeFrameDoc = FindValidacionesFrameManaged(webBrowser1.Document);
            if (_sedeFrameDoc == null)
            {
                WriteUiLog("ValidacionesUsuario frame no encontrado. Reintentando en 1.5s...");
                _sedeSelectionStarted = false;
                return;
            }
            WriteUiLog("Iniciando secuencia sede/contingencia/aceptar con JS.");
            _sedeStep = 0;
            var timer = new System.Windows.Forms.Timer { Interval = 1200 };
            timer.Tick += SedeStepTimer_Tick;
            timer.Start();
        }

        private HtmlDocument FindValidacionesFrameManaged(HtmlDocument doc)
        {
            if (doc == null) return null;
            try
            {
                foreach (HtmlWindow frame in doc.Window.Frames)
                {
                    try
                    {
                        string furl = frame.Url?.ToString() ?? string.Empty;
                        HtmlDocument fd = frame.Document;
                        if (furl.IndexOf("ValidacionesUsuario", StringComparison.OrdinalIgnoreCase) >= 0 && fd != null)
                            return fd;
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        // Ejecuta JS en el frame de ValidacionesUsuario via InvokeScript
        private string JsInFrame(string js)
        {
            if (_sedeFrameDoc == null) return "NO_FRAMEDOC";
            try
            {
                object r = _sedeFrameDoc.InvokeScript("eval", new object[] { js });
                return r == null ? "(null)" : r.ToString();
            }
            catch (Exception ex) { return "EX:" + ex.Message; }
        }

        private void SedeStepTimer_Tick(object sender, EventArgs e)
        {
            var timer = (System.Windows.Forms.Timer)sender;
            _sedeStep++;
            WriteUiLog("SedeStep " + _sedeStep);
            try
            {
                string r;
                switch (_sedeStep)
                {
                    case 1: // Click real OS-level con mouse_event (bypass JS/attachEvent)
                        try {
                            // 1. Posicion del TR dentro del iframe
                            string rowPos = JsInFrame(
                                "(function(){" +
                                "  var rows=document.querySelectorAll('[id^=\\x22SedesDatagrid_r_\\x22]');" +
                                "  var el=null;" +
                                "  for(var i=0;i<rows.length;i++){" +
                                "    var t=(rows[i].innerText||rows[i].textContent||'').toUpperCase();" +
                                "    if(t.indexOf('CALLE 74')>=0){el=rows[i];break;}" +
                                "  }" +
                                "  if(!el){el=document.getElementById('SedesDatagrid_r_0');}" +
                                "  if(!el){return 'NF';}" +
                                "  var r=el.getBoundingClientRect();" +
                                "  var cx=Math.round(r.left+80);" +
                                "  var cy=Math.round(r.top+(r.height/2||10));" +
                                "  return cx+','+cy;" +
                                "})()");
                            WriteUiLog("rowPos en iframe: " + rowPos);
                            if (rowPos == "NF") { WriteUiLog("ROW_NOT_FOUND"); break; }

                            // 2. Posicion del iframe dentro del documento principal
                            string ifrPos = "0,0";
                            try {
                                ifrPos = (string)webBrowser1.Document.InvokeScript("eval", new object[] {
                                    "(function(){" +
                                    "  var all=document.querySelectorAll('iframe,frame');" +
                                    "  for(var i=0;i<all.length;i++){" +
                                    "    try{" +
                                    "      var u=all[i].src||'';" +
                                    "      try{u=all[i].contentWindow.location.href;}catch(e){}" +
                                    "      if(u.indexOf('ValidacionesUsuario')>=0){" +
                                    "        var r=all[i].getBoundingClientRect();" +
                                    "        return Math.round(r.left)+','+Math.round(r.top);" +
                                    "      }" +
                                    "    }catch(e){}" +
                                    "  }" +
                                    "  return '0,0';" +
                                    "})()"
                                }) ?? "0,0";
                            } catch (Exception ex2) { WriteUiLog("ifrPos err: " + ex2.Message); }
                            WriteUiLog("ifrPos en main: " + ifrPos);

                            // 3. Calcular coordenadas de pantalla absolutas
                            var pp = rowPos.Split(',');
                            var ip = ifrPos.Split(',');
                            int rowX = int.Parse(pp[0].Trim()), rowY = int.Parse(pp[1].Trim());
                            int ifrX = int.Parse(ip[0].Trim()), ifrY = int.Parse(ip[1].Trim());
                            var browserPt = webBrowser1.PointToScreen(new System.Drawing.Point(0, 0));
                            int screenX = browserPt.X + ifrX + rowX;
                            int screenY = browserPt.Y + ifrY + rowY;
                            WriteUiLog("Click OS en: " + screenX + "," + screenY);

                            // 4. Enfocar ventana, mover cursor y hacer click real
                            this.Activate();
                            webBrowser1.Focus();
                            System.Threading.Thread.Sleep(200);
                            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(screenX, screenY);
                            System.Threading.Thread.Sleep(150);
                            NativeMouseClick(screenX, screenY);
                            System.Threading.Thread.Sleep(500);

                            // 5. Verificar si Infragistics registró la selección
                            string selCheck = JsInFrame(
                                "(function(){" +
                                "  try{" +
                                "    var g=igtbl_getGridById('SedesDatagrid');" +
                                "    if(g&&g.selectedRow){return 'SEL:'+g.selectedRow.id;}" +
                                "    return 'NO_SEL';" +
                                "  }catch(e){return 'ERR:'+e.message;}" +
                                "})()");
                            WriteUiLog("Sede sel check: " + selCheck);
                        } catch (Exception ex) { WriteUiLog("Sede click err: " + ex.Message); }
                        break;
                    case 2: // Seleccionar FACTURACION directamente via DevExpress API (sin abrir dropdown)
                        r = JsInFrame(
                            "(function(){" +
                            // Estrategia 1: ComboBox por nombre exacto (evita _DDD_L y subcomponentes)
                            "  var cb=window['ContingenciasComboBox'];" +
                            "  if(cb&&typeof cb.GetItemCount==='function'){" +
                            "    for(var i=0;i<cb.GetItemCount();i++){" +
                            "      var it=cb.GetItem(i);" +
                            "      if(it&&(it.text||it.value||'').toUpperCase().indexOf('FACTURACION')>=0){" +
                            "        cb.SetSelectedIndex(i);" +
                            "        return 'SEL_IDX:'+i+':'+it.text;" +
                            "      }" +
                            "    }" +
                            "    return 'NOT_FOUND:'+cb.GetItemCount();" +
                            "  }" +
                            // Estrategia 2: ASPxClientControl collection
                            "  try{var coll=ASPxClientControl.GetControlCollection();" +
                            "    var cb2=coll.GetByName('ContingenciasComboBox');" +
                            "    if(cb2&&typeof cb2.GetItemCount==='function'){" +
                            "      for(var j=0;j<cb2.GetItemCount();j++){" +
                            "        var it2=cb2.GetItem(j);" +
                            "        if(it2&&(it2.text||it2.value||'').toUpperCase().indexOf('FACTURACION')>=0){" +
                            "          cb2.SetSelectedIndex(j);" +
                            "          return 'COLL_IDX:'+j+':'+it2.text;" +
                            "        }" +
                            "      }" +
                            "    }" +
                            "  }catch(e){}" +
                            // Estrategia 3: SetValue directo sobre el ComboBox
                            "  if(cb&&typeof cb.SetValue==='function'){cb.SetValue('FACTURACION');return 'WIN_SV';}" +
                            // Estrategia 4: Abrir dropdown (ultimo recurso)
                            "  var btn=document.getElementById('ContingenciasComboBox_B-1');" +
                            "  if(btn){btn.click();return 'DROPDOWN_OPEN';}" +
                            "  return 'DX-FAIL';" +
                            "})()" );
                        WriteUiLog("Contingencia DX API: " + r);
                        if (r == "DX-FAIL" || r == "DROPDOWN_OPEN" || r == null || r.StartsWith("NOT_FOUND"))
                        {
                            if (r == "DROPDOWN_OPEN")
                            {
                                // Esperar que abra el dropdown y luego hacer click en FACTURACION
                                System.Threading.Thread.Sleep(600);
                                r = JsInFrame(
                                    "(function(){" +
                                    "  var item=document.getElementById('ContingenciasComboBox_DDD_L_LBI1T0');" +
                                    "  if(item){item.click();return 'ITEM_CLICKED';}" +
                                    "  var tds=document.getElementsByTagName('td');" +
                                    "  for(var i=0;i<tds.length;i++){if((tds[i].innerText||'').toUpperCase().indexOf('FACTURACION')>=0){tds[i].click();return 'TEXT_CLICKED:'+tds[i].id;}}" +
                                    "  return 'ITEM_NF';" +
                                    "})()");
                                WriteUiLog("Contingencia dropdown item: " + r);
                            }
                            else
                            {
                                // Fallback: forzar valor directo — usar createEventObject() compatible con IE
                                r = JsInFrame(
                                    "(function(){" +
                                    "  var vi=document.getElementById('ContingenciasComboBox_VI');" +
                                    "  if(vi){vi.value='1';}" +
                                    "  var inp=document.getElementById('ContingenciasComboBox_I');" +
                                    "  if(inp){inp.value='FACTURACION';" +
                                    "    try{inp.fireEvent('onchange');}catch(e){" +
                                    "      try{var ev=document.createEvent('Event');ev.initEvent('change',true,true);inp.dispatchEvent(ev);}catch(e2){}" +
                                    "    }" +
                                    "  }" +
                                    "  return 'DIRECT-SET';" +
                                    "})()");
                                WriteUiLog("Contingencia fallback: " + r);
                            }
                        }
                        break;

                    case 3: // Click Aceptar: serializar Infragistics y loguear hidden field
                        r = JsInFrame(
                            "(function(){" +
                            "  try{if(typeof igtbl_preSubmit==='function'){igtbl_preSubmit();}}catch(e){}" +
                            "  try{if(typeof igtbl_serialize==='function'){igtbl_serialize();}}catch(e){}" +
                            "  try{if(typeof igtbl_serializeGrid==='function'){igtbl_serializeGrid('SedesDatagrid');}}catch(e){}" +
                            "  try{var f=document.forms[0];if(f&&f.onsubmit){f.onsubmit();}}catch(e){}" +
                            "  var hfVal='?';" +
                            "  try{" +
                            "    var hfs=document.getElementsByName('SedesDatagrid');" +
                            "    for(var j=0;j<hfs.length;j++){if(hfs[j].type==='hidden'){hfVal=hfs[j].value.substring(0,60);break;}}" +
                            "  }catch(e){}" +
                            "  var inputs=document.getElementsByTagName('input');" +
                            "  for(var i=0;i<inputs.length;i++){" +
                            "    if(inputs[i].name==='AceptarButton'){inputs[i].click();return 'SUBMIT-INPUT:HF='+hfVal;}" +
                            "  }" +
                            "  var tdB=document.getElementById('AceptarButton_B');" +
                            "  if(tdB){tdB.click();return 'TD-B:HF='+hfVal;}" +
                            "  return 'NOT_FOUND:HF='+hfVal;" +
                            "})()");
                        WriteUiLog("Aceptar: " + r);
                        timer.Stop();
                        _sedeSelectionDone = true;
                        WriteUiLog("Secuencia sede/contingencia/aceptar completada.");
                        break;
                }
            }
            catch (Exception ex)
            {
                WriteUiLog("SedeStep " + _sedeStep + " error: " + ex.Message);
            }
        }

    
        // ─────────────────────────────────────────────────────────────────────
        // LECTURA EXCEL BASE RIPS
        // ─────────────────────────────────────────────────────────────────────

        private void LoadRipsExcel()
        {
            _ripsRecords = new List<RipsRecord>();
            try
            {
                // Ruta del archivo: primero la seleccionada por el usuario, luego la carpeta 'base'
                string xlsxPath = null;
                if (!string.IsNullOrWhiteSpace(_overrideExcelPath) && File.Exists(_overrideExcelPath))
                {
                    xlsxPath = _overrideExcelPath;
                }
                else
                {
                    string baseFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "base");
                    string[] candidates = Directory.Exists(baseFolder)
                        ? Directory.GetFiles(baseFolder, "*.xlsx", SearchOption.TopDirectoryOnly)
                        : new string[0];

                    if (candidates.Length == 0)
                    {
                        WriteUiLog("LoadRipsExcel: no se encontro archivo .xlsx en carpeta 'base'.");
                        return;
                    }
                    xlsxPath = candidates[0];
                }
                WriteUiLog("LoadRipsExcel: leyendo " + xlsxPath);

                using (var pkg = new ExcelPackage(new FileInfo(xlsxPath)))
                {
                    var ws = pkg.Workbook.Worksheets[1];  // EPPlus 4.x es 1-based
                    int lastRow = ws.Dimension?.End.Row ?? 1;

                    // Detectar columnas por cabecera (fila 1)
                    int colCC = -1, colFI = -1, colFF = -1, colConv = -1;
                    int colFinalidad = -1, colDiag = -1, colCausa = -1;
                    int totalCols = ws.Dimension?.End.Column ?? 5;
                    for (int c = 1; c <= totalCols; c++)
                    {
                        string header = (ws.Cells[1, c].Value ?? string.Empty).ToString().Trim().ToUpperInvariant();
                        if (header == "CC") colCC = c;
                        else if (header.Contains("INICIO")) colFI = c;
                        else if (header.Contains("FINALIDAD")) colFinalidad = c;  // ANTES de FIN: "FINALIDAD" contiene "FIN"
                        else if (header.Contains("FIN")) colFF = c;
                        else if (header.Contains("CONVENIO")) colConv = c;
                        else if (header.Contains("DIAGN")) colDiag = c;
                        else if (header.Contains("CAUSA")) colCausa = c;
                    }
                    // Fallback por posicion si no se encontraron por nombre
                    // — evitar asignar FF a una columna ya detectada como FINALIDAD u otra
                    if (colCC  < 0) colCC  = 1;
                    if (colFI  < 0) colFI  = 2;
                    if (colFF  < 0)
                    {
                        // Buscar la primera columna libre que no sea ya CC/FI/FINALIDAD/CONV/DIAG/CAUSA
                        for (int fc = 3; fc <= totalCols; fc++)
                        {
                            if (fc != colCC && fc != colFI && fc != colFinalidad &&
                                fc != colConv && fc != colDiag && fc != colCausa)
                            { colFF = fc; break; }
                        }
                        // Si todas las columnas están ocupadas, dejar -1 (FF vendrá del mes de FI)
                    }
                    if (colConv < 0) colConv = 4;

                    for (int r = 2; r <= lastRow; r++)
                    {
                        string cc = CellToString(ws.Cells[r, colCC].Value);
                        string fi = CellToDate(ws.Cells[r, colFI].Value);
                        // Si no hay columna FF, usar cadena vacía — se calculará como fin del mes de FI
                        string ff = colFF > 0 ? CellToDate(ws.Cells[r, colFF].Value) : string.Empty;
                        // Si FF existe pero no es una fecha válida (ej: "OTRA"), descartar
                        if (!string.IsNullOrWhiteSpace(ff) &&
                            !System.Text.RegularExpressions.Regex.IsMatch(ff, @"^\d{2}-\d{2}-\d{4}$"))
                            ff = string.Empty;
                        string conv = CellToString(ws.Cells[r, colConv].Value);
                        string finalidad = colFinalidad > 0 ? CellToString(ws.Cells[r, colFinalidad].Value) : string.Empty;
                        string diagnostico = colDiag > 0 ? CellToString(ws.Cells[r, colDiag].Value) : string.Empty;
                        string causa = colCausa > 0 ? CellToString(ws.Cells[r, colCausa].Value) : string.Empty;
                        if (string.IsNullOrWhiteSpace(cc)) continue;
                        _ripsRecords.Add(new RipsRecord { CC = cc, FechaInicio = fi, FechaFin = ff, TipoConvenio = conv, Finalidad = finalidad, Diagnostico = diagnostico, CausaExterna = causa });
                        WriteUiLog("  Registro: CC=" + cc + " FI=" + fi + " FF=" + ff + " Conv=" + conv + " Fin=" + finalidad + " Diag=" + diagnostico + " Causa=" + causa);
                    }
                }

                WriteUiLog("LoadRipsExcel: " + _ripsRecords.Count + " registros cargados.");
            }
            catch (Exception ex)
            {
                WriteUiLog("LoadRipsExcel error: " + ex.Message);
            }
        }

        private static string CellToString(object val)
        {
            if (val == null) return string.Empty;
            if (val is double d) return ((long)d).ToString();
            return val.ToString().Trim();
        }

        private static string CellToDate(object val)
        {
            if (val == null) return string.Empty;
            if (val is DateTime dt) return dt.ToString("dd-MM-yyyy");
            if (val is double dbl)
            {
                try { return DateTime.FromOADate(dbl).ToString("dd-MM-yyyy"); }
                catch { }
            }
            string s = val.ToString().Trim();
            // Intentar parsear con multiples formatos para normalizar
            var fmts = new[] { "dd-MM-yyyy", "dd/MM/yyyy", "d-M-yyyy", "d/M/yyyy", "M/d/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "yyyy/MM/dd" };
            if (DateTime.TryParseExact(s, fmts, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime parsed))
                return parsed.ToString("dd-MM-yyyy");
            // Fallback: normalizar separadores si el patron basico coincide
            if (Regex.IsMatch(s, @"^\d{1,2}[-/]\d{1,2}[-/]\d{4}$"))
                return s.Replace('/', '-');
            return s;
        }

        // Parsea una fecha en multiples formatos y retorna DateTime; usa fallback si falla
        private static DateTime ParseRipsDate(string s, DateTime fallback)
        {
            if (string.IsNullOrWhiteSpace(s)) return fallback;
            s = s.Trim();
            var fmts = new[] { "dd-MM-yyyy", "dd/MM/yyyy", "d-M-yyyy", "d/M/yyyy", "M/d/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "yyyy/MM/dd" };
            if (DateTime.TryParseExact(s, fmts, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime d))
                return d;
            if (DateTime.TryParse(s, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime d2))
                return d2;
            return fallback;
        }

        // ─────────────────────────────────────────────────────────────────────
        // FLUJO RIPS
        // ─────────────────────────────────────────────────────────────────────

        // Boton "Continuar bot" — el usuario lo presiona DESPUES de seleccionar el convenio y esperar que cargue
        private void btnContinuarBot_Click(object sender, EventArgs e)
        {
            if (!_waitingForUserConvenio) return;
            _waitingForUserConvenio = false;
            btnContinuarBot.Enabled = false;
            panelConvenio.Visible = false;
            WriteUiLog("Boton Continuar presionado. Re-llenando CC/fechas y haciendo Buscar...");

            var tRefill = new System.Windows.Forms.Timer { Interval = 800 };
            tRefill.Tick += (tsR, teR) =>
            {
                ((System.Windows.Forms.Timer)tsR).Stop();
                try
                {
                    var recR = _ripsRecords[_ripsRecordIndex];
                    // Re-llenar fecha inicio
                    DateTime fiDateR = ParseRipsDate(recR.FechaInicio, new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1));
                    JsInRips("(function(){var dt=new Date(" + fiDateR.Year + "," + (fiDateR.Month-1) + "," + fiDateR.Day + ");var c=window['ctl00_ctl00_ContentPlaceHolder1_mainContent_CntFechas_FechaInicialDateEdit'];if(c&&c.SetDate)c.SetDate(dt);})()");
                    // Re-llenar fecha fin
                    DateTime ffDateR = ParseRipsDate(recR.FechaFin, DateTime.MinValue);
                    if (ffDateR == DateTime.MinValue || ffDateR < fiDateR)
                        ffDateR = new DateTime(fiDateR.Year, fiDateR.Month, DateTime.DaysInMonth(fiDateR.Year, fiDateR.Month));
                    JsInRips("(function(){var dt=new Date(" + ffDateR.Year + "," + (ffDateR.Month-1) + "," + ffDateR.Day + ");var c=window['ctl00_ctl00_ContentPlaceHolder1_mainContent_CntFechas_FechaFinalDateEdit'];if(c&&c.SetDate)c.SetDate(dt);})()");
                    // Re-llenar CC via JS (sin fireEvent para no disparar UpdatePanel)
                    string ccJs = EscapeJs(recR.CC ?? string.Empty);
                    string ccR = JsInRips("(function(){var el=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_PacientesBuscador_ctl01_I');if(!el)return 'NF';el.value='" + ccJs + "';return 'OK:'+el.value;})()");
                    WriteUiLog("Re-fill CC: " + ccR);
                }
                catch (Exception exR) { WriteUiLog("Re-fill err: " + exR.Message); }

                // Esperar 800ms y hacer click en Buscar (JS + OS click de respaldo)
                var tBuscar = new System.Windows.Forms.Timer { Interval = 800 };
                tBuscar.Tick += (tsB, teB) =>
                {
                    ((System.Windows.Forms.Timer)tsB).Stop();
                    string bR = JsInRips(
                        "(function(){" +
                        "  var ids=['ctl00_ctl00_ContentPlaceHolder1_mainContent_ConsultarButton'," +
                        "           'ctl00_ctl00_ContentPlaceHolder1_mainContent_BuscarRipsButton'," +
                        "           'ctl00_ctl00_ContentPlaceHolder1_mainContent_BuscarButton'," +
                        "           'ctl00_ctl00_ContentPlaceHolder1_mainContent_btnBuscar'," +
                        "           'ctl00_ctl00_ContentPlaceHolder1_mainContent_btnConsultar'];" +
                        "  for(var i=0;i<ids.length;i++){var el=document.getElementById(ids[i]);if(el){el.click();return 'JS_OK:'+ids[i];}}" +
                        "  var btns=document.querySelectorAll('input[type=button],input[type=submit],button');" +
                        "  for(var k=0;k<btns.length;k++){var v=(btns[k].value||btns[k].innerText||'').toLowerCase().trim();if(v==='buscar'||v==='consultar'||v.indexOf('buscar')>=0||v.indexOf('consultar')>=0){btns[k].click();return 'BTN_OK:'+btns[k].id+':'+v;}}" +
                        "  return 'BUSCAR_NF';" +
                        "})()");
                    WriteUiLog("Auto-Buscar JS: " + bR);
                    // Respaldo OS click si JS no encontro el boton
                    if (bR == "BUSCAR_NF")
                    {
                        string bPosOS = JsInRips(
                            "(function(){" +
                            "  var btns=document.querySelectorAll('input[type=button],input[type=submit],button');" +
                            "  for(var k=0;k<btns.length;k++){" +
                            "    var v=(btns[k].value||btns[k].innerText||'').toLowerCase().trim();" +
                            "    if(v.indexOf('buscar')>=0||v.indexOf('consultar')>=0){" +
                            "      var r=btns[k].getBoundingClientRect();" +
                            "      if(r.width>0&&r.height>0)return Math.round(r.left+(r.width/2))+','+Math.round(r.top+(r.height/2));" +
                            "    }" +
                            "  }" +
                            "  return 'NF';" +
                            "})()");
                        WriteUiLog("Auto-Buscar OS pos: " + bPosOS);
                        if (bPosOS.Contains(",") && !bPosOS.StartsWith("NF"))
                        {
                            try
                            {
                                string ifrB = GetRipsIframeOffset();
                                var ppB = bPosOS.Split(','); var ipB = ifrB.Split(',');
                                int sxB = webBrowser1.PointToScreen(new System.Drawing.Point(0,0)).X + int.Parse(ipB[0].Trim()) + int.Parse(ppB[0].Trim());
                                int syB = webBrowser1.PointToScreen(new System.Drawing.Point(0,0)).Y + int.Parse(ipB[1].Trim()) + int.Parse(ppB[1].Trim());
                                WriteUiLog("Auto-Buscar Click OS en: " + sxB + "," + syB);
                                this.Activate(); webBrowser1.Focus();
                                System.Threading.Thread.Sleep(200);
                                System.Windows.Forms.Cursor.Position = new System.Drawing.Point(sxB, syB);
                                System.Threading.Thread.Sleep(150);
                                NativeMouseClick(sxB, syB);
                            }
                            catch (Exception exBOS) { WriteUiLog("Auto-Buscar OS err: " + exBOS.Message); }
                        }
                    }
                    _awaitingBuscarResult = true;
                };
                tBuscar.Start();
            };
            tRefill.Start();
        }

        private void StartRipsFlow()
        {
            UpdateBotProgress();
            UpdateSidebarStatus("Procesando CC " + (_ripsRecordIndex + 1) + " de " + (_ripsRecords?.Count ?? 0) + "...");
            if (_ripsRecords == null || _ripsRecords.Count == 0)
            {
                // Si el usuario cargo un Excel y no tiene registros, es un error — no hacer fallback silencioso
                if (!string.IsNullOrWhiteSpace(_overrideExcelPath))
                {
                    WriteUiLog("StartRipsFlow: Excel cargado pero sin registros validos. Abortando.");
                    _ripsFlowDone = true;
                    UpdateSidebarStatus("Error: el archivo Excel no tiene registros validos.");
                    MessageBox.Show(
                        "El archivo Excel seleccionado no contiene registros validos.\n\n" +
                        "Verifique que el archivo tenga las columnas CC, FECHA INICIO, FECHA FIN y CONVENIO.\n\n" +
                        "Use el boton [x] Reiniciar Bot para cargar un archivo diferente.",
                        "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                WriteUiLog("StartRipsFlow: sin registros en BASE RIPS. Usando valores de config.");
                // Fallback a valores de bot-config.json solo si NO se selecciono un Excel
                _ripsRecords = new List<RipsRecord>
                {
                    new RipsRecord
                    {
                        CC = _cedulaPaciente ?? string.Empty,
                        FechaInicio = _ripsFechaInicio ?? "01-04-2026",
                        FechaFin = _ripsFechaFin ?? "30-04-2026"
                    }
                };
            }

            if (_ripsRecordIndex >= _ripsRecords.Count)
            {
                WriteUiLog("StartRipsFlow: todos los registros procesados (" + _ripsRecords.Count + ").");
                _ripsFlowDone = true;
                return;
            }

            var rec = _ripsRecords[_ripsRecordIndex];
            WriteUiLog(string.Format("StartRipsFlow: registro {0}/{1} CC={2} FI={3} FF={4}",
                _ripsRecordIndex + 1, _ripsRecords.Count, rec.CC, rec.FechaInicio, rec.FechaFin));

            _ripsGeneration++;  // invalida cualquier tRetry/tBuscar huerfano del ciclo anterior
            _ripsStep = 0;
            _ripsCcRetries = 0;
            _convenioRetries = 0;
            _buscarSinResultadoRetries = 0;
            _gridEmptyPolls = 0;
            _waitingForUserConvenio = false;
            _awaitingBuscarResult = false;
            _awaitingCompletarForm = false;
            _completarFormUrl = null;
            // Detener y destruir cualquier timer previo antes de crear uno nuevo
            if (_ripsTimer != null) { try { _ripsTimer.Stop(); _ripsTimer.Dispose(); } catch { } _ripsTimer = null; }
            _ripsTimer = new System.Windows.Forms.Timer { Interval = 1200 };
            _ripsTimer.Tick += RipsTimer_Tick;
            _ripsTimer.Start();
        }

        // Helper: ejecuta JS en el frame correcto de RIPS
        // Extrae el path base de una URL (sin querystring ni fragmento)
        private static string UrlBase(string url)
        {
            if (string.IsNullOrEmpty(url)) return string.Empty;
            int q = url.IndexOf('?'); if (q >= 0) url = url.Substring(0, q);
            int f = url.IndexOf('#'); if (f >= 0) url = url.Substring(0, f);
            return url;
        }

        private string JsInRips(string js)
        {
            try
            {
                HtmlDocument mainDoc = webBrowser1.Document;
                if (mainDoc == null) return "NO_DOC";

                // 1) Si tenemos URL base del formulario Completar, buscar ese frame primero
                // Comparar solo el path base (sin querystring) para tolerar cambios post-doPostBack
                if (!string.IsNullOrEmpty(_completarFormUrl))
                {
                    string baseComp = UrlBase(_completarFormUrl);
                    foreach (HtmlWindow frame in mainDoc.Window.Frames)
                    {
                        try
                        {
                            string furl = frame.Document?.Url?.ToString() ?? string.Empty;
                            if (string.Equals(UrlBase(furl), baseComp, StringComparison.OrdinalIgnoreCase))
                            {
                                object r = frame.Document.InvokeScript("eval", new object[] { js });
                                return r?.ToString() ?? "(null)";
                            }
                        }
                        catch { }
                    }
                }

                // 2) Buscar frame de ConsultarEstadoRips o FacturacionPaciente
                foreach (HtmlWindow frame in mainDoc.Window.Frames)
                {
                    try
                    {
                        string furl = frame.Document?.Url?.ToString() ?? string.Empty;
                        if (furl.IndexOf("ConsultarEstadoRips", StringComparison.OrdinalIgnoreCase) >= 0
                            || furl.IndexOf("FacturacionPaciente", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            object result = frame.Document.InvokeScript("eval", new object[] { js });
                            return result?.ToString() ?? "(null)";
                        }
                    }
                    catch { }
                }

                // 3) Fallback: probar TODOS los frames con URL .aspx (excluir about:blank)
                foreach (HtmlWindow frame in mainDoc.Window.Frames)
                {
                    try
                    {
                        string furl = frame.Document?.Url?.ToString() ?? string.Empty;
                        if (furl.IndexOf(".aspx", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            object result = frame.Document.InvokeScript("eval", new object[] { js });
                            string res = result?.ToString() ?? string.Empty;
                            if (!string.IsNullOrEmpty(res) && res != "(null)" && res != "NF")
                                return res;
                        }
                    }
                    catch { }
                }

                // 4) Último recurso: documento principal
                object r2 = mainDoc.InvokeScript("eval", new object[] { js });
                return r2?.ToString() ?? "(null)";
            }
            catch (Exception ex)
            {
                return "ERR:" + ex.Message;
            }
        }

        private void RipsTimer_Tick(object sender, EventArgs e)
        {
            var timer = (System.Windows.Forms.Timer)sender;
            try
            {
                string r;
                switch (_ripsStep)
                {
                    case 0: // Checkbox PRIMERO — OS click dispara UpdatePanel, las fechas se ponen despues
                    {
                        string cbPos0 = JsInRips(
                            "(function(){" +
                            "  var el=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_PacienteCheckBox_I');" +
                            "  if(!el)return 'NF';" +
                            "  var r=el.getBoundingClientRect();" +
                            "  return Math.round(r.left+(r.width/2||6))+','+Math.round(r.top+(r.height/2||6));" +
                            "})()");
                        WriteUiLog("RIPS step0 Checkbox pos: " + cbPos0);
                        if (cbPos0 != "NF" && cbPos0.Contains(","))
                        {
                            try
                            {
                                string ifrOff0 = GetRipsIframeOffset();
                                var pp0 = cbPos0.Split(',');
                                var ip0 = ifrOff0.Split(',');
                                int elX0 = int.Parse(pp0[0].Trim()), elY0 = int.Parse(pp0[1].Trim());
                                int ifrX0 = int.Parse(ip0[0].Trim()), ifrY0 = int.Parse(ip0[1].Trim());
                                var bPt0 = webBrowser1.PointToScreen(new System.Drawing.Point(0, 0));
                                int screenX0 = bPt0.X + ifrX0 + elX0;
                                int screenY0 = bPt0.Y + ifrY0 + elY0;
                                WriteUiLog("RIPS step0 Click OS checkbox en: " + screenX0 + "," + screenY0);
                                this.Activate(); webBrowser1.Focus();
                                System.Threading.Thread.Sleep(200);
                                System.Windows.Forms.Cursor.Position = new System.Drawing.Point(screenX0, screenY0);
                                System.Threading.Thread.Sleep(150);
                                NativeMouseClick(screenX0, screenY0);
                            }
                            catch (Exception ex0c) { WriteUiLog("RIPS step0 checkbox err: " + ex0c.Message); }
                        }
                        _ripsStep++;
                        timer.Interval = 3500; // esperar UpdatePanel
                        break;
                    }

                    case 1: // Llenar campo TipoDocumento (CC) y campo CC — OS mouse click + JS write
                    {
                        // 1a. Seleccionar TipoDocumento = CC antes de llenar la cédula
                        string tipDocR = JsInRips(
                            "(function(){" +
                            "  try{" +
                            "    var cb=window['ctl00_ctl00_ContentPlaceHolder1_mainContent_TipoDocumentoComboBox'];" +
                            "    if(!cb){var coll=ASPxClientControl.GetControlCollection();cb=coll.GetByName('TipoDocumentoComboBox');}" +
                            "    if(cb&&cb.GetItemCount){" +
                            "      for(var i=0;i<cb.GetItemCount();i++){" +
                            "        var it=cb.GetItem(i);" +
                            "        if(it&&(it.text||'').replace(/\\s/g,'').toUpperCase().indexOf('CC')>=0){" +
                            "          cb.SetSelectedIndex(i);return 'TIPODOC_CC:'+it.text;" +
                            "        }" +
                            "      }" +
                            "      cb.SetSelectedIndex(1);return 'TIPODOC_IDX1';" +
                            "    }" +
                            "  }catch(e){}" +
                            "  var vi=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_TipoDocumentoComboBox_VI');" +
                            "  var ti=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_TipoDocumentoComboBox_I');" +
                            "  if(vi)vi.value='1'; if(ti)ti.value='CC';" +
                            "  return 'TIPODOC_DIRECT';" +
                            "})()");
                        WriteUiLog("RIPS step1 TipoDocumento: " + tipDocR);

                        string cc = _ripsRecords[_ripsRecordIndex].CC ?? string.Empty;
                        string inpPos3 = JsInRips(
                            "(function(){" +
                            "  var el=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_PacientesBuscador_ctl01_I');" +
                            "  if(!el)return 'NF';" +
                            "  if(el.offsetParent===null||el.offsetWidth===0)return 'HIDDEN';" +
                            "  var r=el.getBoundingClientRect();" +
                            "  return Math.round(r.left+(r.width/2||50))+','+Math.round(r.top+(r.height/2||8));" +
                            "})()");
                        WriteUiLog("RIPS step1 CC field pos: " + inpPos3);
                        bool ccReady3 = inpPos3.Contains(",") && !inpPos3.StartsWith("NF") && !inpPos3.StartsWith("HIDDEN");
                        if (!ccReady3 && _ripsCcRetries < 5)
                        {
                            _ripsCcRetries++;
                            WriteUiLog("RIPS step1 campo no listo, reintento " + _ripsCcRetries);
                            timer.Interval = 1500;
                            break;
                        }
                        _ripsCcRetries = 0;
                        if (ccReady3)
                        {
                            try
                            {
                                string ifrOff3 = GetRipsIframeOffset();
                                var pp3 = inpPos3.Split(',');
                                var ip3 = ifrOff3.Split(',');
                                int elX3 = int.Parse(pp3[0].Trim()), elY3 = int.Parse(pp3[1].Trim());
                                int ifrX3 = int.Parse(ip3[0].Trim()), ifrY3 = int.Parse(ip3[1].Trim());
                                var bPt3 = webBrowser1.PointToScreen(new System.Drawing.Point(0, 0));
                                int screenX3 = bPt3.X + ifrX3 + elX3;
                                int screenY3 = bPt3.Y + ifrY3 + elY3;
                                WriteUiLog("RIPS step1 Click OS campo CC en: " + screenX3 + "," + screenY3);
                                this.Activate(); webBrowser1.Focus();
                                System.Threading.Thread.Sleep(200);
                                System.Windows.Forms.Cursor.Position = new System.Drawing.Point(screenX3, screenY3);
                                System.Threading.Thread.Sleep(150);
                                NativeMouseClick(screenX3, screenY3);
                                System.Threading.Thread.Sleep(400);
                                string ec3 = EscapeJs(cc);
                                string jsR3 = JsInRips(
                                    "(function(){" +
                                    "  var el=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_PacientesBuscador_ctl01_I');" +
                                    "  if(!el)return 'CC_NF';" +
                                    "  el.value='" + ec3 + "';" +
                                    "  try{el.focus();}catch(e){}" +
                                    "  return 'CC_JS_OK:'+el.value;" +
                                    "})()");
                                WriteUiLog("RIPS step1 CC via JS: " + jsR3);
                            }
                            catch (Exception ex3) { WriteUiLog("RIPS step1 err: " + ex3.Message); }
                        }
                        _ripsStep++;
                        timer.Interval = 1500;
                        break;
                    }

                    case 2: // Llenar FECHAS (fecha inicio y fecha fin) via OS click + SetDate — ANTES de la lupa
                    {
                        try
                        {
                            var rec2 = _ripsRecords[_ripsRecordIndex];
                            // Fecha inicio: OS click para activar, luego SetDate via DevExpress
                            string fPos1 = JsInRips("(function(){var el=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_CntFechas_FechaInicialDateEdit_I');if(!el)return 'NF';var r=el.getBoundingClientRect();if(r.width===0)return 'NF';return Math.round(r.left+(r.width/2||50))+','+Math.round(r.top+(r.height/2||8));})()");
                            WriteUiLog("RIPS step2 FechaInicio pos: " + fPos1);
                            if (fPos1.Contains(",") && !fPos1.StartsWith("NF"))
                            {
                                string ifrF1 = GetRipsIframeOffset(); var ppF1 = fPos1.Split(','); var ipF1 = ifrF1.Split(',');
                                int sxF1 = webBrowser1.PointToScreen(new System.Drawing.Point(0,0)).X + int.Parse(ipF1[0].Trim()) + int.Parse(ppF1[0].Trim());
                                int syF1 = webBrowser1.PointToScreen(new System.Drawing.Point(0,0)).Y + int.Parse(ipF1[1].Trim()) + int.Parse(ppF1[1].Trim());
                                this.Activate(); webBrowser1.Focus(); System.Threading.Thread.Sleep(150);
                                System.Windows.Forms.Cursor.Position = new System.Drawing.Point(sxF1, syF1);
                                System.Threading.Thread.Sleep(150); NativeMouseClick(sxF1, syF1); System.Threading.Thread.Sleep(250);
                            }
                            DateTime fiDate2 = ParseRipsDate(rec2.FechaInicio, new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1));
                            string fiR2 = JsInRips("(function(){var dt=new Date(" + fiDate2.Year + "," + (fiDate2.Month-1) + "," + fiDate2.Day + ");var c=window['ctl00_ctl00_ContentPlaceHolder1_mainContent_CntFechas_FechaInicialDateEdit'];if(c&&c.SetDate){c.SetDate(dt);return 'DX_OK:" + fiDate2.ToString("dd-MM-yyyy") + "';}var inp=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_CntFechas_FechaInicialDateEdit_I');if(inp){inp.value='" + fiDate2.ToString("dd/MM/yyyy") + "';return 'VAL_OK';}return 'NF';})()");
                            WriteUiLog("RIPS step2 FechaInicio set: " + fiR2);
                            // Fecha fin: OS click para activar, luego SetDate via DevExpress
                            string fPos2 = JsInRips("(function(){var el=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_CntFechas_FechaFinalDateEdit_I');if(!el)return 'NF';var r=el.getBoundingClientRect();if(r.width===0)return 'NF';return Math.round(r.left+(r.width/2||50))+','+Math.round(r.top+(r.height/2||8));})()");
                            WriteUiLog("RIPS step2 FechaFin pos: " + fPos2);
                            if (fPos2.Contains(",") && !fPos2.StartsWith("NF"))
                            {
                                string ifrF2 = GetRipsIframeOffset(); var ppF2 = fPos2.Split(','); var ipF2 = ifrF2.Split(',');
                                int sxF2 = webBrowser1.PointToScreen(new System.Drawing.Point(0,0)).X + int.Parse(ipF2[0].Trim()) + int.Parse(ppF2[0].Trim());
                                int syF2 = webBrowser1.PointToScreen(new System.Drawing.Point(0,0)).Y + int.Parse(ipF2[1].Trim()) + int.Parse(ppF2[1].Trim());
                                System.Windows.Forms.Cursor.Position = new System.Drawing.Point(sxF2, syF2);
                                System.Threading.Thread.Sleep(150); NativeMouseClick(sxF2, syF2); System.Threading.Thread.Sleep(250);
                            }
                            DateTime ffDate2 = ParseRipsDate(rec2.FechaFin, DateTime.MinValue);
                            if (ffDate2 == DateTime.MinValue || ffDate2 < fiDate2)
                                ffDate2 = new DateTime(fiDate2.Year, fiDate2.Month, DateTime.DaysInMonth(fiDate2.Year, fiDate2.Month));
                            string ffR2 = JsInRips("(function(){var dt=new Date(" + ffDate2.Year + "," + (ffDate2.Month-1) + "," + ffDate2.Day + ");var c=window['ctl00_ctl00_ContentPlaceHolder1_mainContent_CntFechas_FechaFinalDateEdit'];if(c&&c.SetDate){c.SetDate(dt);return 'DX_OK:" + ffDate2.ToString("dd-MM-yyyy") + "';}var inp=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_CntFechas_FechaFinalDateEdit_I');if(inp){inp.value='" + ffDate2.ToString("dd/MM/yyyy") + "';return 'VAL_OK';}return 'NF';})()");
                            WriteUiLog("RIPS step2 FechaFin set: " + ffR2);
                        }
                        catch (Exception ex2f) { WriteUiLog("RIPS step2 fechas err: " + ex2f.Message); }
                        _ripsStep++;
                        timer.Interval = 1200; // solo espera corta — NO se dispara UpdatePanel porque no usamos fireEvent
                        break;
                    }

                    case 3: // Click LUPA — DevExpress ASPxClientButton.DoClick()
                    {
                        string lupaClick = JsInRips(
                            "(function(){" +
                            "  var btnId='ctl00_ctl00_ContentPlaceHolder1_mainContent_PacientesBuscador_PacientesBuscador_buscarBoton';" +
                            "  var dxBtn=window[btnId];" +
                            "  if(dxBtn&&dxBtn.DoClick){dxBtn.DoClick();return 'DX_DOCLICK_OK';}" +
                            // Fallback: .click() directo sobre el input submit (aunque tenga 0px de tamaño)
                            "  var sub=document.querySelector('input[name=\"ctl00$ctl00$ContentPlaceHolder1$mainContent$PacientesBuscador$PacientesBuscador_buscarBoton\"]');" +
                            "  if(sub){sub.click();return 'SUBMIT_CLICK_OK';}" +
                            "  return 'NF';" +
                            "})()");
                        WriteUiLog("RIPS step3 lupa: " + lupaClick);
                        _ripsStep++;
                        timer.Interval = 4500;
                        break;
                    }

                    case 4: // SELECCIONAR fila del ConveniosDataGrid (Infragistics) que abre la lupa
                    {
                        string convenio4 = (_ripsRecords[_ripsRecordIndex].TipoConvenio ?? "").Trim();
                        WriteUiLog("RIPS step4 seleccionando convenio en ConveniosDataGrid: [" + convenio4 + "]");
                        string eConv4 = EscapeJs(convenio4.ToUpperInvariant());
                        // Diagnostico: estado del ConveniosDataGrid
                        string diag4 = JsInRips(
                            "(function(){" +
                            "  var hf=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_ConveniosDataGrid');" +
                            "  var hfVal=hf?hf.value.substring(0,80):'NO_HF';" +
                            "  var rows=document.querySelectorAll('[id*=ConveniosDataGrid_r_]');" +
                            "  var rowsTxt='';for(var i=0;i<Math.min(rows.length,5);i++)rowsTxt+='['+i+':'+( rows[i].innerText||rows[i].textContent||'').trim().substring(0,40)+']';" +
                            "  return 'HF='+hfVal+'||ROWS='+rows.length+'||'+rowsTxt;" +
                            "})()");
                        WriteUiLog("RIPS step4 DIAG: " + diag4);
                        bool convenioOk = false;
                        // Estrategia 1: Infragistics igtbl API — obtener posicion de la fila y OS click
                        // (row.Element.click() en JS no dispara attachEvent de Infragistics en IE; usar OS click)
                        string igR4 = JsInRips(
                            "(function(){" +
                            "  var conv='" + eConv4 + "';" +
                            "  try{" +
                            "    var g=igtbl_getGridById('ctl00_ctl00_ContentPlaceHolder1_mainContent_ConveniosDataGrid');" +
                            "    if(g&&g.Rows&&g.Rows.length>0){" +
                            "      for(var i=0;i<g.Rows.length;i++){" +
                            "        var row=g.Rows.getRow(i);if(!row)continue;" +
                            "        var rt='';try{rt=(row.Element?(row.Element.innerText||row.Element.textContent||'').toUpperCase():'');}catch(e){}" +
                            "        if(conv===''||rt.indexOf(conv)>=0){" +
                            "          try{var r=row.Element.getBoundingClientRect();if(r.width>0&&r.height>0)return 'IG_POS:'+i+','+Math.round(r.left+r.width/2)+','+Math.round(r.top+r.height/2);}catch(ec){}" +
                            "          try{row.Element.click();return 'IG_CLICK:'+i;}catch(ec2){}" +
                            "        }" +
                            "      }" +
                            "      var r0=g.Rows.getRow(0);" +
                            "      if(r0){try{var rr=r0.Element.getBoundingClientRect();if(rr.width>0)return 'IG_POS:0,'+Math.round(rr.left+rr.width/2)+','+Math.round(rr.top+rr.height/2);}catch(ef){}try{r0.Element.click();return 'IG_FIRST:cnt='+g.Rows.length;}catch(ef2){}}" +
                            "    }" +
                            "  }catch(e){}" +
                            "  return 'IG_NF';" +
                            "})()");
                        WriteUiLog("RIPS step4 IG: " + igR4);
                        // IG_POS:i,x,y => OS click en la fila (dispara todos los handlers de Infragistics en IE)
                        if (igR4.StartsWith("IG_POS:"))
                        {
                            try
                            {
                                var igPosParts = igR4.Substring(7).Split(','); // "i,x,y"
                                if (igPosParts.Length >= 3)
                                {
                                    string ifrOffIG = GetRipsIframeOffset();
                                    var ipIG = ifrOffIG.Split(',');
                                    int sxIG = webBrowser1.PointToScreen(new System.Drawing.Point(0,0)).X + int.Parse(ipIG[0].Trim()) + int.Parse(igPosParts[1].Trim());
                                    int syIG = webBrowser1.PointToScreen(new System.Drawing.Point(0,0)).Y + int.Parse(ipIG[1].Trim()) + int.Parse(igPosParts[2].Trim());
                                    WriteUiLog("RIPS step4 IG OS click: " + sxIG + "," + syIG);
                                    this.Activate(); webBrowser1.Focus();
                                    System.Threading.Thread.Sleep(200);
                                    System.Windows.Forms.Cursor.Position = new System.Drawing.Point(sxIG, syIG);
                                    System.Threading.Thread.Sleep(150);
                                    NativeMouseClick(sxIG, syIG);
                                    convenioOk = true;
                                }
                            }
                            catch (Exception exIG) { WriteUiLog("RIPS step4 IG OS err: " + exIG.Message); }
                        }
                        else if (!igR4.StartsWith("IG_NF")) convenioOk = true;
                        // Estrategia 2: DOM — filas por id ConveniosDataGrid_r_N
                        if (!convenioOk)
                        {
                            string domR4 = JsInRips(
                                "(function(){" +
                                "  var conv='" + eConv4 + "';" +
                                "  var rows=document.querySelectorAll('[id*=ConveniosDataGrid_r_]');" +
                                "  for(var j=0;j<rows.length;j++){" +
                                "    var rt=(rows[j].innerText||rows[j].textContent||'').toUpperCase().trim();" +
                                "    if(conv===''||rt.indexOf(conv)>=0){rows[j].click();return 'DOM_MATCH:'+j+':'+rt.substring(0,30);}" +
                                "  }" +
                                "  if(rows.length>0){rows[0].click();return 'DOM_FIRST:'+rows.length;}" +
                                "  return 'DOM_NF';" +
                                "})()");
                            WriteUiLog("RIPS step4 DOM: " + domR4);
                            if (!domR4.StartsWith("DOM_NF")) convenioOk = true;
                        }
                        // Estrategia 3: OS click sobre elemento visible que contiene el texto del convenio
                        if (!convenioOk)
                        {
                            string convPos4 = JsInRips(
                                "(function(){" +
                                "  var conv='" + eConv4 + "';" +
                                "  var elems=document.querySelectorAll('td,li,span[onclick],div[onclick]');" +
                                "  for(var i=0;i<elems.length;i++){" +
                                "    var t=(elems[i].innerText||elems[i].textContent||'').toUpperCase().trim();" +
                                "    if(conv!==''&&t.indexOf(conv)>=0&&t.length<120){" +
                                "      var r=elems[i].getBoundingClientRect();" +
                                "      if(r.width>0&&r.height>0)return Math.round(r.left+(r.width/2))+','+Math.round(r.top+(r.height/2));" +
                                "    }" +
                                "  }" +
                                "  return 'NF';" +
                                "})()");
                            WriteUiLog("RIPS step4 OS pos: " + convPos4);
                            if (convPos4.Contains(",") && !convPos4.StartsWith("NF"))
                            {
                                try
                                {
                                    string ifrOffC = GetRipsIframeOffset();
                                    var ppC = convPos4.Split(','); var ipC = ifrOffC.Split(',');
                                    int sxC = webBrowser1.PointToScreen(new System.Drawing.Point(0,0)).X + int.Parse(ipC[0].Trim()) + int.Parse(ppC[0].Trim());
                                    int syC = webBrowser1.PointToScreen(new System.Drawing.Point(0,0)).Y + int.Parse(ipC[1].Trim()) + int.Parse(ppC[1].Trim());
                                    WriteUiLog("RIPS step4 Click OS en: " + sxC + "," + syC);
                                    this.Activate(); webBrowser1.Focus();
                                    System.Threading.Thread.Sleep(200);
                                    System.Windows.Forms.Cursor.Position = new System.Drawing.Point(sxC, syC);
                                    System.Threading.Thread.Sleep(150);
                                    NativeMouseClick(sxC, syC);
                                    convenioOk = true;
                                }
                                catch (Exception exC4) { WriteUiLog("RIPS step4 OS err: " + exC4.Message); }
                            }
                        }
                        if (!convenioOk)
                        {
                            WriteUiLog("RIPS step4 ConveniosDataGrid NO encontrado — posible que la lupa no trajo resultados");
                            // Reintentar step4 hasta 4 veces esperando que cargue el AJAX
                            if (_convenioRetries < 4)
                            {
                                _convenioRetries++;
                                WriteUiLog("RIPS step4 reintento " + _convenioRetries + "/4 en 3s...");
                                timer.Interval = 3000;
                                break; // no avanzar _ripsStep, volver a step4
                            }
                            WriteUiLog("RIPS step4 sin filas tras 4 reintentos — continuando igual");
                        }
                        // Serializar Infragistics para asegurar que el convenio quede en el POST de Buscar
                        if (convenioOk) JsInRips("(function(){try{igtbl_preSubmit();}catch(e){}try{igtbl_serialize('ctl00_ctl00_ContentPlaceHolder1_mainContent_ConveniosDataGrid');}catch(e){}})()");
                        if (convenioOk) WriteUiLog("RIPS step4 igtbl_preSubmit ejecutado");
                        _convenioRetries = 0;
                        _ripsStep++;
                        timer.Interval = 2000;  // esperar postback de seleccion de convenio
                        break;
                    }

                    case 5: // Re-establecer FECHAS — atómico: FechaFin primero para evitar NaN por validación interna
                    {
                        try
                        {
                            var rec5 = _ripsRecords[_ripsRecordIndex];
                            DateTime fiDate5 = ParseRipsDate(rec5.FechaInicio, new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1));
                            DateTime ffDate5 = ParseRipsDate(rec5.FechaFin, DateTime.MinValue);
                            if (ffDate5 == DateTime.MinValue || ffDate5 < fiDate5)
                                ffDate5 = new DateTime(fiDate5.Year, fiDate5.Month,
                                    DateTime.DaysInMonth(fiDate5.Year, fiDate5.Month));
                            // Una sola llamada JS: FechaFin primero, luego FechaInicio
                            // Así cuando FechaInicio dispare el evento interno, FechaFin ya tiene valor válido
                            string r5 = JsInRips(
                                "(function(){" +
                                "  var dtFI=new Date(" + fiDate5.Year + "," + (fiDate5.Month - 1) + "," + fiDate5.Day + ");" +
                                "  var dtFF=new Date(" + ffDate5.Year + "," + (ffDate5.Month - 1) + "," + ffDate5.Day + ");" +
                                "  var cFI=window['ctl00_ctl00_ContentPlaceHolder1_mainContent_CntFechas_FechaInicialDateEdit'];" +
                                "  var cFF=window['ctl00_ctl00_ContentPlaceHolder1_mainContent_CntFechas_FechaFinalDateEdit'];" +
                                "  var iFI=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_CntFechas_FechaInicialDateEdit_I');" +
                                "  var iFF=document.getElementById('ctl00_ctl00_ContentPlaceHolder1_mainContent_CntFechas_FechaFinalDateEdit_I');" +
                                "  if(cFF&&cFF.SetDate)cFF.SetDate(dtFF); else if(iFF)iFF.value='" + ffDate5.ToString("dd/MM/yyyy") + "';" +
                                "  if(cFI&&cFI.SetDate)cFI.SetDate(dtFI); else if(iFI)iFI.value='" + fiDate5.ToString("dd/MM/yyyy") + "';" +
                                "  return 'OK:FI=" + fiDate5.ToString("dd-MM-yyyy") + " FF=" + ffDate5.ToString("dd-MM-yyyy") + "';" +
                                "})()");
                            WriteUiLog("RIPS step5 fechas: " + r5);
                        }
                        catch (Exception ex5f) { WriteUiLog("RIPS step5 fechas err: " + ex5f.Message); }
                        _ripsStep++;
                        timer.Interval = 3000; // tiempo suficiente para que el postback del convenio complete
                        break;
                    }

                    case 6: // Click en BUSCAR — DevExpress DoClick o JS directo
                    {
                        // Serializar Infragistics antes de Buscar para que el convenio quede en el POST
                        JsInRips("(function(){try{igtbl_preSubmit();}catch(e){}})()");
                        string buscarR = JsInRips(
                            "(function(){" +
                            // Buscar todos los ASPxClientButton registrados en window y buscar el de Buscar
                            "  var keys=Object.keys(window);" +
                            "  for(var i=0;i<keys.length;i++){" +
                            "    try{" +
                            "      var obj=window[keys[i]];" +
                            "      if(obj&&typeof obj.DoClick==='function'&&keys[i].toLowerCase().indexOf('buscar')>=0&&keys[i].toLowerCase().indexOf('buscador')<0){" +
                            "        obj.DoClick();" +
                            "        return 'DX_DOCLICK:'+keys[i];" +
                            "      }" +
                            "    }catch(e){}" +
                            "  }" +
                            // Fallback: input[type=submit] o button con texto Buscar
                            "  var btns=document.querySelectorAll('input[type=submit],input[type=button],button,a');" +
                            "  for(var j=0;j<btns.length;j++){" +
                            "    var v=(btns[j].value||btns[j].innerText||btns[j].title||'').toLowerCase().trim();" +
                            "    if(v==='buscar'||v.indexOf('buscar')>=0){" +
                            "      btns[j].click();" +
                            "      return 'BTN_CLICK:'+btns[j].id;" +
                            "    }" +
                            "  }" +
                            "  return 'NF';" +
                            "})()");
                        WriteUiLog("RIPS step6 Buscar: " + buscarR);
                        WriteUiLog("RIPS step6 Buscar enviado — esperando carga de pagina...");
                        _awaitingBuscarResult = true;
                        timer.Stop();
                        break;
                    }
                    default:
                        timer.Stop();
                        break;
                }
            }
            catch (Exception ex)
            {
                WriteUiLog("RIPS step " + _ripsStep + " error: " + ex.Message);
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // P/Invoke Win32 mouse click
        // ─────────────────────────────────────────────────────────────────────

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, int dx, int dy, uint cButtons, int dwExtraInfo);
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP   = 0x0004;
        private static void NativeMouseClick(int screenX, int screenY)
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            System.Threading.Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        }

        // ─────────────────────────────────────────────────────────────────────
        // Procesamiento del grid de resultados RIPS
        // ─────────────────────────────────────────────────────────────────────

        private void StartRipsGridProcessing()
        {
            // Volcar HTML del grid a archivo para diagnóstico
            string gridHtml = JsInRips(
                "(function(){" +
                "  var g=document.querySelector('[id*=DXMainTable],[id*=GridView],[id*=gridView],[id*=Grid]');" +
                "  if(g) return g.outerHTML.substring(0,4000);" +
                "  return document.body?document.body.innerHTML.substring(0,4000):'NF';" +
                "})()");
            WriteUiLog("RIPS grid html len=" + (gridHtml ?? "").Length);
            ProcessNextIncompleto();
        }

        // JS de detección del grid: busca filas con ≥5 celdas directas que contengan
        // 'completo' o 'incompleto'. Usa rows[i].cells (solo hijos directos <td>/<th>)
        // para no contar celdas de tablas anidadas dentro de una misma fila.
        private const string JS_GRID_DETECT =
            "(function(){" +
            "  var rows=document.querySelectorAll('tr');" +
            "  var gridRowCount=0;" +
            "  for(var i=0;i<rows.length;i++){" +
            "    var cells=rows[i].cells;" +
            "    if(!cells||cells.length<5)continue;" +
            "    var rt=(rows[i].innerText||rows[i].textContent||'').replace(/\\u00a0/g,' ').toLowerCase();" +
            "    if(rt.indexOf('completo')<0&&rt.indexOf('incompleto')<0)continue;" +
            "    gridRowCount++;" +
            "    if(rt.indexOf('incompleto')>=0){" +
            "      var els=rows[i].querySelectorAll('a,input[type=button],input[type=image],input[type=submit],button,td[onclick],div[onclick],span[onclick]');" +
            "      for(var j=0;j<els.length;j++){" +
            "        var v=((els[j].value||els[j].innerText||els[j].title||els[j].alt||'').replace(/\\u00a0/g,' ')).toLowerCase().trim();" +
            "        if(v.indexOf('complet')>=0&&v.indexOf('incomplet')<0){" +
            "          var oc=els[j].getAttribute('onclick')||'';" +
            "          if(oc){try{eval(oc);}catch(e){try{els[j].click();}catch(e2){}}}" +
            "          else{try{els[j].click();}catch(e){}}" +
            "          return 'CLICKED:'+els[j].tagName+'['+els[j].id+']';" +
            "        }" +
            "      }" +
            "      return 'INCOMPLETO_SIN_BTN:'+rt.replace(/\\s+/g,' ').substring(0,100);" +
            "    }" +
            "  }" +
            "  if(gridRowCount>0)return 'TODOS_COMPLETOS';" +
            "  return 'GRID_VACIO:rows='+rows.length;" +
            "})()";

        // Ejecuta JS_GRID_DETECT en un HtmlDocument y devuelve el resultado
        private string RunGridDetectIn(HtmlDocument doc)
        {
            if (doc == null) return "GRID_VACIO";
            try
            {
                object r = doc.InvokeScript("eval", new object[] { JS_GRID_DETECT });
                return r?.ToString() ?? "GRID_VACIO";
            }
            catch { return "GRID_VACIO"; }
        }

        // Escanea recursivamente (hasta depth niveles) todos los frames del documento
        // buscando el que contiene el grid. Devuelve el resultado y la URL del frame.
        private string ScanGridAllFrames(HtmlDocument doc, int depth, out string foundUrl)
        {
            foundUrl = string.Empty;
            if (doc == null || depth < 0) return "GRID_VACIO";

            // Intentar en este documento primero
            string res = RunGridDetectIn(doc);
            string url = doc.Url?.ToString() ?? string.Empty;
            WriteUiLog("GridScan[" + (url.Length > 60 ? url.Substring(url.Length - 60) : url) + "] → " + (res.Length > 60 ? res.Substring(0, 60) : res));
            if (!res.StartsWith("GRID_VACIO") && !res.StartsWith("ERR"))
            {
                foundUrl = url;
                return res;
            }

            // Bajar a frames hijos
            try
            {
                foreach (HtmlWindow frame in doc.Window.Frames)
                {
                    try
                    {
                        HtmlDocument fd = frame.Document;
                        if (fd == null) continue;
                        string childRes = ScanGridAllFrames(fd, depth - 1, out foundUrl);
                        if (!childRes.StartsWith("GRID_VACIO") && !childRes.StartsWith("ERR"))
                            return childRes;
                    }
                    catch { }
                }
            }
            catch { }

            return "GRID_VACIO";
        }

        private void ProcessNextIncompleto()
        {
            // Escanear TODOS los frames recursivamente desde C# (IE bloquea esto desde JS
            // pero no desde el host managed). Esto garantiza que encontramos el frame correcto
            // sin importar la profundidad del iframe anidado.
            string foundInFrame;
            string result = "GRID_VACIO";
            try
            {
                HtmlDocument mainDoc = webBrowser1.Document;
                result = ScanGridAllFrames(mainDoc, 4, out foundInFrame);
                WriteUiLog("RIPS grid incompleto: " + result + (foundInFrame.Length > 0 ? " [en: ..." + (foundInFrame.Length > 50 ? foundInFrame.Substring(foundInFrame.Length - 50) : foundInFrame) + "]" : " [no frame encontrado]"));
            }
            catch (Exception exScan)
            {
                WriteUiLog("ProcessNextIncompleto ERR: " + exScan.Message);
            }

            if (result != null && result.StartsWith("CLICKED"))
            {
                // Esperar que cargue el formulario completar
                _buscarSinResultadoRetries = 0;
                _awaitingCompletarForm = true;
                WriteUiLog("RIPS grid: esperando formulario Completar...");
            }
            else if (result != null && result.StartsWith("TODOS_COMPLETOS"))
            {
                // Todo ya esta completo — avanzar al siguiente registro sin reintentos
                WriteUiLog("RIPS grid: todo completo. Avanzando al siguiente registro.");
                _buscarSinResultadoRetries = 0;
                AdvanceToNextRipsRecord();
            }
            else
            {
                // Primero re-chequear el grid varias veces: puede estar cargando via AJAX
                const int MaxPolls = 4;
                if (_gridEmptyPolls < MaxPolls)
                {
                    _gridEmptyPolls++;
                    WriteUiLog(string.Format(
                        "RIPS grid: grid vacio, recheck AJAX {0}/{1}. CC={2}",
                        _gridEmptyPolls, MaxPolls,
                        _ripsRecords[_ripsRecordIndex].CC));
                    int myGenPoll = _ripsGeneration;
                    var tPoll = new System.Windows.Forms.Timer { Interval = 3000 };
                    tPoll.Tick += (tsPo, tePo) =>
                    {
                        ((System.Windows.Forms.Timer)tsPo).Stop();
                        if (_ripsGeneration != myGenPoll) return;
                        ProcessNextIncompleto();
                    };
                    tPoll.Start();
                    return;
                }
                // Grid sigue vacio despues de todos los polls → probable error servidor
                _gridEmptyPolls = 0;
                const int MaxBuscarRetries = 5;
                _buscarSinResultadoRetries++;
                WriteUiLog(string.Format(
                    "RIPS grid: grid vacio/error (intento {0}/{1}). CC={2}",
                    _buscarSinResultadoRetries, MaxBuscarRetries,
                    _ripsRecords[_ripsRecordIndex].CC));

                if (_buscarSinResultadoRetries < MaxBuscarRetries)
                {
                    // Reintentar: primero cerrar el dialogo de error (si sigue abierto), luego re-Buscar
                    WriteUiLog("RIPS grid: cerrando dialogo de error y re-clickando Buscar...");
                    int myGen = _ripsGeneration;  // capturar generacion actual para detectar ciclos obsoletos
                    var tRetry = new System.Windows.Forms.Timer { Interval = 2000 };
                    tRetry.Tick += (tsRe, teRe) =>
                    {
                        ((System.Windows.Forms.Timer)tsRe).Stop();
                        // Si ya avanzamos a otro ciclo (TODOS_COMPLETOS u otro CC), ignorar
                        if (_ripsGeneration != myGen) return;
                        _awaitingBuscarResult = false;
                        _awaitingCompletarForm = false;
                        _completarFormUrl = null;

                        // Paso 1: buscar y cerrar el dialogo "Aceptar" en doc principal y todos los frames
                        const string JS_ACEPTAR =
                            "(function(){" +
                            "  var all=document.querySelectorAll('input,button,td,a,span,div');" +
                            "  for(var i=0;i<all.length;i++){" +
                            "    var v=(all[i].value||all[i].innerText||'').replace(/\\s/g,'').toLowerCase();" +
                            "    if(v==='aceptar'){try{all[i].click();}catch(ex){} return 'ACEPTAR_OK:'+all[i].tagName;}" +
                            "  }" +
                            "  return 'NO_DIALOG';" +
                            "})()";
                        string dismissed = "NO_DIALOG";
                        HtmlDocument mainDoc = webBrowser1.Document;
                        if (mainDoc != null)
                        {
                            try
                            {
                                object r0 = mainDoc.InvokeScript("eval", new object[] { JS_ACEPTAR });
                                dismissed = r0?.ToString() ?? "NO_DIALOG";
                                WriteUiLog("Dismiss mainDoc: " + dismissed);
                            }
                            catch { }
                            if (!dismissed.StartsWith("ACEPTAR_OK"))
                            {
                                try
                                {
                                    foreach (HtmlWindow frame in mainDoc.Window.Frames)
                                    {
                                        try
                                        {
                                            HtmlDocument fd = frame.Document;
                                            if (fd == null) continue;
                                            object r1 = fd.InvokeScript("eval", new object[] { JS_ACEPTAR });
                                            string rs1 = r1?.ToString() ?? string.Empty;
                                            WriteUiLog("Dismiss frame[" + (fd.Url?.ToString() ?? "") + "]: " + rs1);
                                            if (rs1.StartsWith("ACEPTAR_OK")) { dismissed = rs1; break; }
                                        }
                                        catch { }
                                    }
                                }
                                catch { }
                            }
                        }

                        // Paso 2: esperar que el dialogo se cierre, luego hacer Buscar
                        int delay = dismissed.StartsWith("ACEPTAR_OK") ? 2500 : 1000;
                        var tBuscar = new System.Windows.Forms.Timer { Interval = delay };
                        tBuscar.Tick += (tsB, teB) =>
                        {
                            ((System.Windows.Forms.Timer)tsB).Stop();
                            // Si ya avanzamos a otro ciclo, no re-clickar Buscar
                            if (_ripsGeneration != myGen) return;
                            string reBuscar = JsInRips(
                                "(function(){" +
                                "  var keys=Object.keys(window);" +
                                "  for(var i=0;i<keys.length;i++){" +
                                "    try{var obj=window[keys[i]];" +
                                "      if(obj&&typeof obj.DoClick==='function'&&" +
                                "         keys[i].toLowerCase().indexOf('buscar')>=0&&" +
                                "         keys[i].toLowerCase().indexOf('buscador')<0){" +
                                "        obj.DoClick();return 'DX_DOCLICK:'+keys[i];}" +
                                "    }catch(e){}" +
                                "  }" +
                                "  var btns=document.querySelectorAll('input[type=submit],input[type=button],button,a');" +
                                "  for(var j=0;j<btns.length;j++){" +
                                "    var v=(btns[j].value||btns[j].innerText||btns[j].title||'').toLowerCase().trim();" +
                                "    if(v==='buscar'||v.indexOf('buscar')>=0){btns[j].click();return 'BTN_CLICK:'+btns[j].id;}" +
                                "  }" +
                                "  return 'NF';" +
                                "})()");
                            WriteUiLog("RIPS grid reintento Buscar: " + reBuscar);
                            _awaitingBuscarResult = true;
                        };
                        tBuscar.Start();
                    };
                    tRetry.Start();
                }
                else
                {
                    // Agotados los reintentos — avanzar al siguiente registro
                    WriteUiLog("RIPS grid: maximo de reintentos alcanzado. Avanzando al siguiente registro.");
                    _buscarSinResultadoRetries = 0;
                    AdvanceToNextRipsRecord();
                }
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // Flujo formulario COMPLETAR RIPS
        // ─────────────────────────────────────────────────────────────────────

        private void StartCompletarFlow()
        {
            WriteUiLog("Iniciando flujo Completar RIPS step 0...");
            // Loguear todos los keys de window con DoClick o SetValue
            try
            {
                string keysLog = JsInRips(
                    "(function(){var r=[];for(var k in window){try{var o=window[k];if(o&&(typeof o.DoClick==='function'||typeof o.SetValue==='function'))r.push(k);}catch(e){}}return r.join('|').substring(0,2000);})()");
                WriteUiLog("Completar DX-keys: " + keysLog);
            }
            catch { }
            _completarStep = 0;
            if (_completarTimer != null) { _completarTimer.Stop(); _completarTimer.Dispose(); }
            _completarTimer = new System.Windows.Forms.Timer { Interval = 1200 };
            _completarTimer.Tick += CompletarTimer_Tick;
            _completarTimer.Start();
        }

        private void CompletarTimer_Tick(object sender, EventArgs e)
        {
            var timer = (System.Windows.Forms.Timer)sender;
            const string PFX    = "ctl00_ctl00_ContentPlaceHolder1_mainContent_";
            const string PFXBTN = "ctl00_ctl00_ContentPlaceHolder1_upButtonContent_";
            const string ID_FIN      = PFX + "ConsultaProcedimientoPanel_FinalidadComboBox";
            const string ID_DIAG_INP = PFX + "ConsultaProcedimientoPanel_DiagnosticosBuscador_ctl01_I";
            const string ID_DIAG_BTN = PFX + "ConsultaProcedimientoPanel_DiagnosticosBuscador_DiagnosticosBuscador_buscarBoton";
            const string ID_CAUSA    = PFX + "ConsultaPanel_CausaExternaComboBox";
            const string ID_GUARD    = PFXBTN + "GuardarButton";
            const string GRID_ID     = PFX + "ConsultaProcedimientoPanel_DiagnosticosDataGrid";
            try
            {
                var rec = _ripsRecords[_ripsRecordIndex];
                switch (_completarStep)
                {
                    case 0: // Finalidad — DevExpress ComboBox
                    {
                        string fin = EscapeJs(rec.Finalidad ?? "");
                        string r0 = JsInRips(
                            "(function(){" +
                            "  var fin='" + fin + "'.toLowerCase();" +
                            "  var cb=window['" + ID_FIN + "'];" +
                            "  if(cb&&typeof cb.GetItemCount==='function'){" +
                            "    for(var i=0;i<cb.GetItemCount();i++){var it=cb.GetItem(i);" +
                            "      if(it&&(it.text||'').toLowerCase().indexOf(fin)>=0){cb.SetValue(it.value);return 'FIN_OK:'+it.text;}" +
                            "    }" +
                            "    if(cb.GetItemCount()>1){cb.SetValue(cb.GetItem(1).value);return 'FIN_FIRST:'+cb.GetItem(1).text;}" +
                            "  }" +
                            "  return 'FIN_NF_cbNull='+(cb==null);" +
                            "})()");
                        WriteUiLog("Completar step0 Finalidad: " + r0);
                        _completarStep++; timer.Interval = 800;
                        break;
                    }
                    case 1: // Verificar si el grid ya tiene diagnostico; si no, escribir código en el campo
                    {
                        string diag = EscapeJs(rec.Diagnostico ?? "");
                        string r1 = JsInRips(
                            "(function(){" +
                            // Primero verificar si el grid Infragistics ya tiene filas
                            "  try{var g=igtbl_getGridById('" + GRID_ID + "');if(g&&g.Rows&&g.Rows.length>0)return 'GRID_YA_TIENE:'+g.Rows.length;}catch(ex){}" +
                            // Si no tiene filas, escribir el código en el input del buscador
                            "  var inp=document.getElementById('" + ID_DIAG_INP + "');" +
                            "  if(inp){inp.value='" + diag + "';try{inp.focus();}catch(e){}return 'DIAG_INP_OK';}" +
                            "  var el=document.querySelector('[id*=DiagnosticosBuscador] input[type=text],[id*=DiagnosticosBuscador] input:not([type])');" +
                            "  if(el){el.value='" + diag + "';return 'DIAG_INP_FALLBACK:'+el.id;}" +
                            "  return 'DIAG_INP_NF';" +
                            "})()");
                        WriteUiLog("Completar step1 DiagCheck/Inp: " + r1);
                        // Si el grid ya tiene filas, saltar directo al step de TipoDxPrincipal
                        if (r1.StartsWith("GRID_YA_TIENE")) { _completarStep = 4; timer.Interval = 400; }
                        else { _completarStep++; timer.Interval = 600; }
                        break;
                    }
                    case 2: // Click lupa diagnósticos
                    {
                        string r2 = JsInRips(
                            "(function(){" +
                            "  var btn=window['" + ID_DIAG_BTN + "'];" +
                            "  if(btn&&typeof btn.DoClick==='function'){btn.DoClick();return 'LUPA_OK';}" +
                            "  var el=document.querySelector('[id*=DiagnosticosBuscador][id*=buscarBoton]');" +
                            "  if(el){try{el.click();}catch(x){}return 'LUPA_CLICK:'+el.id;}" +
                            "  return 'LUPA_NF';" +
                            "})()");
                        WriteUiLog("Completar step2 Lupa: " + r2);
                        _completarStep++; timer.Interval = 4500; // esperar que el __doPostBack recargue el frame
                        break;
                    }
                    case 3: // Seleccionar primera fila del grid de resultados del buscador (Infragistics)
                    {
                        string diag3 = EscapeJs((rec.Diagnostico ?? "").ToUpperInvariant());
                        string r3 = JsInRips(
                            "(function(){" +
                            "  var txt='" + diag3 + "'.toLowerCase();" +
                            // El buscador usa un grid separado — buscar por ID que contenga 'Buscador' y 'Grid' o 'Result'
                            "  var sels=['[id*=BuscadorGrid] tr','[id*=ResultGrid] tr','[id*=DiagBusc] tr'," +
                            "            '[id*=BuscResult] tr','[id*=DiagnosticosBuscador] td a'," +
                            "            '[id*=DiagnosticosBuscador] tr[onclick]'];" +
                            "  for(var s=0;s<sels.length;s++){" +
                            "    var rows=document.querySelectorAll(sels[s]);" +
                            "    for(var i=0;i<rows.length;i++){" +
                            "      var rt=(rows[i].innerText||rows[i].textContent||'').toLowerCase().trim();" +
                            "      if(rt.length>1&&(txt===''||rt.indexOf(txt)>=0)){" +
                            "        var oc=rows[i].getAttribute('onclick')||'';if(oc){try{eval(oc);}catch(er){rows[i].click();}return 'SEL_OC:'+rt.substring(0,40);}" +
                            "        rows[i].click();return 'SEL_CLICK:'+rt.substring(0,40);" +
                            "      }" +
                            "    }" +
                            "  }" +
                            // Fallback: cualquier tr que contenga el código del diagnóstico
                            "  var allTr=document.querySelectorAll('tr');" +
                            "  for(var j=0;j<allTr.length;j++){" +
                            "    var rt2=(allTr[j].innerText||allTr[j].textContent||'').toLowerCase().trim();" +
                            "    if(txt!==''&&rt2.indexOf(txt)>=0&&rt2.length<200){" +
                            "      var oc2=allTr[j].getAttribute('onclick')||'';" +
                            "      if(oc2&&(oc2.indexOf('seleccion')>=0||oc2.indexOf('Seleccion')>=0||oc2.indexOf('igtbl')>=0)){" +
                            "        try{eval(oc2);}catch(e){allTr[j].click();}return 'TR_OC:'+rt2.substring(0,40);" +
                            "      }" +
                            "    }" +
                            "  }" +
                            "  return 'SEL_NF:trs='+document.querySelectorAll('tr').length;" +
                            "})()");
                        WriteUiLog("Completar step3 SelDiag: " + r3);
                        _completarStep++; timer.Interval = 2000; // esperar que el grid principal se actualice
                        break;
                    }
                    case 4: // TipoDiagnosticoPrincipal = "Confirmado repetido" (valor 3)
                    {
                        string r4 = JsInRips(
                            "(function(){" +
                            "  var grid=igtbl_getGridById('" + GRID_ID + "');" +
                            "  if(!grid) return 'GRID_NF';" +
                            "  var rows=grid.Rows;" +
                            "  if(!rows||rows.length===0) return 'GRID_EMPTY';" +
                            "  for(var i=0;i<rows.length;i++){" +
                            "    var row=rows.getRow(i);if(!row) continue;" +
                            "    var tipoDx;try{tipoDx=row.getCellFromKey('TipoDiagnostico');}catch(ex){}" +
                            "    if(tipoDx&&tipoDx.getValue().toString()==='0'){" +
                            "      var cell=row.getCellFromKey('TipoDiagnosticoPrincipal');" +
                            "      var idRow=row.getCellFromKey('Identificador').getValue();" +
                            // setValue('3') pone el valor y abre el dropdown en IE
                            "      cell.setValue('3');" +
                            // Buscar el <select> visible que quedó abierto y seleccionar "Confirmado repetido"
                            "      var sels=document.querySelectorAll('select');" +
                            "      for(var s=0;s<sels.length;s++){" +
                            "        var sel=sels[s];" +
                            "        for(var o=0;o<sel.options.length;o++){" +
                            "          if((sel.options[o].text||'').toLowerCase().indexOf('confirmado repetido')>=0){" +
                            "            sel.selectedIndex=o;" +
                            "            try{sel.fireEvent('onchange');}catch(e1){}" +  // IE: cerrar y commitear
                            "            try{PageMethods.ActualizarDiagnotico(idRow,'3','TipoDiagnosticoPrincipal',function(){},function(){});}catch(e2){}" +
                            "            return 'SEL_CONF_REPETIDO:'+sel.options[o].text+' idRow='+idRow;" +
                            "          }" +
                            "        }" +
                            "      }" +
                            // Fallback: solo PageMethods si no encontramos el select
                            "      try{PageMethods.ActualizarDiagnotico(idRow,'3','TipoDiagnosticoPrincipal',function(){},function(){});}catch(ex3){}" +
                            "      return 'PM_ONLY_3:idRow='+idRow;" +
                            "    }" +
                            "  }" +
                            "  return 'NO_PRINCIPAL_ROW:len='+rows.length;" +
                            "})()");
                        WriteUiLog("Completar step4 TipoDxPrincipal: " + r4);
                        _completarStep++; timer.Interval = 2000; // esperar commit + PageMethod
                        break;
                    }
                    case 5: // Causa externa — DevExpress ComboBox
                    {
                        string causa = EscapeJs(rec.CausaExterna ?? "");
                        string frameUrl = JsInRips("window.location.href");
                        WriteUiLog("Completar step5 frameURL=" + frameUrl + " causaExcel=" + (rec.CausaExterna ?? "(vacío)"));
                        string r5 = JsInRips(
                            "(function(){" +
                            // norm(): quitar tildes sin normalize() — IE no lo soporta
                            "  function norm(s){" +
                            "    return s.toLowerCase().trim()" +
                            "      .replace(/[áàäâã]/g,'a').replace(/[éèëê]/g,'e')" +
                            "      .replace(/[íìïî]/g,'i').replace(/[óòöôõ]/g,'o')" +
                            "      .replace(/[úùüû]/g,'u').replace(/ñ/g,'n').replace(/ç/g,'c');" +
                            "  }" +
                            "  var causa=norm('" + causa + "');" +
                            "  var cb=window['" + ID_CAUSA + "'];" +
                            "  if(!cb||typeof cb.GetItemCount!=='function') return 'CAUSA_NF_cbNull=true';" +
                            "  var opts=[];for(var i=0;i<cb.GetItemCount();i++){var it=cb.GetItem(i);if(it)opts.push(i+':'+it.text);}" +
                            "  var optsStr=opts.join('||').substring(0,600);" +
                            // 1) Exacta
                            "  for(var i=0;i<cb.GetItemCount();i++){var it=cb.GetItem(i);" +
                            "    if(it&&norm(it.text||'')===causa){cb.SetValue(it.value);return 'EXACT:'+it.text;}" +
                            "  }" +
                            // 2) ComboBox contiene todo el texto del Excel
                            "  for(var i=0;i<cb.GetItemCount();i++){var it=cb.GetItem(i);" +
                            "    if(it&&causa.length>3&&norm(it.text||'').indexOf(causa)>=0){cb.SetValue(it.value);return 'CB_CONT:'+it.text;}" +
                            "  }" +
                            // 3) Excel contiene todo el texto del ComboBox
                            "  for(var i=0;i<cb.GetItemCount();i++){var it=cb.GetItem(i);" +
                            "    var itext=norm(it&&it.text||'');" +
                            "    if(itext.length>5&&causa.indexOf(itext)>=0){cb.SetValue(it.value);return 'XL_CONT:'+it.text;}" +
                            "  }" +
                            // 4) Primera palabra larga del Excel está en el texto del ComboBox
                            "  var palabras=causa.split(/[\\s\\-]+/);" +
                            "  for(var p=0;p<palabras.length;p++){" +
                            "    if(palabras[p].length<5) continue;" +
                            "    for(var i=0;i<cb.GetItemCount();i++){var it=cb.GetItem(i);" +
                            "      if(it&&norm(it.text||'').indexOf(palabras[p])>=0){cb.SetValue(it.value);return 'WORD:'+palabras[p]+':'+it.text;}" +
                            "    }" +
                            "  }" +
                            "  return 'NOMATCH:causa='+causa+'||OPTS:'+optsStr;" +
                            "})()");
                        WriteUiLog("Completar step5 CausaExt: " + r5);
                        _completarStep++; timer.Interval = 800;
                        break;
                    }
                    case 6: // Guardar
                    {
                        string r6 = JsInRips(
                            "(function(){" +
                            "  var btn=window['" + ID_GUARD + "'];" +
                            "  if(btn&&typeof btn.DoClick==='function'){btn.DoClick();return 'GUARD_OK';}" +
                            "  var btns=document.querySelectorAll('input[type=button],input[type=submit],button');" +
                            "  for(var j=0;j<btns.length;j++){var v=(btns[j].value||btns[j].innerText||'').toLowerCase();" +
                            "    if(v.indexOf('guardar')>=0){btns[j].click();return 'GUARD_BTN:'+btns[j].id;}}" +
                            "  return 'GUARD_NF';" +
                            "})()");
                        WriteUiLog("Completar step6 Guardar: " + r6);
                        timer.Stop();
                        _awaitingBuscarResult = true;
                        break;
                    }
                    default:
                        timer.Stop();
                        break;
                }
            }
            catch (Exception ex)
            {
                WriteUiLog("Completar step " + _completarStep + " error: " + ex.Message);
            }
        }

        private void AdvanceToNextRipsRecord()
        {
            _ripsRecordIndex++;
            WriteUiLog(string.Format("Registro {0} completado. Total={1}", _ripsRecordIndex, _ripsRecords != null ? _ripsRecords.Count : 0));
            UpdateBotProgress();
            SaveProgressState();
            if (_ripsRecords != null && _ripsRecordIndex < _ripsRecords.Count)
            {
                // Resetear todos los flags de estado de flujo,
                // EXCEPTO _ripsFlowStarted que se mantiene en true durante el delay
                // para que eventos DocumentCompleted espurios (AJAX, sub-frames) no
                // re-inicien el flujo en la página vieja antes de la navegación.
                _ripsStep = 0;
                _ripsCcRetries = 0;
                _convenioRetries = 0;
                _buscarSinResultadoRetries = 0;
                _waitingForUserConvenio = false;
                _awaitingBuscarResult = false;
                _awaitingCompletarForm = false;
                _completarFormUrl = null;
                if (_ripsTimer != null) { try { _ripsTimer.Stop(); _ripsTimer.Dispose(); } catch { } _ripsTimer = null; }
                _ripsGeneration++;  // invalida tRetry/tBuscar/tPoll huerfanos del ciclo actual
                _gridEmptyPolls = 0;
                string nextCC = _ripsRecords[_ripsRecordIndex].CC ?? "?";
                WriteUiLog("Avanzando a siguiente CC=" + nextCC);
                var tNext = new System.Windows.Forms.Timer { Interval = 2500 };
                tNext.Tick += (tsN, teN) =>
                {
                    ((System.Windows.Forms.Timer)tsN).Stop();
                    const string ripsUrlN = "http://181.51.196.194/Panacea/FacturacionPaciente/Facturacion_Pacientes/ConsultarEstadoRipsForm.aspx?IdOpcion=828";
                    // Poner _ripsFlowStarted = false justo antes de navegar,
                    // así el DocumentCompleted de la nueva página arranca el flujo correctamente
                    _ripsFlowStarted = false;
                    bool navOk = false;
                    try
                    {
                        HtmlDocument mainDoc = webBrowser1.Document;
                        if (mainDoc != null)
                        {
                            foreach (HtmlWindow frame in mainDoc.Window.Frames)
                            {
                                try
                                {
                                    string furl = frame.Document?.Url?.ToString() ?? string.Empty;
                                    if (furl.IndexOf("ConsultarEstadoRips", StringComparison.OrdinalIgnoreCase) >= 0
                                        || furl.IndexOf("FacturacionPaciente", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        frame.Document.InvokeScript("eval", new object[] { "window.location.href='" + ripsUrlN + "';" });
                                        WriteUiLog("Nav siguiente: frame RIPS JS OK. CC=" + nextCC);
                                        navOk = true;
                                        break;
                                    }
                                }
                                catch { }
                            }
                            if (!navOk)
                            {
                                mainDoc.InvokeScript("eval", new object[] { "window.location.href='" + ripsUrlN + "';" });
                                WriteUiLog("Nav siguiente: mainDoc JS OK. CC=" + nextCC);
                                navOk = true;
                            }
                        }
                    }
                    catch { }
                    if (!navOk)
                    {
                        webBrowser1.Navigate(ripsUrlN);
                        WriteUiLog("Nav siguiente: webBrowser1.Navigate fallback. CC=" + nextCC);
                    }
                };
                tNext.Start();
            }
            else
            {
                _ripsFlowDone = true;
                WriteUiLog("=== Flujo RIPS completado. Todos los registros procesados. ===");
                UpdateBotProgress();
                SaveProgressState();
                UpdateSidebarStatus("Completado: todos los registros procesados.");
            }
        }

        private string GetRipsIframeOffset()
        {
            try
            {
                return (string)webBrowser1.Document.InvokeScript("eval", new object[] {
                    "(function(){" +
                    "  var all=document.querySelectorAll('iframe,frame');" +
                    "  for(var i=0;i<all.length;i++){" +
                    "    try{" +
                    "      var u=all[i].src||'';" +
                    "      try{u=all[i].contentWindow.location.href;}catch(e){}" +
                    "      if(u.indexOf('ConsultarEstadoRips')>=0){" +
                    "        var r=all[i].getBoundingClientRect();" +
                    "        return Math.round(r.left)+','+Math.round(r.top);" +
                    "      }" +
                    "    }catch(e){}" +
                    "  }" +
                    "  return '0,0';" +
                    "})()"
                }) ?? "0,0";
            }
            catch { return "0,0"; }
        }

        // ── NUEVOS MÉTODOS DE UI LATERAL ────────────────────────────────────

        // Seleccionar archivo Excel manualmente
        private void btnBrowseExcel_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "Seleccionar archivo base Excel";
                dlg.Filter = "Archivos Excel (*.xlsx)|*.xlsx";
                dlg.CheckFileExists = true;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _overrideExcelPath = dlg.FileName;
                    txtExcelPath.Text = System.IO.Path.GetFileName(dlg.FileName);
                    WriteUiLog("Archivo base seleccionado: " + dlg.FileName);
                    // Recargar lista de registros con el nuevo archivo
                    LoadRipsExcel();
                    int cntExcel = _ripsRecords?.Count ?? 0;
                    WriteUiLog("Registros cargados: " + cntExcel);
                    UpdateSidebarStatus(cntExcel > 0
                        ? cntExcel + " registros cargados del Excel."
                        : "Advertencia: el Excel no contiene registros validos.");
                    UpdateBotProgress();
                }
            }
        }

        // Correr Bot
        private void btnCorrerBot_Click(object sender, EventArgs e)
        {
            if (_ripsRecords == null || _ripsRecords.Count == 0)
            {
                if (!string.IsNullOrWhiteSpace(_overrideExcelPath))
                {
                    try { LoadRipsExcel(); UpdateBotProgress(); }
                    catch (Exception exXls) { WriteUiLog("LoadRipsExcel en bot click: " + exXls.Message); }
                }
                if (_ripsRecords == null || _ripsRecords.Count == 0)
                {
                    MessageBox.Show("Cargue primero un archivo base Excel (.xlsx).",
                        "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            if (_ripsFlowStarted && !_ripsFlowDone)
            {
                MessageBox.Show("El bot ya está corriendo.", "Aviso",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            _botPaused = false;
            _ripsRecordIndex = 0;
            _ripsFlowStarted = false;
            _ripsFlowDone = false;
            _ripsNavigated = false; // resetear para que 2da+ ejecucion navegue a RIPS
            UpdateSidebarStatus("Iniciando bot...");
            UpdateBotProgress();
            // Si la sede ya fue seleccionada, ir directo al modulo RIPS; si no, flujo completo desde login
            if (_sedeSelectionDone)
                webBrowser1.Navigate("http://181.51.196.194/Panacea/FacturacionPaciente/Facturacion_Pacientes/ConsultarEstadoRipsForm.aspx?IdOpcion=828");
            else
                webBrowser1.Navigate(PanaceaUrl);
        }

        // Pausar Bot
        private void btnPausarBot_Click(object sender, EventArgs e)
        {
            if (!_ripsFlowStarted || _ripsFlowDone)
            {
                MessageBox.Show("El bot no está corriendo.", "Aviso",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            _botPaused = true;
            _ripsTimer?.Stop();
            _completarTimer?.Stop();
            UpdateSidebarStatus("Bot pausado.");
            WriteUiLog("Bot pausado por el usuario.");
        }

        // Continuar (desde sidebar, para reanudar luego de pausa)
        private void btnContinuarSidebar_Click(object sender, EventArgs e)
        {
            if (_botPaused)
            {
                _botPaused = false;
                UpdateSidebarStatus("Reanudando bot...");
                WriteUiLog("Bot reanudado por el usuario.");
                // Reactivar el timer que corresponda
                if (_awaitingCompletarForm || _completarStep > 0)
                    _completarTimer?.Start();
                else if (_ripsFlowStarted && !_ripsFlowDone)
                    _ripsTimer?.Start();
                return;
            }
            // Si está esperando selección de convenio manual, delegar al método original
            if (_waitingForUserConvenio)
            {
                btnContinuarBot_Click(sender, e);
            }
        }

        // Generar informe Excel
        private void btnGenerarInforme_Click(object sender, EventArgs e)
        {
            try
            {
                var state = ProgressState.Load();
                string path = ReportGenerator.Generate(state);
                System.Diagnostics.Process.Start(path);
                WriteUiLog("Informe generado: " + path);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al generar el informe:\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // Reiniciar Bot — limpia todo el estado y recarga el Excel
        private void btnResetBot_Click(object sender, EventArgs e)
        {
            var res = MessageBox.Show(
                "Esto detiene el bot, limpia el estado guardado y recarga el archivo Excel.\n\n" +
                "Use esto cuando el bot quede trabado, en cache, o al cambiar de PC / archivo.\n\n" +
                "\u00bfDesea continuar?",
                "Reiniciar Bot", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (res != DialogResult.Yes) return;

            // 1. Detener todos los timers
            _ripsTimer?.Stop();
            _completarTimer?.Stop();
            _autoLoginTimer?.Stop();
            _authOutcomeTimer?.Stop();

            // 2. Resetear todos los flags de estado
            _ripsFlowStarted       = false;
            _ripsFlowDone          = false;
            _botPaused             = false;
            _waitingForUserConvenio = false;
            _ripsRecordIndex       = 0;
            _ripsRecords           = null;
            _sedeSelectionDone     = false;
            _ripsNavigated         = false;
            _awaitingCompletarForm = false;
            _completarStep         = 0;

            // 3. Limpiar estado persistido en disco
            ProgressState.Clear();

            // 4. Si el Excel seleccionado no existe en esta PC, limpiar la ruta
            if (!string.IsNullOrWhiteSpace(_overrideExcelPath) && !File.Exists(_overrideExcelPath))
            {
                WriteUiLog("Reset: archivo no existe en esta PC: " + _overrideExcelPath);
                _overrideExcelPath = string.Empty;
                txtExcelPath.Text  = string.Empty;
                MessageBox.Show(
                    "El archivo Excel de la sesion anterior no existe en esta PC.\n\nSeleccione uno nuevo con el boton ...",
                    "Archivo no encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // 5. Siempre intentar recargar registros (override path O carpeta 'base')
            try
            {
                LoadRipsExcel();
                int cnt = _ripsRecords?.Count ?? 0;
                WriteUiLog("Reset: " + cnt + " registros cargados.");
                if (cnt > 0)
                    UpdateSidebarStatus("Bot reiniciado. " + cnt + " registros listos.");
                else
                    UpdateSidebarStatus("Bot reiniciado. Sin registros — seleccione un archivo Excel.");
            }
            catch (Exception exXls) { WriteUiLog("Reset: error al recargar Excel: " + exXls.Message); }

            // 6. Actualizar UI
            UpdateBotProgress();
            WriteUiLog("Bot reiniciado. Indice=0. ProgressState borrado.");
        }
        private void btnAdmin_Click(object sender, EventArgs e)
        {
            var form = new AdminForm();
            form.Show(this);
        }

        // Actualizar barra de progreso y etiqueta
        private void UpdateBotProgress()
        {
            try
            {
                int total = _ripsRecords?.Count ?? 0;
                int done  = _ripsRecordIndex;
                lblProgress.Text = string.Format("Registros: {0} / {1}", done, total);
                if (total > 0)
                {
                    progressBar.Maximum = total;
                    progressBar.Value   = Math.Min(done, total);
                }
                else
                {
                    progressBar.Value = 0;
                }
            }
            catch { }
        }

        // Actualizar etiqueta de estado en la barra lateral
        private void UpdateSidebarStatus(string message)
        {
            try
            {
                lblStatus.Text = "  " + message;
            }
            catch { }
        }

        // Agregar una línea al log de la barra lateral
        private void AppendSidebarLog(string message)
        {
            try
            {
                string line = "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message;
                if (txtLog.InvokeRequired)
                {
                    txtLog.BeginInvoke(new Action(() => AppendLine(line)));
                }
                else
                {
                    AppendLine(line);
                }
            }
            catch { }
        }

        private void AppendLine(string line)
        {
            txtLog.AppendText(line + "\r\n");
            // Mantener scroll al final
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
        }

        // Guardar estado de progreso actual
        private void SaveProgressState()
        {
            try
            {
                var state = ProgressState.Load();
                state.TotalRecords = _ripsRecords?.Count ?? 0;
                state.LastUpdate   = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // Registrar el registro actual como procesado si aún no fue registrado
                if (_ripsRecords != null && _ripsRecordIndex > 0 &&
                    _ripsRecordIndex <= _ripsRecords.Count)
                {
                    var rec = _ripsRecords[_ripsRecordIndex - 1];
                    if (state.ProcessedRecords == null)
                        state.ProcessedRecords = new System.Collections.Generic.List<ProcessedRecord>();

                    bool exists = false;
                    foreach (var p in state.ProcessedRecords)
                    {
                        if (p.CC == rec.CC && p.FechaInicio == rec.FechaInicio)
                        { exists = true; break; }
                    }
                    if (!exists)
                    {
                        state.ProcessedRecords.Add(new ProcessedRecord
                        {
                            CC          = rec.CC,
                            FechaInicio = rec.FechaInicio,
                            FechaFin    = rec.FechaFin,
                            Convenio    = rec.TipoConvenio,
                            Estado      = "Completado",
                            Timestamp   = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                        });
                    }
                }
                state.Save();
            }
            catch { }
        }

    }
}
