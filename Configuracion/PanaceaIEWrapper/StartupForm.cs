using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;

namespace PanaceaIEWrapper
{
    internal sealed class StartupForm : Form
    {
        public string SelectedExcelPath { get; private set; }
        public string PanaceaUsername   { get; private set; }
        public string PanaceaPassword   { get; private set; }

        private readonly ModernTextBox _txtExcel;
        private readonly ModernTextBox _txtUser;
        private readonly ModernTextBox _txtPass;
        private readonly Label         _lblError;

        // Colores del tema
        private static readonly Color C_BG        = Color.FromArgb(18,  18,  32);   // fondo oscuro
        private static readonly Color C_CARD       = Color.FromArgb(28,  28,  45);   // card
        private static readonly Color C_BORDER     = Color.FromArgb(55,  55,  80);   // borde sutil
        private static readonly Color C_ACCENT     = Color.FromArgb(99,  102, 241);  // violeta/indigo
        private static readonly Color C_ACCENT2    = Color.FromArgb(139, 92,  246);  // violeta
        private static readonly Color C_TEXT       = Color.FromArgb(230, 230, 250);  // texto principal
        private static readonly Color C_TEXT_DIM   = Color.FromArgb(130, 130, 160);  // texto atenuado
        private static readonly Color C_SUCCESS    = Color.FromArgb(52,  211, 153);  // verde esmeralda
        private static readonly Color C_ERROR_BG   = Color.FromArgb(60,  20,  30);
        private static readonly Color C_ERROR_TEXT = Color.FromArgb(252, 165, 165);

        public StartupForm()
        {
            SuspendLayout();
            Text            = "Panacea RIPS";
            FormBorderStyle = FormBorderStyle.None;   // sin bordes del SO — custom border
            StartPosition   = FormStartPosition.CenterScreen;
            ClientSize      = new Size(460, 540);
            MaximizeBox     = false;
            MinimizeBox     = false;
            BackColor       = C_BG;
            Font            = new Font("Segoe UI", 9.5f);

            // Borde redondeado custom (GDI)
            this.Paint += (s, e) =>
            {
                using (var pen = new Pen(C_ACCENT, 1.5f))
                    e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
            };

            // Arrastrar ventana con mouse sobre header
            bool dragging = false;
            Point dragStart = Point.Empty;

            // ── BARRA SUPERIOR (drag + close) ────────────────────────────────
            var pnlTop = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 40,
                BackColor = Color.FromArgb(22, 22, 38),
                Cursor    = Cursors.SizeAll
            };
            pnlTop.MouseDown += (s, e) => { dragging = true; dragStart = e.Location; };
            pnlTop.MouseMove += (s, e) =>
            {
                if (!dragging) return;
                Location = new Point(Left + e.X - dragStart.X, Top + e.Y - dragStart.Y);
            };
            pnlTop.MouseUp += (s, e) => dragging = false;

            var lblAppTitle = new Label
            {
                Text      = "  ⬡  PANACEA RIPS  v1.3.4",
                Dock      = DockStyle.Fill,
                ForeColor = C_TEXT_DIM,
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            lblAppTitle.MouseDown += (s, e) => { dragging = true; dragStart = e.Location; };
            lblAppTitle.MouseMove += (s, e) =>
            {
                if (!dragging) return;
                Location = new Point(Left + e.X - dragStart.X, Top + e.Y - dragStart.Y);
            };
            lblAppTitle.MouseUp += (s, e) => dragging = false;

            // Boton cerrar (X)
            var btnClose = new Button
            {
                Text      = "✕",
                Size      = new Size(40, 40),
                Dock      = DockStyle.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                ForeColor = C_TEXT_DIM,
                Font      = new Font("Segoe UI", 10f),
                Cursor    = Cursors.Hand,
                TabStop   = false
            };
            btnClose.FlatAppearance.BorderSize  = 0;
            btnClose.FlatAppearance.MouseOverBackColor = Color.FromArgb(200, 30, 30);
            btnClose.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            // Botón Admin (esquina superior derecha, junto al X)
            var btnAdmin = new Button
            {
                Text      = "⚙",
                Size      = new Size(40, 40),
                Dock      = DockStyle.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                ForeColor = C_TEXT_DIM,
                Font      = new Font("Segoe UI", 12f),
                Cursor    = Cursors.Hand,
                TabStop   = false
            };
            btnAdmin.FlatAppearance.BorderSize  = 0;
            btnAdmin.FlatAppearance.MouseOverBackColor = Color.FromArgb(55, 55, 80);
            btnAdmin.Click += (s, e) =>
            {
                var admin = new AdminForm();
                admin.ShowDialog(this);
            };
            var ttAdmin = new ToolTip();
            ttAdmin.SetToolTip(btnAdmin, "Panel de administración");

            pnlTop.Controls.Add(lblAppTitle);
            pnlTop.Controls.Add(btnClose);
            pnlTop.Controls.Add(btnAdmin);

            // ── HERO / LOGO ──────────────────────────────────────────────────
            var pnlHero = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 120,
                BackColor = Color.Transparent
            };
            pnlHero.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // Gradiente de fondo
                using (var br = new LinearGradientBrush(
                    new Rectangle(0, 0, pnlHero.Width, pnlHero.Height),
                    Color.FromArgb(99, 102, 241), Color.FromArgb(139, 92, 246),
                    LinearGradientMode.Horizontal))
                {
                    g.FillRectangle(br, 0, 0, pnlHero.Width, pnlHero.Height);
                }

                // Círculo decorativo
                using (var pen = new Pen(Color.FromArgb(60, 255, 255, 255), 1))
                {
                    g.DrawEllipse(pen, pnlHero.Width - 100, -40, 160, 160);
                    g.DrawEllipse(pen, pnlHero.Width - 60, 20, 80, 80);
                }

                // Icono robot
                using (var fIcon = new Font("Segoe UI Emoji", 32f))
                using (var br2 = new SolidBrush(Color.White))
                    g.DrawString("🤖", fIcon, br2, new PointF(28, 18));

                // Título
                using (var fTitle = new Font("Segoe UI", 22f, FontStyle.Bold))
                using (var br3 = new SolidBrush(Color.White))
                    g.DrawString("Panacea RIPS Bot", fTitle, br3, new PointF(100, 22));

                // Subtítulo
                using (var fSub = new Font("Segoe UI", 10f))
                using (var br4 = new SolidBrush(Color.FromArgb(200, 255, 255, 255)))
                    g.DrawString("Automatización de radicación RIPS", fSub, br4, new PointF(102, 72));
            };

            // ── CARD CENTRAL ─────────────────────────────────────────────────
            var pnlCard = new Panel
            {
                Dock      = DockStyle.Fill,
                BackColor = C_CARD,
                Padding   = new Padding(36, 28, 36, 24)
            };
            pnlCard.Paint += (s, e) =>
            {
                // línea top accent
                using (var br = new LinearGradientBrush(
                    new Rectangle(0, 0, pnlCard.Width, 2),
                    C_ACCENT, C_ACCENT2, LinearGradientMode.Horizontal))
                    e.Graphics.FillRectangle(br, 0, 0, pnlCard.Width, 2);
            };

            // ── CAMPO: ARCHIVO EXCEL ─────────────────────────────────────────
            int y = 18;

            var lblExcel = FieldLabel("📂  Archivo base (.xlsx)");
            lblExcel.Location = new Point(0, y);
            y += 26;

            var pnlExcelRow = new Panel
            {
                Location  = new Point(0, y),
                Size      = new Size(388, 38),
                BackColor = Color.Transparent
            };
            _txtExcel = new ModernTextBox(C_CARD, C_BORDER, C_ACCENT, C_TEXT)
            {
                Location = new Point(0, 0),
                Size     = new Size(344, 38),
                ReadOnly = true
            };
            SetHint(_txtExcel, "Haga clic en › para seleccionar...");

            var btnBrowse = AccentButton("›", C_ACCENT, 38, 38);
            btnBrowse.Location = new Point(350, 0);
            btnBrowse.Font     = new Font("Segoe UI", 16f, FontStyle.Bold);
            btnBrowse.Click   += BtnBrowse_Click;
            pnlExcelRow.Controls.Add(_txtExcel);
            pnlExcelRow.Controls.Add(btnBrowse);
            y += 46;

            // ── SEPARADOR ───────────────────────────────────────────────────
            var sep1 = CardSep(y); y += 20;

            // ── CAMPO: USUARIO ───────────────────────────────────────────────
            var lblUser = FieldLabel("👤  Usuario Panacea");
            lblUser.Location = new Point(0, y); y += 26;
            _txtUser = new ModernTextBox(C_CARD, C_BORDER, C_ACCENT, C_TEXT)
            {
                Location = new Point(0, y),
                Size     = new Size(388, 38)
            };
            SetHint(_txtUser, "Número de documento o usuario");
            y += 48;

            // ── CAMPO: CONTRASEÑA ────────────────────────────────────────────
            var lblPass = FieldLabel("🔒  Contraseña");
            lblPass.Location = new Point(0, y); y += 26;
            _txtPass = new ModernTextBox(C_CARD, C_BORDER, C_ACCENT, C_TEXT)
            {
                Location              = new Point(0, y),
                Size                  = new Size(388, 38),
                UseSystemPasswordChar = true
            };
            SetHint(_txtPass, "Contraseña de Panacea");
            y += 48;

            // ── ERROR LABEL ──────────────────────────────────────────────────
            _lblError = new Label
            {
                Location  = new Point(0, y),
                Size      = new Size(388, 26),
                BackColor = C_ERROR_BG,
                ForeColor = C_ERROR_TEXT,
                Font      = new Font("Segoe UI", 8.5f),
                TextAlign = ContentAlignment.MiddleCenter,
                Text      = string.Empty,
                Visible   = false,
                Padding   = new Padding(6, 0, 0, 0)
            };
            y += 34;

            // ── BOTÓN INICIAR ────────────────────────────────────────────────
            var btnStart = GradientStartButton();
            btnStart.Location = new Point(0, y);
            btnStart.Size     = new Size(388, 48);
            btnStart.Click   += BtnStart_Click;
            AcceptButton      = btnStart;
            y += 58;

            // ── FOOTER ───────────────────────────────────────────────────────
            var lblFooter = new Label
            {
                Text      = "v1.3.4",
                Location  = new Point(0, y),
                Size      = new Size(388, 18),
                ForeColor = C_TEXT_DIM,
                Font      = new Font("Segoe UI", 7.5f),
                TextAlign = ContentAlignment.MiddleCenter
            };

            pnlCard.Controls.Add(lblExcel);
            pnlCard.Controls.Add(pnlExcelRow);
            pnlCard.Controls.Add(sep1);
            pnlCard.Controls.Add(lblUser);
            pnlCard.Controls.Add(_txtUser);
            pnlCard.Controls.Add(lblPass);
            pnlCard.Controls.Add(_txtPass);
            pnlCard.Controls.Add(_lblError);
            pnlCard.Controls.Add(btnStart);
            pnlCard.Controls.Add(lblFooter);

            Controls.Add(pnlCard);
            Controls.Add(pnlHero);
            Controls.Add(pnlTop);

            LoadSavedCredentials();
            ResumeLayout();
        }

        // ── EVENTOS ──────────────────────────────────────────────────────────
        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title           = "Seleccionar archivo base Excel";
                dlg.Filter          = "Archivos Excel (*.xlsx)|*.xlsx";
                dlg.CheckFileExists = true;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _txtExcel.Text    = System.IO.Path.GetFileName(dlg.FileName);
                    _txtExcel.Tag     = dlg.FileName;
                    _txtExcel.ForeColor = C_TEXT;
                    _lblError.Visible = false;
                }
            }
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            string excelPath = _txtExcel.Tag as string ?? string.Empty;
            string user      = _txtUser.ForeColor == C_TEXT_DIM ? string.Empty : _txtUser.Text.Trim();
            string pass      = _txtPass.UseSystemPasswordChar && _txtPass.ForeColor == C_TEXT_DIM
                               ? string.Empty : _txtPass.Text;

            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            { ShowError("⚠  Seleccione un archivo Excel válido."); return; }
            if (string.IsNullOrWhiteSpace(user))
            { ShowError("⚠  Ingrese su usuario de Panacea."); _txtUser.Focus(); return; }
            if (string.IsNullOrWhiteSpace(pass))
            { ShowError("⚠  Ingrese su contraseña."); _txtPass.Focus(); return; }

            SelectedExcelPath = excelPath;
            PanaceaUsername   = user;
            PanaceaPassword   = pass;
            SaveCredentials(user);
            DialogResult = DialogResult.OK;
            Close();
        }

        // ── PERSISTENCIA ─────────────────────────────────────────────────────
        private static readonly string CredFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PanaceaRIPS", "cred.dat");

        private void LoadSavedCredentials()
        {
            try
            {
                if (!File.Exists(CredFile)) return;
                string[] lines = File.ReadAllLines(CredFile);
                if (lines.Length >= 1 && !string.IsNullOrWhiteSpace(lines[0]))
                {
                    _txtUser.Text      = lines[0];
                    _txtUser.ForeColor = C_TEXT;
                }
            }
            catch { }
        }

        private static void SaveCredentials(string user)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(CredFile));
                File.WriteAllLines(CredFile, new[] { user });
            }
            catch { }
        }

        // ── HELPERS ──────────────────────────────────────────────────────────
        private void ShowError(string msg)
        {
            _lblError.Text    = msg;
            _lblError.Visible = true;
        }

        private static Label FieldLabel(string text)
        {
            return new Label
            {
                Text      = text,
                AutoSize  = true,
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = C_TEXT_DIM
            };
        }

        private static Panel CardSep(int y)
        {
            return new Panel
            {
                Location  = new Point(0, y),
                Size      = new Size(388, 1),
                BackColor = C_BORDER
            };
        }

        private static void SetHint(ModernTextBox tb, string hint)
        {
            tb.Text      = hint;
            tb.ForeColor = C_TEXT_DIM;
            tb.GotFocus += (s, e) =>
            {
                if (tb.Text == hint) { tb.Text = string.Empty; tb.ForeColor = C_TEXT; }
            };
            tb.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(tb.Text)) { tb.Text = hint; tb.ForeColor = C_TEXT_DIM; }
            };
        }

        private static Button AccentButton(string text, Color back, int w, int h)
        {
            var btn = new Button
            {
                Text      = text,
                Size      = new Size(w, h),
                BackColor = back,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor    = Cursors.Hand,
                TabStop   = false
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        // Botón con gradiente pintado a mano
        private static Button GradientStartButton()
        {
            var btn = new Button
            {
                Text      = "▶   Iniciar Bot",
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = C_ACCENT,
                Font      = new Font("Segoe UI", 12f, FontStyle.Bold),
                Cursor    = Cursors.Hand,
                UseVisualStyleBackColor = false
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.Paint += (s, e) =>
            {
                var g  = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var rc = new Rectangle(0, 0, btn.Width, btn.Height);
                using (var br = new LinearGradientBrush(rc, C_ACCENT, C_ACCENT2, LinearGradientMode.Horizontal))
                    g.FillRectangle(br, rc);
                var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                using (var tb = new SolidBrush(Color.White))
                    g.DrawString(btn.Text, btn.Font, tb, rc, sf);
            };
            return btn;
        }
    }

    // ── CONTROL AUXILIAR: TextBox con borde coloreado al foco ─────────────
    internal sealed class ModernTextBox : TextBox
    {
        private readonly Color _bgColor;
        private readonly Color _borderNormal;
        private readonly Color _borderFocus;

        public ModernTextBox(Color bg, Color borderNormal, Color borderFocus, Color fg)
        {
            _bgColor      = bg;
            _borderNormal = borderNormal;
            _borderFocus  = borderFocus;
            BackColor     = bg;
            ForeColor     = fg;
            BorderStyle   = BorderStyle.FixedSingle;
            Font          = new Font("Segoe UI", 10f);
        }

        protected override void OnGotFocus(EventArgs e)
        {
            base.OnGotFocus(e);
            // Redraw parent to show focus border
            Parent?.Invalidate(new Rectangle(Left - 2, Top - 2, Width + 4, Height + 4));
        }

        protected override void OnLostFocus(EventArgs e)
        {
            base.OnLostFocus(e);
            Parent?.Invalidate(new Rectangle(Left - 2, Top - 2, Width + 4, Height + 4));
        }
    }
}
