using System.Drawing;
using System.Windows.Forms;

namespace PanaceaIEWrapper
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        // Browser area
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.Panel panelConvenio;
        private System.Windows.Forms.Label lblConvenioMsg;
        private System.Windows.Forms.Button btnContinuarBot;
        private System.Windows.Forms.Panel panelBrowser;

        // Sidebar controls
        private System.Windows.Forms.Panel panelSidebar;
        private System.Windows.Forms.TextBox txtExcelPath;
        private System.Windows.Forms.Button btnBrowseExcel;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button btnCorrerBot;
        private System.Windows.Forms.Button btnPausarBot;
        private System.Windows.Forms.Button btnContinuarSidebar;
        private System.Windows.Forms.Button btnGenerarInforme;
        private System.Windows.Forms.Button btnAdmin;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.TextBox txtLog;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();

            // ── SIDEBAR ──────────────────────────────────────────────────────
            this.panelSidebar = new System.Windows.Forms.Panel();
            this.panelSidebar.SuspendLayout();

            // Header bar (blue)
            var pnlHead = new System.Windows.Forms.Panel
            {
                Dock = System.Windows.Forms.DockStyle.Top,
                Height = 52,
                BackColor = System.Drawing.Color.FromArgb(0, 84, 166)
            };
            var lblTitle = new System.Windows.Forms.Label
            {
                Text = "  Panacea RIPS Bot",
                Dock = System.Windows.Forms.DockStyle.Fill,
                ForeColor = System.Drawing.Color.White,
                Font = new System.Drawing.Font("Segoe UI", 13f, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            pnlHead.Controls.Add(lblTitle);

            // Excel file picker section
            var pnlExcel = new System.Windows.Forms.Panel
            {
                Dock = System.Windows.Forms.DockStyle.Top,
                Height = 62,
                Padding = new System.Windows.Forms.Padding(8, 6, 8, 4)
            };
            var lblExcelLbl = new System.Windows.Forms.Label
            {
                Text = "Archivo base (.xlsx):",
                Dock = System.Windows.Forms.DockStyle.Top,
                Height = 18,
                Font = new System.Drawing.Font("Segoe UI", 8.5f),
                ForeColor = System.Drawing.Color.FromArgb(80, 80, 80)
            };
            var pnlExcelRow = new System.Windows.Forms.Panel { Dock = System.Windows.Forms.DockStyle.Fill };
            this.txtExcelPath = new System.Windows.Forms.TextBox
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                Font = new System.Drawing.Font("Segoe UI", 8.5f),
                ReadOnly = true,
                BackColor = System.Drawing.Color.FromArgb(248, 248, 248)
            };
            this.btnBrowseExcel = new System.Windows.Forms.Button
            {
                Text = "...",
                Dock = System.Windows.Forms.DockStyle.Right,
                Width = 36,
                FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(224, 224, 224),
                Font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold),
                Cursor = System.Windows.Forms.Cursors.Hand
            };
            this.btnBrowseExcel.FlatAppearance.BorderSize = 0;
            pnlExcelRow.Controls.Add(this.txtExcelPath);
            pnlExcelRow.Controls.Add(this.btnBrowseExcel);
            pnlExcel.Controls.Add(pnlExcelRow);
            pnlExcel.Controls.Add(lblExcelLbl);

            // Progress bar section
            var pnlProg = new System.Windows.Forms.Panel
            {
                Dock = System.Windows.Forms.DockStyle.Top,
                Height = 50,
                Padding = new System.Windows.Forms.Padding(8, 4, 8, 4)
            };
            this.lblProgress = new System.Windows.Forms.Label
            {
                Text = "Registros: 0 / 0",
                Dock = System.Windows.Forms.DockStyle.Top,
                Height = 20,
                Font = new System.Drawing.Font("Segoe UI", 8.5f),
                ForeColor = System.Drawing.Color.FromArgb(60, 60, 60)
            };
            this.progressBar = new System.Windows.Forms.ProgressBar
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                Style = System.Windows.Forms.ProgressBarStyle.Continuous,
                Value = 0,
                Minimum = 0,
                Maximum = 100
            };
            pnlProg.Controls.Add(this.progressBar);
            pnlProg.Controls.Add(this.lblProgress);

            // Buttons section
            var pnlBotones = new System.Windows.Forms.Panel
            {
                Dock = System.Windows.Forms.DockStyle.Top,
                Height = 204,
                Padding = new System.Windows.Forms.Padding(8, 6, 8, 0)
            };
            this.btnCorrerBot        = MakeSideBtn(">> Correr Bot",       System.Drawing.Color.FromArgb(40,  167,  69), 42);
            this.btnPausarBot        = MakeSideBtn("|| Pausar Bot",        System.Drawing.Color.FromArgb(255, 153,   0), 36);
            this.btnContinuarSidebar = MakeSideBtn(">> Continuar",         System.Drawing.Color.FromArgb(0,  120, 215), 36);
            this.btnGenerarInforme   = MakeSideBtn("[i] Generar Informe",  System.Drawing.Color.FromArgb(108, 117, 125), 36);
            this.btnAdmin            = MakeSideBtn("[A] Panel Admin",      System.Drawing.Color.FromArgb(52,  58,  64), 36);
            // Add in visual order (first = top)
            pnlBotones.Controls.Add(this.btnCorrerBot);
            pnlBotones.Controls.Add(this.btnPausarBot);
            pnlBotones.Controls.Add(this.btnContinuarSidebar);
            pnlBotones.Controls.Add(this.btnGenerarInforme);
            pnlBotones.Controls.Add(this.btnAdmin);

            // Status label
            this.lblStatus = new System.Windows.Forms.Label
            {
                Text = "  Esperando...",
                Dock = System.Windows.Forms.DockStyle.Top,
                Height = 30,
                Font = new System.Drawing.Font("Segoe UI", 8.5f),
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                ForeColor = System.Drawing.Color.FromArgb(60, 60, 60),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new System.Windows.Forms.Padding(8, 0, 0, 0)
            };

            // Log area label
            var lblLogTitle = new System.Windows.Forms.Label
            {
                Text = "  Registro de actividad:",
                Dock = System.Windows.Forms.DockStyle.Top,
                Height = 20,
                Font = new System.Drawing.Font("Segoe UI", 8f),
                ForeColor = System.Drawing.Color.FromArgb(100, 100, 100),
                BackColor = System.Drawing.Color.White
            };

            // Log text box
            this.txtLog = new System.Windows.Forms.TextBox
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
                Font = new System.Drawing.Font("Consolas", 7.5f),
                BackColor = System.Drawing.Color.FromArgb(30, 30, 30),
                ForeColor = System.Drawing.Color.FromArgb(180, 230, 180),
                BorderStyle = System.Windows.Forms.BorderStyle.None
            };

            // Separator line between sections
            var sepLine = new System.Windows.Forms.Panel
            {
                Dock = System.Windows.Forms.DockStyle.Top,
                Height = 1,
                BackColor = System.Drawing.Color.FromArgb(210, 210, 210)
            };

            // Assemble sidebar — Fill first, then Top controls in visual order
            this.panelSidebar.Controls.Add(this.txtLog);         // Fill
            this.panelSidebar.Controls.Add(lblLogTitle);          // Top (last visible top)
            this.panelSidebar.Controls.Add(this.lblStatus);       // Top
            this.panelSidebar.Controls.Add(pnlBotones);           // Top
            this.panelSidebar.Controls.Add(pnlProg);              // Top
            this.panelSidebar.Controls.Add(pnlExcel);             // Top
            this.panelSidebar.Controls.Add(sepLine);              // Top
            this.panelSidebar.Controls.Add(pnlHead);              // Top (topmost)

            this.panelSidebar.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelSidebar.Width = 265;
            this.panelSidebar.BackColor = System.Drawing.Color.White;
            this.panelSidebar.ResumeLayout(false);

            // ── BROWSER PANEL ─────────────────────────────────────────────────
            this.panelBrowser = new System.Windows.Forms.Panel
            {
                Dock = System.Windows.Forms.DockStyle.Fill
            };

            // convenio wait bar
            this.panelConvenio = new System.Windows.Forms.Panel
            {
                BackColor = System.Drawing.Color.FromArgb(255, 243, 205),
                Dock = System.Windows.Forms.DockStyle.Bottom,
                Height = 44,
                Visible = false
            };
            this.lblConvenioMsg = new System.Windows.Forms.Label
            {
                AutoSize = false,
                Dock = System.Windows.Forms.DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new System.Windows.Forms.Padding(8, 0, 0, 0),
                Font = new System.Drawing.Font("Segoe UI", 9.5f),
                Text = "Seleccione el convenio y haga clic en Aceptar. Cuando la pagina cargue, presione el boton ->"
            };
            this.btnContinuarBot = new System.Windows.Forms.Button
            {
                Dock = System.Windows.Forms.DockStyle.Right,
                Width = 160,
                Text = ">> Continuar bot",
                Font = new System.Drawing.Font("Segoe UI", 9.5f, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.FromArgb(0, 120, 215),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = System.Windows.Forms.FlatStyle.Flat
            };
            this.btnContinuarBot.FlatAppearance.BorderSize = 0;
            this.btnContinuarBot.Click += new System.EventHandler(this.btnContinuarBot_Click);

            this.panelConvenio.Controls.Add(this.lblConvenioMsg);
            this.panelConvenio.Controls.Add(this.btnContinuarBot);

            // WebBrowser
            this.webBrowser1 = new System.Windows.Forms.WebBrowser
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                MinimumSize = new System.Drawing.Size(20, 20),
                Name = "webBrowser1",
                TabIndex = 0
            };
            this.webBrowser1.DocumentCompleted +=
                new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser1_DocumentCompleted);

            this.panelBrowser.Controls.Add(this.webBrowser1);
            this.panelBrowser.Controls.Add(this.panelConvenio);

            // ── FORM ──────────────────────────────────────────────────────────
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1220, 700);
            this.MinimumSize = new System.Drawing.Size(900, 500);
            this.Controls.Add(this.panelBrowser);
            this.Controls.Add(this.panelSidebar);
            this.Name = "MainForm";
            this.Text = "Panacea RIPS Bot";
            this.Load += new System.EventHandler(this.MainForm_Load);

            // Wire up sidebar button events
            this.btnCorrerBot.Click        += new System.EventHandler(this.btnCorrerBot_Click);
            this.btnPausarBot.Click        += new System.EventHandler(this.btnPausarBot_Click);
            this.btnContinuarSidebar.Click += new System.EventHandler(this.btnContinuarSidebar_Click);
            this.btnBrowseExcel.Click      += new System.EventHandler(this.btnBrowseExcel_Click);
            this.btnGenerarInforme.Click   += new System.EventHandler(this.btnGenerarInforme_Click);
            this.btnAdmin.Click            += new System.EventHandler(this.btnAdmin_Click);
        }

        private static System.Windows.Forms.Button MakeSideBtn(
            string text, System.Drawing.Color back, int height)
        {
            var btn = new System.Windows.Forms.Button
            {
                Text = text,
                Dock = System.Windows.Forms.DockStyle.Top,
                Height = height,
                BackColor = back,
                ForeColor = System.Drawing.Color.White,
                FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Segoe UI", 9.5f, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new System.Windows.Forms.Padding(10, 0, 0, 0),
                Cursor = System.Windows.Forms.Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }
    }
}
