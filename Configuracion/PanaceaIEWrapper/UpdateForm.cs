using System;
using System.Drawing;
using System.Windows.Forms;

namespace PanaceaIEWrapper
{
    internal sealed class UpdateForm : Form
    {
        public bool ShouldUpdate { get; private set; }

        public UpdateForm(string currentVersion, string newVersion)
        {
            SuspendLayout();

            Text = "Panacea RIPS";
            ClientSize = new Size(440, 222);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9.5f);
            AutoScaleMode = AutoScaleMode.None;

            // ── Barra superior azul ──────────────────────────────────
            var pnlHeader = new Panel
            {
                BackColor = Color.FromArgb(0, 84, 166),
                Dock = DockStyle.Top,
                Height = 56
            };
            var lblAppName = new Label
            {
                Text = "  Panacea RIPS",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(14, 0, 0, 0)
            };
            pnlHeader.Controls.Add(lblAppName);

            // ── Cuerpo ───────────────────────────────────────────────
            var lblTitle = new Label
            {
                Text = "Nueva versión disponible",
                Font = new Font("Segoe UI", 12f, FontStyle.Bold),
                ForeColor = Color.FromArgb(20, 20, 20),
                Location = new Point(24, 72),
                AutoSize = true
            };

            var lblCurrent = new Label
            {
                Text = "Versión instalada:   " + currentVersion,
                ForeColor = Color.Gray,
                Location = new Point(24, 110),
                AutoSize = true
            };

            var lblNew = new Label
            {
                Text = "Nueva versión:        " + newVersion,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 128, 0),
                Location = new Point(24, 132),
                AutoSize = true
            };

            // ── Separador ────────────────────────────────────────────
            var sep = new Panel
            {
                BackColor = Color.FromArgb(218, 218, 218),
                Location = new Point(0, 162),
                Size = new Size(440, 1)
            };

            // ── Botones ──────────────────────────────────────────────
            var btnUpdate = new Button
            {
                Text = "Actualizar ahora",
                Size = new Size(154, 38),
                Location = new Point(152, 172),
                BackColor = Color.FromArgb(0, 84, 166),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnUpdate.FlatAppearance.BorderSize = 0;
            btnUpdate.Click += (s, e) => { ShouldUpdate = true; Close(); };

            var btnLater = new Button
            {
                Text = "Más tarde",
                Size = new Size(104, 38),
                Location = new Point(318, 172),
                BackColor = Color.FromArgb(236, 236, 236),
                ForeColor = Color.FromArgb(60, 60, 60),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnLater.FlatAppearance.BorderSize = 0;
            btnLater.FlatAppearance.MouseOverBackColor = Color.FromArgb(210, 210, 210);
            btnLater.Click += (s, e) => { ShouldUpdate = false; Close(); };

            AcceptButton = btnUpdate;
            CancelButton = btnLater;

            Controls.Add(pnlHeader);
            Controls.Add(lblTitle);
            Controls.Add(lblCurrent);
            Controls.Add(lblNew);
            Controls.Add(sep);
            Controls.Add(btnUpdate);
            Controls.Add(btnLater);

            ResumeLayout();
        }
    }

    // ── Ventana de progreso de descarga ──────────────────────────────
    internal sealed class DownloadProgressForm : Form
    {
        public DownloadProgressForm()
        {
            Text = "Panacea RIPS";
            ClientSize = new Size(360, 96);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ControlBox = false;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9.5f);

            var lbl = new Label
            {
                Text = "Descargando actualización, por favor espere...",
                AutoSize = false,
                Size = new Size(320, 28),
                Location = new Point(20, 16),
                TextAlign = ContentAlignment.MiddleLeft
            };

            var pb = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Size = new Size(320, 18),
                Location = new Point(20, 56)
            };

            Controls.Add(lbl);
            Controls.Add(pb);
        }
    }
}
