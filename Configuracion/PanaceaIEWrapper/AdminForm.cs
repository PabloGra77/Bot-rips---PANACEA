using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace PanaceaIEWrapper
{
    internal sealed class AdminForm : Form
    {
        private readonly ListView _lvRecords;
        private readonly Label _lblTotal;
        private readonly Label _lblDone;
        private readonly Label _lblPending;
        private readonly Label _lblLastUpdate;

        public AdminForm()
        {
            SuspendLayout();
            Text = "Panacea RIPS — Panel de Administracion";
            ClientSize = new Size(760, 520);
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(640, 400);
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9.5f);

            // Barra superior azul
            var pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 52,
                BackColor = Color.FromArgb(0, 84, 166)
            };
            var lblTitle = new Label
            {
                Text = "  Panel de Administracion",
                Dock = DockStyle.Fill,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlHeader.Controls.Add(lblTitle);

            // Fila de estadisticas
            var pnlStats = new Panel
            {
                Dock = DockStyle.Top,
                Height = 56,
                BackColor = Color.FromArgb(240, 244, 250),
                Padding = new Padding(16, 0, 16, 0)
            };
            _lblTotal   = MakeStatLabel("Total: —",        Color.FromArgb(40,  40,  40));
            _lblDone    = MakeStatLabel("Procesados: —",   Color.FromArgb(0,  128,  0));
            _lblPending = MakeStatLabel("Pendientes: —",   Color.FromArgb(180, 80,   0));
            _lblLastUpdate = MakeStatLabel("Ultima act.: —", Color.FromArgb(80, 80, 80));
            _lblTotal.Location      = new Point(16, 16);
            _lblDone.Location       = new Point(168, 16);
            _lblPending.Location    = new Point(336, 16);
            _lblLastUpdate.Location = new Point(510, 16);
            pnlStats.Controls.Add(_lblTotal);
            pnlStats.Controls.Add(_lblDone);
            pnlStats.Controls.Add(_lblPending);
            pnlStats.Controls.Add(_lblLastUpdate);

            // ListView
            _lvRecords = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9f)
            };
            _lvRecords.Columns.Add("#",            42);
            _lvRecords.Columns.Add("Cedula (CC)",  110);
            _lvRecords.Columns.Add("Fecha Inicio", 100);
            _lvRecords.Columns.Add("Fecha Fin",    100);
            _lvRecords.Columns.Add("Convenio",     140);
            _lvRecords.Columns.Add("Estado",        96);
            _lvRecords.Columns.Add("Fecha/Hora",   140);

            // Barra inferior
            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 52,
                BackColor = Color.FromArgb(242, 242, 242)
            };
            var sepBottom = new Panel
            {
                Dock = DockStyle.Top,
                Height = 1,
                BackColor = Color.FromArgb(210, 210, 210)
            };
            var btnReport   = MakeBtn("Generar Informe",  Color.FromArgb(0,  84, 166), new Point(14,  9));
            var btnRefresh  = MakeBtn("Actualizar",       Color.FromArgb(70, 70,  70), new Point(186, 9));
            var btnClear    = MakeBtn("Limpiar progreso", Color.FromArgb(190, 50, 50), new Point(358, 9));
            var btnClose    = MakeBtn("Cerrar",           Color.FromArgb(130,130,130), new Point(636, 9));
            btnClose.Width  = 110;

            btnReport.Click  += (s, e) => OnGenerateReport();
            btnRefresh.Click += (s, e) => LoadData();
            btnClear.Click   += (s, e) => OnClear();
            btnClose.Click   += (s, e) => Close();

            pnlBottom.Controls.Add(sepBottom);
            pnlBottom.Controls.Add(btnReport);
            pnlBottom.Controls.Add(btnRefresh);
            pnlBottom.Controls.Add(btnClear);
            pnlBottom.Controls.Add(btnClose);

            Controls.Add(_lvRecords);
            Controls.Add(pnlStats);
            Controls.Add(pnlBottom);
            Controls.Add(pnlHeader);

            ResumeLayout();
            LoadData();
        }

        private void LoadData()
        {
            try
            {
                var state = ProgressState.Load();
                int total   = state.TotalRecords;
                int done    = state.ProcessedRecords?.Count ?? 0;
                int pending = total - done;

                _lblTotal.Text      = "Total: " + total;
                _lblDone.Text       = "Procesados: " + done;
                _lblPending.Text    = "Pendientes: " + pending;
                _lblLastUpdate.Text = "Ultima act.: " + (string.IsNullOrWhiteSpace(state.LastUpdate) ? "—" : state.LastUpdate);

                _lvRecords.Items.Clear();
                if (state.ProcessedRecords != null)
                {
                    int n = 1;
                    foreach (var r in state.ProcessedRecords)
                    {
                        var item = new ListViewItem(n++.ToString());
                        item.SubItems.Add(r.CC);
                        item.SubItems.Add(r.FechaInicio);
                        item.SubItems.Add(r.FechaFin);
                        item.SubItems.Add(r.Convenio);
                        item.SubItems.Add(r.Estado);
                        item.SubItems.Add(r.Timestamp);

                        item.BackColor = r.Estado == "Completado" ? Color.FromArgb(198, 239, 206)
                                        : r.Estado == "Error"      ? Color.FromArgb(255, 199, 206)
                                        :                             Color.FromArgb(255, 235, 156);
                        _lvRecords.Items.Add(item);
                    }
                }
            }
            catch { }
        }

        private void OnGenerateReport()
        {
            try
            {
                var state = ProgressState.Load();
                string path = ReportGenerator.Generate(state);
                Process.Start(path);
                MessageBox.Show("Informe guardado en el escritorio:\n" + System.IO.Path.GetFileName(path),
                    "Informe generado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al generar el informe:\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void OnClear()
        {
            if (MessageBox.Show("Se borrara el historial de progreso guardado.\n¿Continuar?",
                    "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;
            ProgressState.Clear();
            LoadData();
        }

        // Helpers
        private static Label MakeStatLabel(string text, Color color)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                ForeColor = color,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                BackColor = Color.Transparent
            };
        }

        private static Button MakeBtn(string text, Color back, System.Drawing.Point location)
        {
            var btn = new Button
            {
                Text = text,
                Size = new Size(164, 34),
                Location = location,
                BackColor = back,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }
    }
}
