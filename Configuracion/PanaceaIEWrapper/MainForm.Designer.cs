using System.Windows.Forms;

namespace PanaceaIEWrapper
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private WebBrowser webBrowser1;
        private System.Windows.Forms.Panel panelConvenio;
        private System.Windows.Forms.Label lblConvenioMsg;
        private System.Windows.Forms.Button btnContinuarBot;

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.panelConvenio = new System.Windows.Forms.Panel();
            this.lblConvenioMsg = new System.Windows.Forms.Label();
            this.btnContinuarBot = new System.Windows.Forms.Button();
            this.panelConvenio.SuspendLayout();
            this.SuspendLayout();

            // panelConvenio — barra visible solo cuando bot espera convenio
            this.panelConvenio.BackColor = System.Drawing.Color.FromArgb(255, 243, 205);
            this.panelConvenio.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelConvenio.Height = 44;
            this.panelConvenio.Visible = false;
            this.panelConvenio.Controls.Add(this.lblConvenioMsg);
            this.panelConvenio.Controls.Add(this.btnContinuarBot);

            // lblConvenioMsg
            this.lblConvenioMsg.AutoSize = false;
            this.lblConvenioMsg.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblConvenioMsg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblConvenioMsg.Padding = new System.Windows.Forms.Padding(8, 0, 0, 0);
            this.lblConvenioMsg.Font = new System.Drawing.Font("Segoe UI", 9.5f);
            this.lblConvenioMsg.Text = "Seleccione el convenio y haga clic en Aceptar. Cuando la pagina cargue, presione el boton →";

            // btnContinuarBot
            this.btnContinuarBot.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnContinuarBot.Width = 160;
            this.btnContinuarBot.Text = "▶  Continuar bot";
            this.btnContinuarBot.Font = new System.Drawing.Font("Segoe UI", 9.5f, System.Drawing.FontStyle.Bold);
            this.btnContinuarBot.BackColor = System.Drawing.Color.FromArgb(0, 120, 215);
            this.btnContinuarBot.ForeColor = System.Drawing.Color.White;
            this.btnContinuarBot.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnContinuarBot.Click += new System.EventHandler(this.btnContinuarBot_Click);

            // webBrowser1
            this.webBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser1.Location = new System.Drawing.Point(0, 0);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(800, 406);
            this.webBrowser1.TabIndex = 0;
            this.webBrowser1.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser1_DocumentCompleted);

            this.panelConvenio.ResumeLayout(false);

            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.panelConvenio);
            this.Name = "MainForm";
            this.Text = "Panacea";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
        }
    }
}
