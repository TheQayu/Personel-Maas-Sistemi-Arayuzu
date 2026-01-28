using System.Drawing;
using System.Windows.Forms;

namespace denemelikimid
{
    public partial class Form1
    {
        private void LoadAyarlarView()
        {
            panelContent.Controls.Clear();

            Label lblInfo = new Label
            {
                Text = "⚙️ Ayarlar\n\nSistem ayarları bu bölümde yapılandırılacaktır.",
                Font = new Font("Segoe UI", 14, FontStyle.Regular),
                ForeColor = colorTextPrimary,
                AutoSize = true,
                Location = new Point(30, 30)
            };

            panelContent.Controls.Add(lblInfo);
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(282, 253);
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        private void Form1_Load(object sender, System.EventArgs e)
        {

        }
    }
}




