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

    }
}




