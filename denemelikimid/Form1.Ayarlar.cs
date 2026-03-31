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
                Text = "⚙️ Hakkında\n\nBu Uygulama Bursa Uludağ Üniversitesi İdari Ve Mali İşler Daire Başkanlığı Adına Yapılmıştır.\n\n V1.0 By Tanrıverdi",
                Font = new Font("Segoe UI", 10, FontStyle.Italic),
                ForeColor = colorTextSecondary,
                AutoSize = true,
                Location = new Point(30, 30)
            };

            panelContent.Controls.Add(lblInfo);
        }

    }
}




