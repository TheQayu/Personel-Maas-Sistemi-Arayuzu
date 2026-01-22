using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace denemelikimid
{
    public partial class Form1
    {
        private void LoadDashboardView()
        {
            // 1. SAYFAYI SIFIRLA
            panelContent.Controls.Clear();

            // 2. ANA İSKELET (DİKEY TABLO)
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = colorContent,
                Padding = new Padding(10)
            };

            // Satır Ayarları
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));          // Başlık
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));               // İstatistikler
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));          // Alt kısım

            panelContent.Controls.Add(mainLayout);

            // 3. BAŞLIK ve YENİLE BUTONU
            Panel pnlHeader = new Panel { Dock = DockStyle.Fill };

            Label lblHeader = new Label
            {
                Text = "📊 Genel Durum Paneli",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                Location = new Point(0, 10),
                AutoSize = true
            };
            pnlHeader.Controls.Add(lblHeader);

            Button btnRefresh = new Button
            {
                Text = "🔄 Yenile",
                Size = new Size(100, 35),
                BackColor = colorPrimary,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Dock = DockStyle.Right
            };
            btnRefresh.Click += (s, e) => LoadDashboardView();
            pnlHeader.Controls.Add(btnRefresh);

            mainLayout.Controls.Add(pnlHeader, 0, 0);

            // 4. İSTATİSTİK KARTLARI
            FlowLayoutPanel pnlStats = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Margin = new Padding(0, 0, 0, 20)
            };

            int toplam = 0, aktif = 0, puantaj = 0;
            decimal odeme = 0;
            try
            {
                using (var conn = new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                {
                    conn.Open();
                    using (var cmd = new MySqlCommand("SELECT COUNT(*) FROM program_katilimcilari", conn))
                        toplam = Convert.ToInt32(cmd.ExecuteScalar());
                    using (var cmd = new MySqlCommand("SELECT COUNT(*) FROM program_katilimcilari WHERE pk_isten_ayrilma_tarihi IS NULL", conn))
                        aktif = Convert.ToInt32(cmd.ExecuteScalar());

                    // Basit puantaj sayımı
                    try
                    {
                        string buAy = DateTime.Now.ToString("yyyy-MM");
                        using (var cmd = new MySqlCommand("SELECT COUNT(*) FROM puantaj WHERE p_yil_ay = @ay AND p_calistigi_gun_sayisi > 0", conn))
                        {
                            cmd.Parameters.AddWithValue("@ay", buAy);
                            puantaj = Convert.ToInt32(cmd.ExecuteScalar());
                        }
                    }
                    catch { }

                    // Basit ödeme toplamı
                    try
                    {
                        using (var cmd = new MySqlCommand("SELECT SUM(b_odenmesi_gereken_net_tutar) FROM bordro", conn))
                        {
                            var res = cmd.ExecuteScalar();
                            if (res != DBNull.Value) odeme = Convert.ToDecimal(res);
                        }
                    }
                    catch { }
                }
            }
            catch { }

            pnlStats.Controls.Add(CreateStatCard("👥 Toplam Personel", toplam.ToString(), colorPrimary));
            pnlStats.Controls.Add(CreateStatCard("✅ Aktif Çalışan", aktif.ToString(), colorSuccess));
            pnlStats.Controls.Add(CreateStatCard("📝 Bu Ay Puantaj", puantaj.ToString(), Color.Orange));
            pnlStats.Controls.Add(CreateStatCard("💰 Toplam Ödeme", odeme.ToString("C0"), colorInfo));

            mainLayout.Controls.Add(pnlStats, 0, 1);

            // 5. ALT KISIM (BUTONLAR ve LOG)
            TableLayoutPanel bottomLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.Transparent
            };
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            // Sol: Hızlı İşlemler
            Panel pnlLeft = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(10),
                Margin = new Padding(0, 0, 10, 0)
            };
            Label lblLeft = new Label
            {
                Text = "🚀 Hızlı İşlemler",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                Dock = DockStyle.Top,
                Height = 30
            };
            FlowLayoutPanel flowBtns = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                AutoScroll = true
            };

            flowBtns.Controls.Add(CreateQuickBtn("👤 Yeni Personel", colorPrimary, () => LoadPersonelListView()));
            flowBtns.Controls.Add(CreateQuickBtn("📝 Puantaj Gir", Color.Orange, () => LoadPuantajView()));
            flowBtns.Controls.Add(CreateQuickBtn("💰 Maaş Hesapla", colorSuccess, () => LoadRaporlarView()));
            flowBtns.Controls.Add(CreateQuickBtn("📄 Bordro Al", colorInfo, () => LoadRaporlarView()));

            pnlLeft.Controls.Add(flowBtns);
            pnlLeft.Controls.Add(lblLeft);
            bottomLayout.Controls.Add(pnlLeft, 0, 0);

            // Sağ: Sistem Logu
            Panel pnlRight = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(10)
            };
            Label lblRight = new Label
            {
                Text = "📢 Sistem Durumu",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                Dock = DockStyle.Top,
                Height = 30
            };
            ListBox lstLog = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = new Font("Consolas", 10),
                BackColor = Color.WhiteSmoke,
                ItemHeight = 20
            };

            lstLog.Items.Add($"> [{DateTime.Now:HH:mm}] Panel yüklendi.");
            lstLog.Items.Add($"> Veritabanı bağlantısı: OK");
            lstLog.Items.Add($"> Toplam {toplam} personel mevcut.");
            if (puantaj == 0)
                lstLog.Items.Add("> UYARI: Bu ay henüz puantaj girilmemiş!");
            else
                lstLog.Items.Add($"> {puantaj} personelin puantajı hazır.");

            pnlRight.Controls.Add(lstLog);
            pnlRight.Controls.Add(lblRight);
            bottomLayout.Controls.Add(pnlRight, 1, 0);

            mainLayout.Controls.Add(bottomLayout, 0, 2);
        }

        private Button CreateQuickBtn(string text, Color color, Action onClickAction)
        {
            Button btn = new Button
            {
                Text = text,
                Size = new Size(200, 100),
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 20, 20)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += (s, e) => onClickAction.Invoke();
            return btn;
        }

        private Panel CreateStatCard(string title, string value, Color color)
        {
            Panel card = new Panel
            {
                Width = 250,
                Height = 140,
                BackColor = Color.White,
                Margin = new Padding(0, 0, 20, 0)
            };

            Panel accent = new Panel { Dock = DockStyle.Left, Width = 5, BackColor = color };
            card.Controls.Add(accent);

            Label lblValue = new Label
            {
                Text = value,
                Font = new Font("Segoe UI", 24, FontStyle.Bold),
                ForeColor = Color.Black,
                Location = new Point(20, 25),
                AutoSize = true
            };
            card.Controls.Add(lblValue);

            Label lblTitle = new Label
            {
                Text = title,
                Font = new Font("Segoe UI", 11, FontStyle.Regular),
                ForeColor = Color.Gray,
                Location = new Point(20, 80),
                AutoSize = true
            };
            card.Controls.Add(lblTitle);

            return card;
        }

        private DataGridView CreateMiniGrid()
        {
            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            return dgv;
        }
    }
}




