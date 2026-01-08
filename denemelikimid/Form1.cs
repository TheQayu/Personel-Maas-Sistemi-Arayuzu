using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using denemelikimid.DataBase;
using System.Globalization;
using MySql.Data.MySqlClient;

namespace denemelikimid
{
    public partial class Form1 : Form
    {
        private DateTime secilenTarih = DateTime.Now;
        // Renkler
        private Color colorPrimary = Color.FromArgb(67, 97, 238);      // Modern Mavi                                                                 
        private Color colorPrimaryDark = Color.FromArgb(52, 76, 186);
        private Color colorSecondary = Color.FromArgb(255, 255, 255);  // Beyaz
        private Color colorSidebar = Color.FromArgb(30, 41, 59);       // Koyu Gri-Mavi
        private Color colorSidebarHover = Color.FromArgb(51, 65, 85);
        private Color colorHeader = Color.FromArgb(248, 250, 252);    // Açık Gri
        private Color colorContent = Color.FromArgb(241, 245, 249);    // Çok Açık Gri
        private Color colorSuccess = Color.FromArgb(34, 197, 94);      // Yeşil
        private Color colorInfo = Color.FromArgb(59, 130, 246);        // Açık Mavi
        private Color colorDanger = Color.FromArgb(239, 68, 68);       // Kırmızı
        private Color colorTextPrimary = Color.FromArgb(15, 23, 42);   // Koyu Metin
        private Color colorTextSecondary = Color.FromArgb(100, 116, 139); // Gri Metin

        // UI Bileşenleri
        private Panel panelSidebar;
        private Panel panelHeader;
        private Panel panelContent;
        private Label lblTitle;
        private Label lblSubtitle;
        private Button btnLogout;
        private Panel panelLogo;
        private DataGridView dgvMain;
        private string currentView = "Dashboard";

        public Form1()
        {
            SetupForm();
            InitializeUI();
            this.Resize += Form1_Resize;

        }

        private void SetupForm()
        {
            this.Text = "Üniversite Personel Takip Sistemi";
            this.Size = new Size(1400, 800);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.MinimumSize = new Size(1000, 600);
            this.BackColor = colorContent;
        }

        private void InitializeUI()
        {

            CreateSidebar();


            CreateHeader();


            CreateContentArea();


            LoadDashboardView();


            panelHeader.SendToBack();


            panelSidebar.SendToBack();


            panelContent.BringToFront();
        }

        private void CreateSidebar()
        {

            panelSidebar = new Panel();
            panelSidebar.Dock = DockStyle.Left;
            panelSidebar.Width = 280;
            panelSidebar.BackColor = colorSidebar;
            panelSidebar.Padding = new Padding(0);
            this.Controls.Add(panelSidebar);

            // Logo Bölümü
            panelLogo = new Panel();
            panelLogo.Dock = DockStyle.Top;
            panelLogo.Height = 120;
            panelLogo.BackColor = colorPrimary;
            panelLogo.Padding = new Padding(20, 20, 20, 20);
            panelSidebar.Controls.Add(panelLogo);

            Label lblLogo = new Label();
            lblLogo.Text = "🏛️\nÜNİVERSİTE\nPERSONEL SİSTEMİ";
            lblLogo.ForeColor = Color.White;
            lblLogo.Font = new Font("Segoe UI", 14, FontStyle.Bold);
            lblLogo.TextAlign = ContentAlignment.MiddleCenter;
            lblLogo.Dock = DockStyle.Fill;
            lblLogo.AutoSize = false;
            panelLogo.Controls.Add(lblLogo);

            // Menü Butonları Container
            Panel panelMenu = new Panel();
            panelMenu.Dock = DockStyle.Fill;
            panelMenu.BackColor = Color.Transparent;
            panelMenu.Padding = new Padding(15, 20, 15, 20);
            panelSidebar.Controls.Add(panelMenu);

            // Menü Butonları


            AddMenuButton(panelMenu, "📊 Ana Sayfa", "Dashboard", true);
            AddMenuButton(panelMenu, "👥 Personel Listesi", "PersonelListesi");
            AddMenuButton(panelMenu, "📝 Puantaj Girişi", "Puantaj");
            AddMenuButton(panelMenu, "📊 Raporlar & Bordro", "Raporlar");
            AddMenuButton(panelMenu, "📄 Excel İşlemleri", "Excel");
            AddMenuButton(panelMenu, "⚙️ Ayarlar", "Ayarlar");

            // Alt Kısım - Kullanıcı Bilgisi
            Panel panelUser = new Panel();
            panelUser.Dock = DockStyle.Bottom;
            panelUser.Height = 80;
            panelUser.BackColor = Color.FromArgb(20, 30, 45);
            panelUser.Padding = new Padding(15);
            panelSidebar.Controls.Add(panelUser);

            Label lblUser = new Label();
            lblUser.Text = "👤 Admin Kullanıcı";
            lblUser.ForeColor = Color.White;
            lblUser.Font = new Font("Segoe UI", 10, FontStyle.Regular);
            lblUser.Dock = DockStyle.Fill;
            lblUser.TextAlign = ContentAlignment.MiddleLeft;
            panelUser.Controls.Add(lblUser);
            panelLogo.SendToBack();
            panelUser.SendToBack();



            panelSidebar.Controls.Cast<Control>().FirstOrDefault(c => c.Dock == DockStyle.Fill)?.BringToFront();
        }

        private void AddMenuButton(Panel parent, string text, string viewName, bool isActive = false)
        {
            Button btn = new Button();
            btn.Height = 55;
            btn.Dock = DockStyle.Top;
            btn.Text = text;
            btn.TextAlign = ContentAlignment.MiddleLeft;
            btn.Padding = new Padding(20, 0, 0, 0);
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.ForeColor = isActive ? Color.White : Color.FromArgb(200, 200, 200);
            btn.BackColor = isActive ? colorPrimary : Color.Transparent;
            btn.Font = new Font("Segoe UI", 11, FontStyle.Regular);
            btn.Cursor = Cursors.Hand;
            btn.TextAlign = ContentAlignment.MiddleLeft;
            btn.Margin = new Padding(0, 0, 0, 8);

            if (isActive)
            {
                btn.BackColor = colorPrimary;
            }

            btn.Click += (s, e) =>
            {
                // Tüm butonları sıfırla
                foreach (Control ctrl in parent.Controls)
                {
                    if (ctrl is Button)
                    {
                        ctrl.BackColor = Color.Transparent;
                        ctrl.ForeColor = Color.FromArgb(200, 200, 200);
                    }
                }
                // Aktif butonu işaretle
                btn.BackColor = colorPrimary;
                btn.ForeColor = Color.White;

                currentView = viewName;
                lblTitle.Text = text.Replace("📊 ", "").Replace("👥 ", "").Replace("📝 ", "").Replace("📄 ", "").Replace("⚙️ ", "");
                LoadView(viewName);
            };

            btn.MouseEnter += (s, e) =>
            {
                if (btn.BackColor != colorPrimary)
                {
                    btn.BackColor = colorSidebarHover;
                    btn.ForeColor = Color.White;
                }
            };

            btn.MouseLeave += (s, e) =>
            {
                if (btn.BackColor != colorPrimary)
                {
                    btn.BackColor = Color.Transparent;
                    btn.ForeColor = Color.FromArgb(200, 200, 200);
                }
            };

            parent.Controls.Add(btn);
            btn.BringToFront();
        }

        private void CreateHeader()
        {
            panelHeader = new Panel();
            panelHeader.Dock = DockStyle.Top;
            panelHeader.Height = 80;
            panelHeader.BackColor = colorHeader;
            panelHeader.Padding = new Padding(30, 0, 30, 0);
            this.Controls.Add(panelHeader);

            //panelHeader.BringToFront();

            // Başlık
            Panel panelTitle = new Panel();
            panelTitle.Dock = DockStyle.Left;
            panelTitle.Width = 400;
            panelTitle.BackColor = Color.Transparent;
            panelHeader.Controls.Add(panelTitle);

            lblTitle = new Label();
            lblTitle.Text = "Ana Sayfa";
            lblTitle.Font = new Font("Segoe UI", 20, FontStyle.Bold);
            lblTitle.ForeColor = colorTextPrimary;
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(0, 15);
            panelTitle.Controls.Add(lblTitle);

            lblSubtitle = new Label();
            lblSubtitle.Text = "Sisteme hoş geldiniz";
            lblSubtitle.Font = new Font("Segoe UI", 10, FontStyle.Regular);
            lblSubtitle.ForeColor = colorTextSecondary;
            lblSubtitle.AutoSize = true;
            lblSubtitle.Location = new Point(0, 55);
            panelTitle.Controls.Add(lblSubtitle);

            // Sağ Taraf - Çıkış Butonu
            Panel panelActions = new Panel();
            panelActions.Dock = DockStyle.Right;
            panelActions.Width = 200;
            panelActions.BackColor = Color.Transparent;
            panelHeader.Controls.Add(panelActions);

            btnLogout = new Button();
            btnLogout.Text = "🚪 Çıkış Yap";
            btnLogout.Size = new Size(150, 45);
            btnLogout.Location = new Point(25, 17);
            btnLogout.BackColor = colorDanger;
            btnLogout.ForeColor = Color.White;
            btnLogout.FlatStyle = FlatStyle.Flat;
            btnLogout.FlatAppearance.BorderSize = 0;
            btnLogout.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            btnLogout.Cursor = Cursors.Hand;
            btnLogout.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            btnLogout.MouseEnter += (s, e) =>
            {
                btnLogout.BackColor = Color.FromArgb(220, 38, 38);
            };
            btnLogout.MouseLeave += (s, e) =>
            {
                btnLogout.BackColor = colorDanger;
            };

            panelActions.Controls.Add(btnLogout);
        }

        private void CreateContentArea()
        {
            panelContent = new Panel();
            panelContent.Dock = DockStyle.Fill;
            panelContent.BackColor = colorContent;
            panelContent.Padding = new Padding(30);
            this.Controls.Add(panelContent);
            //panelContent.SendToBack();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            // Responsive ayarlamalar burada yapılabilir
        }

        private void LoadView(string viewName)
        {
            panelContent.Controls.Clear();

            switch (viewName)
            {
                case "Dashboard":
                    LoadDashboardView();
                    break;
                case "PersonelListesi":
                    LoadPersonelListView();
                    break;
                case "Puantaj":
                    LoadPuantajView();
                    break;
                case "Excel":
                    LoadExcelView();
                    break;
                case "Ayarlar":
                    LoadAyarlarView();
                    break;
                case "Raporlar":
                    LoadRaporlarView();
                    break;
            }
        }
        // --- BU BLOĞU KOMPLE KOPYALA VE ESKİSİNİN YERİNE YAPIŞTIR ---

        private void LoadDashboardView()
        {
            // --- 1. SAYFAYI SIFIRLA ---
            panelContent.Controls.Clear();

            // --- 2. ANA İSKELET (DİKEY TABLO) ---
            // Sayfayı alt alta 3 satıra bölüyoruz. Bu yapı kaymayı engeller.
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = colorContent,
                Padding = new Padding(10)
            };

            // Satır Ayarları:
            // 1. Satır: Başlık (Sabit 60px)
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));
            // 2. Satır: İstatistikler (İçeriği kadar yer kaplasın - Auto)
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            // 3. Satır: Alt Kısım (Kalan tüm alanı kaplasın - %100)
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            panelContent.Controls.Add(mainLayout);

            // --- 3. SATIR 1: BAŞLIK ve YENİLE BUTONU ---
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

            // Başlık panelini ana tablonun 0. satırına ekle
            mainLayout.Controls.Add(pnlHeader, 0, 0);


            // --- 4. SATIR 2: İSTATİSTİK KARTLARI ---
            FlowLayoutPanel pnlStats = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Margin = new Padding(0, 0, 0, 20) // Alt kısımla arasına boşluk
            };

            // Verileri Çek
            int toplam = 0, aktif = 0, puantaj = 0; decimal odeme = 0;
            try
            {
                using (var conn = new MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                {
                    conn.Open();
                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand("SELECT COUNT(*) FROM program_katilimcilari", conn)) toplam = Convert.ToInt32(cmd.ExecuteScalar());
                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand("SELECT COUNT(*) FROM program_katilimcilari WHERE pk_isten_ayrilma_tarihi IS NULL", conn)) aktif = Convert.ToInt32(cmd.ExecuteScalar());

                    // Basit puantaj sayımı
                    try
                    {
                        string buAy = DateTime.Now.ToString("yyyy-MM");
                        using (var cmd = new MySql.Data.MySqlClient.MySqlCommand("SELECT COUNT(*) FROM puantaj WHERE p_yil_ay = @ay AND p_calistigi_gun_sayisi > 0", conn))
                        {
                            cmd.Parameters.AddWithValue("@ay", buAy);
                            puantaj = Convert.ToInt32(cmd.ExecuteScalar());
                        }
                    }
                    catch { }

                    // Basit ödeme toplamı
                    try
                    {
                        using (var cmd = new MySql.Data.MySqlClient.MySqlCommand("SELECT SUM(b_odenmesi_gereken_net_tutar) FROM bordro", conn))
                        {
                            var res = cmd.ExecuteScalar();
                            if (res != DBNull.Value) odeme = Convert.ToDecimal(res);
                        }
                    }
                    catch { }
                }
            }
            catch { }

            // Kartları ekle (Yardımcı metodun class içinde olduğunu varsayıyoruz)
            pnlStats.Controls.Add(CreateStatCard("👥 Toplam Personel", toplam.ToString(), colorPrimary));
            pnlStats.Controls.Add(CreateStatCard("✅ Aktif Çalışan", aktif.ToString(), colorSuccess));
            pnlStats.Controls.Add(CreateStatCard("📝 Bu Ay Puantaj", puantaj.ToString(), Color.Orange));
            pnlStats.Controls.Add(CreateStatCard("💰 Toplam Ödeme", odeme.ToString("C0"), colorInfo));

            // İstatistik panelini ana tablonun 1. satırına ekle
            mainLayout.Controls.Add(pnlStats, 0, 1);


            // --- 5. SATIR 3: ALT KISIM (BUTONLAR ve LOG) ---
            // Burayı da kendi içinde ikiye bölen bir tablo yapıyoruz (Sol ve Sağ)
            TableLayoutPanel bottomLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.Transparent
            };
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F)); // Sol %50
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F)); // Sağ %50

            // -- SOL TARAFA (BUTONLAR) --
            Panel pnlLeft = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(10), Margin = new Padding(0, 0, 10, 0) };
            Label lblLeft = new Label { Text = "🚀 Hızlı İşlemler", Font = new Font("Segoe UI", 12, FontStyle.Bold), ForeColor = colorTextPrimary, Dock = DockStyle.Top, Height = 30 };
            FlowLayoutPanel flowBtns = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, AutoScroll = true };

            // Butonları ekle (CreateQuickBtn metodu aşağıda olmalı)
            flowBtns.Controls.Add(CreateQuickBtn("👤 Yeni Personel", colorPrimary, () => LoadPersonelListView()));
            flowBtns.Controls.Add(CreateQuickBtn("📝 Puantaj Gir", Color.Orange, () => LoadPuantajView()));
            flowBtns.Controls.Add(CreateQuickBtn("💰 Maaş Hesapla", colorSuccess, () => LoadRaporlarView()));
            flowBtns.Controls.Add(CreateQuickBtn("📄 Bordro Al", colorInfo, () => LoadRaporlarView()));

            pnlLeft.Controls.Add(flowBtns);
            pnlLeft.Controls.Add(lblLeft);
            bottomLayout.Controls.Add(pnlLeft, 0, 0); // Sol hücreye ekle

            // -- SAĞ TARAFA (SİSTEM LOGU) --
            Panel pnlRight = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(10) };
            Label lblRight = new Label { Text = "📢 Sistem Durumu", Font = new Font("Segoe UI", 12, FontStyle.Bold), ForeColor = colorTextPrimary, Dock = DockStyle.Top, Height = 30 };
            ListBox lstLog = new ListBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, Font = new Font("Consolas", 10), BackColor = Color.WhiteSmoke, ItemHeight = 20 };

            lstLog.Items.Add($"> [{DateTime.Now:HH:mm}] Panel yüklendi.");
            lstLog.Items.Add($"> Veritabanı bağlantısı: OK");
            lstLog.Items.Add($"> Toplam {toplam} personel mevcut.");
            if (puantaj == 0) lstLog.Items.Add("> UYARI: Bu ay henüz puantaj girilmemiş!");
            else lstLog.Items.Add($"> {puantaj} personelin puantajı hazır.");

            pnlRight.Controls.Add(lstLog);
            pnlRight.Controls.Add(lblRight);
            bottomLayout.Controls.Add(pnlRight, 1, 0); // Sağ hücreye ekle

            // Alt düzeni ana tablonun 2. satırına (en alta) ekle
            mainLayout.Controls.Add(bottomLayout, 0, 2);
        }

        // --- YARDIMCI METOT: HIZLI BUTON OLUŞTURUCU ---
        // Bunu LoadDashboardView'in dışına, en alta ekle
        private Button CreateQuickBtn(string text, Color color, Action onClickAction)
        {
            Button btn = new Button
            {
                Text = text,
                Size = new Size(200, 100), // Büyük Kare Butonlar
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 20, 20) // Aralarına boşluk
            };
            btn.FlatAppearance.BorderSize = 0;
            // Butona basınca ilgili sayfayı açsın
            btn.Click += (s, e) => onClickAction.Invoke();
            return btn;
        }

        // --- YARDIMCI METOT (LoadDashboardView DIŞINA, Sınıf içine yapıştır) ---
        // Eğer bu zaten varsa eskisini silip bunu yapıştırın, renk ayarları güncellendi.


        // --- YARDIMCI: MİNİ TABLO OLUŞTURUCU ---
        // Bunu LoadDashboardView'in hemen altına (dışına) ekle
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

        // --- YARDIMCI METOT: İSTATİSTİK KARTI OLUŞTURUCU ---
        // Bu metot LoadDashboardView'in dışındadır ama Form1 sınıfının içindedir.
        private Panel CreateStatCard(string title, string value, Color color)
        {
            Panel card = new Panel
        {
                Width = 250,
                Height = 140,
                BackColor = Color.White,
                Margin = new Padding(0, 0, 20, 0)
            };

            // Renkli Sol Çizgi
            Panel accent = new Panel { Dock = DockStyle.Left, Width = 5, BackColor = color };
            card.Controls.Add(accent);

            // Sayı Değeri
            Label lblValue = new Label
            {
                Text = value,
                Font = new Font("Segoe UI", 24, FontStyle.Bold),
                ForeColor = Color.Black, // Hata almamak için Black yaptık
                Location = new Point(20, 25),
                AutoSize = true
            };
            card.Controls.Add(lblValue);

            // Başlık
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

        private void CreateStatCard(Panel parent, string title, string value, Color accentColor, int index)
        {
            Panel card = new Panel();
            card.Width = 280;
            card.Height = 130;
            card.BackColor = Color.White;
            card.Location = new Point(index * 300 + (index * 20), 0);
            card.Padding = new Padding(20);
            card.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            parent.Controls.Add(card);

            // Accent Bar
            Panel accentBar = new Panel();
            accentBar.Dock = DockStyle.Left;
            accentBar.Width = 5;
            accentBar.BackColor = accentColor;
            card.Controls.Add(accentBar);

            Label lblValue = new Label();
            lblValue.Text = value;
            lblValue.Font = new Font("Segoe UI", 28, FontStyle.Bold);
            lblValue.ForeColor = colorTextPrimary;
            lblValue.AutoSize = true;
            lblValue.Location = new Point(25, 20);
            card.Controls.Add(lblValue);

            Label lblTitle = new Label();
            lblTitle.Text = title;
            lblTitle.Font = new Font("Segoe UI", 11, FontStyle.Regular);
            lblTitle.ForeColor = colorTextSecondary;
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(25, 70);
            card.Controls.Add(lblTitle);
        }

        private void LoadPersonelListView()
        {
            // --- 1. ANA KAPLAYICI ---
            panelContent.Controls.Clear();
            Panel panelContainer = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10), BackColor = colorContent };
            panelContent.Controls.Add(panelContainer);

            Label lblHeader = new Label
            {
                Text = "👥 Personel Yönetimi",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                Dock = DockStyle.Top,
                Height = 50
            };
            panelContainer.Controls.Add(lblHeader);

            // --- 2. SOL PANEL (GİRİŞ FORMU - FlowLayout) ---
            FlowLayoutPanel pnlInput = new FlowLayoutPanel();
            pnlInput.Dock = DockStyle.Left;
            pnlInput.Width = 360;
            pnlInput.BackColor = Color.White;
            pnlInput.Padding = new Padding(20);
            pnlInput.FlowDirection = FlowDirection.TopDown;
            pnlInput.WrapContents = false;
            pnlInput.AutoScroll = true;
            panelContainer.Controls.Add(pnlInput);

            Label lblFormBaslik = new Label { Text = "Yeni Personel Ekle", Font = new Font("Segoe UI", 14, FontStyle.Bold), ForeColor = colorPrimary, AutoSize = true, Margin = new Padding(0, 0, 0, 20) };
            pnlInput.Controls.Add(lblFormBaslik);

            TextBox txtTc = AddInputControl(pnlInput, "TC Kimlik No:", 11);
            TextBox txtAd = AddInputControl(pnlInput, "Adı Soyadı:");
            TextBox txtIban = AddInputControl(pnlInput, "IBAN (TR):");
            TextBox txtGorev = AddInputControl(pnlInput, "Görev Yeri:");

            Label lblTarih = new Label { Text = "İşe Başlama Tarihi:", AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.Gray, Margin = new Padding(0, 10, 0, 5) };
            DateTimePicker dtpBaslama = new DateTimePicker { Width = 300, Height = 35, Format = DateTimePickerFormat.Short, Font = new Font("Segoe UI", 10), Margin = new Padding(0, 0, 0, 20) };
            pnlInput.Controls.Add(lblTarih);
            pnlInput.Controls.Add(dtpBaslama);

            Button btnKaydet = new Button { Text = "💾 Kaydet", Width = 300, Height = 50, BackColor = colorPrimary, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 11, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(0, 10, 0, 0) };
            btnKaydet.FlatAppearance.BorderSize = 0;
            pnlInput.Controls.Add(btnKaydet);

            // --- 3. SAĞ PANEL (LİSTE, EXCEL VE ARAMA) ---
            Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20, 0, 0, 0) };
            panelContainer.Controls.Add(pnlRight);
            pnlRight.BringToFront();

            // Üst Bar (Excel Butonu ve Arama Kutusu)
            Panel pnlRightTop = new Panel { Dock = DockStyle.Top, Height = 60 };
            pnlRight.Controls.Add(pnlRightTop);

            // Excel Butonu (Sola yaslı)
            Button btnExcelImport = CreateModernButton("📥 Excel'den Yükle", colorSuccess, 0, pnlRightTop);
            btnExcelImport.Width = 200;
            btnExcelImport.Location = new Point(0, 5);

            // ARAMA KUTUSU (Sağa yaslı veya butonun yanında)
            // 1. Etiket
            Label lblAra = new Label
            {
                Text = "🔍 Ara:",
                AutoSize = true,
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                Location = new Point(220, 15),
                ForeColor = Color.Gray
            };
            pnlRightTop.Controls.Add(lblAra);

            // 2. Textbox
            TextBox txtAra = new TextBox
            {
                Location = new Point(280, 12),
                Width = 250,
                Font = new Font("Segoe UI", 11)
            };
            pnlRightTop.Controls.Add(txtAra);

            // DataGridView
            DataGridView dgvPersonel = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AllowUserToAddRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvPersonel.ColumnHeadersDefaultCellStyle.BackColor = colorSidebar;
            dgvPersonel.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPersonel.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            dgvPersonel.ColumnHeadersHeight = 40;
            dgvPersonel.EnableHeadersVisualStyles = false;

            pnlRight.Controls.Add(dgvPersonel);
            pnlRightTop.SendToBack();

            // --- FONKSİYONLAR ---

            void PersonelListele()
            {
                try
                {
                    using (var conn = new MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        string sql = "SELECT pk_tc AS 'TC', pk_ad_soyad AS 'Ad Soyad', pk_iban_no AS 'IBAN', pk_gorev_yeri AS 'Görev', pk_is_baslama_tarihi AS 'Başlama' FROM program_katilimcilari";
                        using (var da = new MySql.Data.MySqlClient.MySqlDataAdapter(sql, conn))
                        {
                            DataTable dt = new DataTable(); da.Fill(dt); dgvPersonel.DataSource = dt;
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            }
            PersonelListele();

            // *** ARAMA MANTIĞI (FİLTRELEME) ***
            txtAra.TextChanged += (s, e) =>
            {
                DataTable dt = dgvPersonel.DataSource as DataTable;
                if (dt != null)
                {
                    string aranan = txtAra.Text.Trim().Replace("'", "''"); // Tırnak işareti hatasını önle

                    if (string.IsNullOrEmpty(aranan))
                    {
                        dt.DefaultView.RowFilter = ""; // Boşsa filtreyi kaldır
                    }
                    else
                    {
                        // SQL Sorgusu gibi çalışır ama veritabanına gitmez, RAM'de filtreler. Hızlıdır.
                        // Hem Ad Soyad hem de TC içinde arama yapar.
                        dt.DefaultView.RowFilter = string.Format("[Ad Soyad] LIKE '%{0}%' OR [TC] LIKE '%{0}%'", aranan);
                    }
                }
            };

            // KAYDET
            btnKaydet.Click += (s, e) =>
            {
                if (txtTc.Text == "" || txtAd.Text == "") { MessageBox.Show("TC ve Ad zorunlu."); return; }
                try
                {
                    using (var conn = new MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        string sqlPer = @"INSERT INTO program_katilimcilari (pk_tc, pk_ad_soyad, pk_iban_no, pk_gorev_yeri, pk_is_baslama_tarihi) VALUES (@tc, @ad, @iban, @gorev, @tarih)";
                        var cmd = new MySql.Data.MySqlClient.MySqlCommand(sqlPer, conn);
                        cmd.Parameters.AddWithValue("@tc", txtTc.Text); cmd.Parameters.AddWithValue("@ad", txtAd.Text);
                        cmd.Parameters.AddWithValue("@iban", txtIban.Text); cmd.Parameters.AddWithValue("@gorev", txtGorev.Text); cmd.Parameters.AddWithValue("@tarih", dtpBaslama.Value);
                        cmd.ExecuteNonQuery();

                        string sqlPua = @"INSERT IGNORE INTO puantaj (p_tc, p_ad_soyad, p_iban, p_ise_baslama_tarihi) VALUES (@tc, @ad, @iban, @tarih)";
                        var cmdPua = new MySql.Data.MySqlClient.MySqlCommand(sqlPua, conn);
                        cmdPua.Parameters.AddWithValue("@tc", txtTc.Text); cmdPua.Parameters.AddWithValue("@ad", txtAd.Text); cmdPua.Parameters.AddWithValue("@iban", txtIban.Text); cmdPua.Parameters.AddWithValue("@tarih", dtpBaslama.Value);
                        cmdPua.ExecuteNonQuery();
                    }
                    MessageBox.Show("✅ Personel eklendi.");
                    PersonelListele();
                    txtTc.Clear(); txtAd.Clear(); txtIban.Clear(); txtGorev.Clear();
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            };

            // EXCEL IMPORT
            btnExcelImport.Click += (s, e) =>
            {
                OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel|*.xlsx", Title = "Personel Listesi" };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var workbook = new ClosedXML.Excel.XLWorkbook(ofd.FileName))
                        {
                            var ws = workbook.Worksheet(1);
                            var rows = ws.RangeUsed().RowsUsed().Skip(1);
                            using (var conn = new MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                            {
                                conn.Open();
                                int sayac = 0;
                                foreach (var row in rows)
                                {
                                    string tc = row.Cell(1).GetValue<string>(); string ad = row.Cell(2).GetValue<string>();
                                    string iban = row.Cell(3).GetValue<string>(); string gorev = row.Cell(4).GetValue<string>();
                                    DateTime tarih = DateTime.Now; try { tarih = row.Cell(5).GetDateTime(); } catch { }

                                    string sqlPer = "INSERT IGNORE INTO program_katilimcilari (pk_tc, pk_ad_soyad, pk_iban_no, pk_gorev_yeri, pk_is_baslama_tarihi) VALUES (@tc, @ad, @iban, @gorev, @tarih)";
                                    var cmdPer = new MySql.Data.MySqlClient.MySqlCommand(sqlPer, conn);
                                    cmdPer.Parameters.AddWithValue("@tc", tc); cmdPer.Parameters.AddWithValue("@ad", ad);
                                    cmdPer.Parameters.AddWithValue("@iban", iban); cmdPer.Parameters.AddWithValue("@gorev", gorev); cmdPer.Parameters.AddWithValue("@tarih", tarih);
                                    cmdPer.ExecuteNonQuery();

                                    string sqlPua = "INSERT IGNORE INTO puantaj (p_tc, p_ad_soyad, p_iban, p_ise_baslama_tarihi) VALUES (@tc, @ad, @iban, @tarih)";
                                    var cmdPua = new MySql.Data.MySqlClient.MySqlCommand(sqlPua, conn);
                                    cmdPua.Parameters.AddWithValue("@tc", tc); cmdPua.Parameters.AddWithValue("@ad", ad); cmdPua.Parameters.AddWithValue("@iban", iban); cmdPua.Parameters.AddWithValue("@tarih", tarih);
                                    cmdPua.ExecuteNonQuery();
                                    sayac++;
                                }
                                MessageBox.Show($"✅ {sayac} personel eklendi.");
                            }
                        }
                        PersonelListele();
                    }
                    catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
                }
            };
        }
        private void LoadRaporlarView()
        {
            // --- TEMİZLİK VE ANA DÜZEN ---
            panelContent.Controls.Clear();
            TableLayoutPanel tlpMain = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = colorContent, Padding = new Padding(10) };
            tlpMain.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlpMain.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            panelContent.Controls.Add(tlpMain);

            // --- ÜST PANEL ---
            Panel pnlTopContainer = new Panel { AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10) };

            Label lblHeader = new Label { Text = "📊 Bordro ve Muhtasar İşlemleri", Font = new Font("Segoe UI", 16, FontStyle.Bold), ForeColor = colorTextPrimary, Dock = DockStyle.Top, Height = 45 };
            pnlTopContainer.Controls.Add(lblHeader);

            FlowLayoutPanel flowTools = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = true, Padding = new Padding(0, 10, 0, 0) };

            // Ücret Kutusu
            Panel pnlUcret = new Panel { Width = 160, Height = 60, Margin = new Padding(0, 0, 10, 0) };
            Label lblUcret = new Label { Text = "Günlük Brüt (TL):", Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.Gray };
            NumericUpDown numUcret = new NumericUpDown { Location = new Point(0, 25), Width = 150, Height = 35, Maximum = 10000, DecimalPlaces = 2, Value = 666.75M, Font = new Font("Segoe UI", 11) };
            pnlUcret.Controls.Add(lblUcret); pnlUcret.Controls.Add(numUcret);
            flowTools.Controls.Add(pnlUcret);

            // Butonlar (Yardımcı metodlar aşağıda tanımlanmalı)
            Button btnHesapla = CreateActionButton("⚙️ 1. Hesapla", Color.Orange);
            Button btnMuhtasar = CreateActionButton("📄 2. Muhtasar İndir", colorSuccess);
            Button btnBordro = CreateActionButton("📑 3. Bordro İndir", colorInfo);

            flowTools.Controls.Add(btnHesapla);
            flowTools.Controls.Add(btnMuhtasar);
            flowTools.Controls.Add(btnBordro);

            pnlTopContainer.Controls.Add(flowTools);
            flowTools.BringToFront(); lblHeader.BringToFront();
            tlpMain.Controls.Add(pnlTopContainer, 0, 0);

            // --- ALT TABLO ---
            DataGridView dgvOnizleme = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, BorderStyle = BorderStyle.FixedSingle, RowHeadersVisible = false, ReadOnly = true, AllowUserToAddRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
            dgvOnizleme.ColumnHeadersDefaultCellStyle.BackColor = colorSidebar; dgvOnizleme.ColumnHeadersDefaultCellStyle.ForeColor = Color.White; dgvOnizleme.EnableHeadersVisualStyles = false;
            tlpMain.Controls.Add(dgvOnizleme, 0, 1);

            // --- LİSTELEME FONKSİYONU ---
            void ListeyiGuncelle()
            {
                try
                {
                    using (var conn = new MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        string sql = "SELECT b_tc AS 'TC', b_ad_soyad AS 'Ad Soyad', b_aylik_calisilan_gun AS 'Gün', b_tahakkuk_toplami AS 'Brüt', b_odenmesi_gereken_net_tutar AS 'NET MAAŞ' FROM bordro";
                        using (var da = new MySql.Data.MySqlClient.MySqlDataAdapter(sql, conn)) { DataTable dt = new DataTable(); da.Fill(dt); dgvOnizleme.DataSource = dt; }
                    }
                }
                catch { }
            }
            ListeyiGuncelle();

            // --- BUTON OLAYLARI ---
            btnHesapla.Click += (s, e) =>
            {
                try
                {
                    decimal gunlukBruk = numUcret.Value;
                    using (var conn = new MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        new MySql.Data.MySqlClient.MySqlCommand("TRUNCATE TABLE bordro", conn).ExecuteNonQuery();
                        new MySql.Data.MySqlClient.MySqlCommand("TRUNCATE TABLE muhtasar_raporu", conn).ExecuteNonQuery();
                        new MySql.Data.MySqlClient.MySqlCommand("TRUNCATE TABLE banka_listesi", conn).ExecuteNonQuery();

                        string sqlPuantaj = "SELECT * FROM puantaj WHERE p_calistigi_gun_sayisi > 0";
                        var cmdGet = new MySql.Data.MySqlClient.MySqlCommand(sqlPuantaj, conn);
                        var dr = cmdGet.ExecuteReader();
                        DataTable dtPuantaj = new DataTable(); dtPuantaj.Load(dr);

                        foreach (DataRow row in dtPuantaj.Rows)
                        {
                            string tc = row["p_tc"].ToString(); string ad = row["p_ad_soyad"].ToString();
                            string iban = row["p_iban"].ToString(); int gun = Convert.ToInt32(row["p_calistigi_gun_sayisi"]);

                            decimal brutUcret = gun * gunlukBruk;
                            decimal sgkPrimi = brutUcret * 0.14M;
                            decimal damgaVergisi = brutUcret * 0.00759M;
                            decimal gelirVergisiMatrahi = brutUcret - sgkPrimi;
                            decimal gelirVergisi = gelirVergisiMatrahi * 0.15M;
                            decimal netUcret = brutUcret - (sgkPrimi + damgaVergisi + gelirVergisi);

                            // Ekleme Sorguları
                            string sqlBordro = "INSERT INTO bordro (b_tc, b_ad_soyad, b_gorev_yeri, b_aylik_calisilan_gun, b_tahakkuk_toplami, b_sosyal_guvenlik_primi, b_gelir_vergisi_kesintisi, b_damga_vergisi_kesintisi, b_odenmesi_gereken_net_tutar) VALUES (@tc, @ad, 'Merkez', @gun, @brut, @sgk, @gv, @dv, @net)";
                            using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sqlBordro, conn))
                            {
                                cmd.Parameters.AddWithValue("@tc", tc); cmd.Parameters.AddWithValue("@ad", ad); cmd.Parameters.AddWithValue("@gun", gun); cmd.Parameters.AddWithValue("@brut", brutUcret); cmd.Parameters.AddWithValue("@sgk", sgkPrimi); cmd.Parameters.AddWithValue("@gv", gelirVergisi); cmd.Parameters.AddWithValue("@dv", damgaVergisi); cmd.Parameters.AddWithValue("@net", netUcret); cmd.ExecuteNonQuery();
                            }
                            string sqlMuhtasar = "INSERT INTO muhtasar_raporu (mh_tc, mh_ad_soyad, mh_prim_odeme_gunu, mh_hak_edilen_ucret, mh_doneme_ait_gelir_vergisi_matrahi, mh_gelir_vergisi_kesintisi, mh_damga_vergisi_kesintisi) VALUES (@tc, @ad, @gun, @brut, @matrah, @gv, @dv)";
                            using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sqlMuhtasar, conn))
                            {
                                cmd.Parameters.AddWithValue("@tc", tc); cmd.Parameters.AddWithValue("@ad", ad); cmd.Parameters.AddWithValue("@gun", gun); cmd.Parameters.AddWithValue("@brut", brutUcret); cmd.Parameters.AddWithValue("@matrah", gelirVergisiMatrahi); cmd.Parameters.AddWithValue("@gv", gelirVergisi); cmd.Parameters.AddWithValue("@dv", damgaVergisi); cmd.ExecuteNonQuery();
                            }
                            string sqlBanka = "INSERT INTO banka_listesi (bl_tc, bl_ad_soyad, bl_iban_no, bl_tutar) VALUES (@tc, @ad, @iban, @net)";
                            using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sqlBanka, conn))
                            {
                                cmd.Parameters.AddWithValue("@tc", tc); cmd.Parameters.AddWithValue("@ad", ad); cmd.Parameters.AddWithValue("@iban", iban); cmd.Parameters.AddWithValue("@net", netUcret); cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    MessageBox.Show("✅ Hesaplama tamamlandı!");
                    ListeyiGuncelle();
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            };

            btnMuhtasar.Click += (s, e) => ExportTableToExcel("muhtasar_raporu", "Muhtasar_Raporu");
            btnBordro.Click += (s, e) => ExportTableToExcel("bordro", "Personel_Bordrosu");
        }


        // --- YARDIMCI BUTON OLUŞTURUCU (Bu metodu sınıf içine ekle) ---
        private Button CreateActionButton(string text, Color color)
        {
            return new Button
            {
                Text = text,
                Size = new Size(200, 55), // Geniş ve yüksek butonlar
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 15, 10) // Sağdan ve alttan boşluk bırak (Çakışmayı önler)
            };
        }

        // Ortak Excel Çıktı Fonksiyonu (Kod tekrarını önlemek için)
        private void ExportTableToExcel(string tableName, string fileName)
        {
            try
            {
                DataTable dt = new DataTable();
                using (var conn = new MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                {
                    conn.Open();
                    using (var da = new MySql.Data.MySqlClient.MySqlDataAdapter($"SELECT * FROM {tableName}", conn))
                    {
                        da.Fill(dt);
                    }
                }

                if (dt.Rows.Count == 0) { MessageBox.Show("Tabloda veri yok. Önce 'Hesapla' butonuna basın."); return; }

                using (var workbook = new ClosedXML.Excel.XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Rapor");
                    ws.Cell(1, 1).InsertTable(dt);
                    ws.Columns().AdjustToContents();

                    SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel Dosyası|*.xlsx", FileName = $"{fileName}_{DateTime.Now:yyyy-MM}.xlsx" };
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        workbook.SaveAs(sfd.FileName);
                        MessageBox.Show("✅ Dosya kaydedildi.");
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Excel Hatası: " + ex.Message); }
        }

        // --- YENİ YARDIMCI METOT (Bu metod Form1 class'ının içinde herhangi bir yere ekleyin) ---
        // Bu metod, giriş kutularını FlowLayoutPanel içine düzgünce ekler.
        private TextBox AddInputControl(FlowLayoutPanel parent, string labelText, int maxLength = 100)
        {
            // Etiket
            Label lbl = new Label
            {
                Text = labelText,
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.Gray,
                Margin = new Padding(0, 5, 0, 5) // Üstten ve alttan boşluk
            };
            parent.Controls.Add(lbl);

            // Kutu
            TextBox txt = new TextBox
            {
                Width = 300,
                Height = 35, // Yükseklik veriyoruz
                Font = new Font("Segoe UI", 11),
                MaxLength = maxLength,
                Margin = new Padding(0, 0, 0, 15) // Bir sonraki elemanla arasına boşluk koy
            };
            parent.Controls.Add(txt);

            return txt;
        }

        // Yardımcı Input Oluşturucu (Kodu kısaltmak için)
        private TextBox CreateInput(Panel parent, string labelText, ref int yPos, int maxLength = 100)
        {
            Label lbl = new Label { Text = labelText, Location = new Point(20, yPos), AutoSize = true, Font = new Font("Segoe UI", 10) };
            parent.Controls.Add(lbl);

            TextBox txt = new TextBox { Location = new Point(20, yPos + 25), Width = 300, Font = new Font("Segoe UI", 10), MaxLength = maxLength };
            parent.Controls.Add(txt);

            yPos += 65; // Bir sonraki eleman için aşağı kay
            return txt;
        }



        private void LoadExcelView()
        {
            // --- ARAYÜZ OLUŞTURMA KISMI ---
            Panel panelContainer = new Panel();
            panelContainer.Dock = DockStyle.Fill;
            panelContainer.BackColor = Color.Transparent;
            panelContainer.Padding = new Padding(30);
            panelContent.Controls.Add(panelContainer);

            // Başlık
            Label lblHeader = new Label();
            lblHeader.Text = "📄 İŞKUR Puantaj ve Banka Entegrasyonu";
            lblHeader.Font = new Font("Segoe UI", 16, FontStyle.Bold);
            lblHeader.ForeColor = colorTextPrimary;
            lblHeader.AutoSize = true;
            lblHeader.Dock = DockStyle.Top;
            panelContainer.Controls.Add(lblHeader);

            // Araçlar Paneli
            Panel pnlTools = new Panel();
            pnlTools.Dock = DockStyle.Top;
            pnlTools.Height = 120;
            pnlTools.BackColor = Color.Transparent;
            pnlTools.Padding = new Padding(0, 20, 0, 0);
            panelContainer.Controls.Add(pnlTools);
            pnlTools.BringToFront();

            // ---------------------------------------------------------
            // 1. BUTON: EXCEL'DEN PUANTAJ YÜKLE (puantaj tablosuna)
            // ---------------------------------------------------------
            Button btnImport = CreateModernButton("📥 1. Puantaj Yükle", colorSuccess, 0, pnlTools);
            btnImport.Click += (s, e) =>
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel Dosyaları|*.xlsx;*.xls";
                ofd.Title = "Puantaj Dosyasını Seçin";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook(ofd.FileName))
                        {
                            var worksheet = workbook.Worksheet(1); // İlk sayfa
                            var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Başlığı atla

                            // DİKKAT: Veritabanı adını 'iskur' yaptık
                            using (var conn = new MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                            {
                                conn.Open();

                                // Temiz kurulum için önce eski puantajı siliyoruz (Çakışma olmasın diye)
                                new MySql.Data.MySqlClient.MySqlCommand("TRUNCATE TABLE puantaj", conn).ExecuteNonQuery();

                                foreach (var row in rows)
                                {
                                    // Excel'deki sütun sırası: 1:TC, 2:Ad Soyad, 3:IBAN, 4:Gün Sayısı
                                    string tc = row.Cell(1).GetValue<string>();
                                    string adSoyad = row.Cell(2).GetValue<string>();
                                    string iban = row.Cell(3).GetValue<string>();
                                    int gunSayisi = 0;
                                    int.TryParse(row.Cell(4).GetValue<string>(), out gunSayisi);

                                    // SQL İsimleri 'iskur.sql' dosyasına göre uyarlandı:
                                    // p_tc, p_ad_soyad, p_iban, p_calistigi_gun_sayisi
                                    string query = @"INSERT INTO puantaj 
                                           (p_tc, p_ad_soyad, p_iban, p_calistigi_gun_sayisi, p_ise_baslama_tarihi) 
                                           VALUES (@tc, @ad, @iban, @gun, CURDATE())";

                                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(query, conn))
                                    {
                                        cmd.Parameters.AddWithValue("@tc", tc);
                                        cmd.Parameters.AddWithValue("@ad", adSoyad);
                                        cmd.Parameters.AddWithValue("@iban", iban);
                                        cmd.Parameters.AddWithValue("@gun", gunSayisi);
                                        cmd.ExecuteNonQuery();
                }
                                }
                            }
                        }
                        MessageBox.Show("✅ Puantaj listesi başarıyla yüklendi!", "İşlem Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Hata: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            };
            pnlTools.Controls.Add(btnImport);

            // ---------------------------------------------------------
            // 2. BUTON: MAAŞ HESAPLA (puantaj -> banka_listesi tablosuna)
            // ---------------------------------------------------------
            Button btnCalculate = CreateModernButton("💰 2. Maaş Hesapla", Color.Orange, 1, pnlTools);
            btnCalculate.Click += (s, e) =>
            {
                try
                {
                    float gunlukUcret = 500.0f; // Burayı istersen bir kutucuktan alabilirsin

                    using (var conn = new MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();

                        // Banka listesini temizle
                        new MySql.Data.MySqlClient.MySqlCommand("TRUNCATE TABLE banka_listesi", conn).ExecuteNonQuery();

                        // Puantaj tablosundan verileri alıp hesaplayarak banka listesine atıyoruz.
                        // Sütun isimleri: bl_tc, bl_ad_soyad, bl_iban_no, bl_tutar
                        string sql = @"INSERT INTO banka_listesi (bl_tc, bl_ad_soyad, bl_iban_no, bl_tutar)
                               SELECT p_tc, p_ad_soyad, p_iban, (p_calistigi_gun_sayisi * @ucret) 
                               FROM puantaj 
                               WHERE p_calistigi_gun_sayisi > 0";

                        using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@ucret", gunlukUcret);
                            int sayi = cmd.ExecuteNonQuery();
                            MessageBox.Show($"✅ {sayi} kişinin maaşı hesaplandı ve banka listesine yazıldı.", "Tamamlandı");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hesaplama Hatası: " + ex.Message);
                }
            };
            pnlTools.Controls.Add(btnCalculate);

            // ---------------------------------------------------------
            // 3. BUTON: BANKA LİSTESİ İNDİR (banka_listesi -> Excel)
            // ---------------------------------------------------------
            Button btnExport = CreateModernButton("📤 3. Banka Listesi İndir", colorInfo, 2, pnlTools);
            btnExport.Click += (s, e) =>
            {
                try
                {
                    DataTable dt = new DataTable();
                    using (var conn = new MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        // Veritabanındaki 'bl_' sütunlarını çekiyoruz
                        string sql = "SELECT bl_ad_soyad AS 'Ad Soyad', bl_iban_no AS 'IBAN', bl_tutar AS 'Tutar' FROM banka_listesi";
                        using (var da = new MySql.Data.MySqlClient.MySqlDataAdapter(sql, conn))
                        {
                            da.Fill(dt);
                        }
                    }

                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Excel Dosyası|*.xlsx";
                    sfd.FileName = $"Banka_Odeme_Listesi_{DateTime.Now:yyyy-MM-dd}.xlsx";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Banka Listesi");
                            worksheet.Cell(1, 1).InsertTable(dt);
                            worksheet.Columns().AdjustToContents();
                            workbook.SaveAs(sfd.FileName);
                        }
                        MessageBox.Show("✅ Banka listesi Excel dosyası olarak kaydedildi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Dışa Aktarma Hatası: " + ex.Message);
                }
            };
            pnlTools.Controls.Add(btnExport);
            pnlTools.SendToBack();

            // Alt Bilgi
            Label lblInfo = new Label();
            lblInfo.Text = "ℹ️ Sistem 'iskur' veritabanına bağlıdır. Excel dosyanızda sırasıyla: TC, Ad Soyad, IBAN ve Gün Sayısı olmalıdır.";
            lblInfo.Font = new Font("Segoe UI", 10, FontStyle.Italic);
            lblInfo.ForeColor = colorTextSecondary;
            lblInfo.Dock = DockStyle.Bottom;
            lblInfo.Padding = new Padding(0, 20, 0, 0);
            lblInfo.Height = 100;
            panelContainer.Controls.Add(lblInfo);
        }

        private Button CreateModernButton(string text, Color backColor, int index, Panel parent)
        {
            Button btn = new Button();
            btn.Text = text;
            btn.Size = new Size(220, 50);

            btn.Location = new Point(index * 240, 15);
            btn.BackColor = backColor;
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;
            btn.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            Color hoverColor = Color.FromArgb(
                Math.Max(0, backColor.R - 20),
                Math.Max(0, backColor.G - 20),
                Math.Max(0, backColor.B - 20)
            );

            btn.MouseEnter += (s, e) =>
            {
                btn.BackColor = hoverColor;
            };
            btn.MouseLeave += (s, e) =>
            {
                btn.BackColor = backColor;
            };

            return btn;
        }

        // Sınıf seviyesinde şu değişkenin olduğundan emin ol:
        // private DateTime secilenTarih = DateTime.Now;

        private void LoadPuantajView()
        {
            // --- 1. SAYFA TEMİZLİĞİ ---
            panelContent.Controls.Clear();

            // ANA DÜZENLEYİCİ: Ekranı dikeyde 2 parçaya bölen tablo yapısı
            TableLayoutPanel tlpMain = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = colorContent,
                Padding = new Padding(10)
            };
            // 1. Satır: Otomatik Yükseklik (İçeriğe göre)
            tlpMain.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            // 2. Satır: Kalan her yeri kapla (%100)
            tlpMain.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            panelContent.Controls.Add(tlpMain);

            // --- 2. ÜST KISIM (BAŞLIK, TARİH, BUTONLAR) -> Satır 0 ---
            Panel pnlTopContainer = new Panel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 10) // Tablo ile arasına boşluk
            };

            // A) Başlık
            Label lblHeader = new Label
            {
                Text = "📝 Personel Puantaj Girişi",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                Dock = DockStyle.Top,
                Height = 40
            };
            pnlTopContainer.Controls.Add(lblHeader);

            // B) Araç Çubuğu (Tarih ve Butonlar)
            FlowLayoutPanel flowTools = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(0, 10, 0, 0)
            };

            // Tarih Seçici
            DateTimePicker dtpDonem = new DateTimePicker
            {
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "MMMM yyyy",
                Width = 200,
                Height = 40,
                Font = new Font("Segoe UI", 12),
                Margin = new Padding(0, 5, 20, 10)
            };
            dtpDonem.Value = secilenTarih;
            flowTools.Controls.Add(dtpDonem);

            // Butonlar
            Button btnExcelExport = new Button
            {
                Text = "📤 Excel Oluştur",
                Size = new Size(160, 45),
                BackColor = colorSuccess,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 10, 5)
            };
            btnExcelExport.FlatAppearance.BorderSize = 0;
            flowTools.Controls.Add(btnExcelExport);

            Button btnKaydet = new Button
            {
                Text = "💾 Kaydet",
                Size = new Size(140, 45),
                BackColor = colorPrimary,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 10, 5)
            };
            btnKaydet.FlatAppearance.BorderSize = 0;
            flowTools.Controls.Add(btnKaydet);

            // Araç çubuğunu üst panele ekle
            pnlTopContainer.Controls.Add(flowTools);

            // Sıralama (Başlık en üstte, Araçlar onun altında)
            flowTools.BringToFront();
            lblHeader.SendToBack(); // Dock.Top mantığında, en son eklenen veya SendToBack yapılan en üstte durur. 
                                    // Ama biz Panel kullandık. Dock=Top sırası: Kodda son eklenen en üste çıkar.
                                    // O yüzden lblHeader'ı en son ekleyelim veya BringToFront yapalım.
            lblHeader.BringToFront();

            // Bilgi Etiketi
            Label lblInfo = new Label
            {
                Text = "ℹ️ Bilgi: Hücrelere tıklayarak durumu değiştirin (X: Çalıştı, İ: İzinli, R: Raporlu). Haftada maksimum 3 gün çalışılabilir.",
                AutoSize = true,
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 10, FontStyle.Italic),
                Dock = DockStyle.Bottom,
                Padding = new Padding(5, 5, 0, 0)
            };
            pnlTopContainer.Controls.Add(lblInfo);

            // Üst Paneli Ana Tabloya Ekle
            tlpMain.Controls.Add(pnlTopContainer, 0, 0);

            // --- 3. TABLO (GRID) -> Satır 1 ---
            DataGridView dgvPuantaj = new DataGridView
            {
                Dock = DockStyle.Fill, // Bulunduğu hücreyi doldur
                BackgroundColor = Color.White,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                BorderStyle = BorderStyle.FixedSingle
            };
            dgvPuantaj.ColumnHeadersHeight = 40;
            dgvPuantaj.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvPuantaj.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvPuantaj.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);

            // Grid'i Ana Tabloya Ekle
            tlpMain.Controls.Add(dgvPuantaj, 0, 1);

            // --- FONKSİYONLAR ---

            void GridOlustur()
            {
                dgvPuantaj.Columns.Clear();
                dgvPuantaj.Rows.Clear();

                dgvPuantaj.Columns.Add("colTc", "TC Kimlik");
                dgvPuantaj.Columns.Add("colAd", "Ad Soyad");
                dgvPuantaj.Columns[0].ReadOnly = true;
                dgvPuantaj.Columns[1].ReadOnly = true;
                dgvPuantaj.Columns[0].Width = 100;
                dgvPuantaj.Columns[1].Width = 150;
                dgvPuantaj.Columns[1].Frozen = true;

                int gunSayisi = DateTime.DaysInMonth(secilenTarih.Year, secilenTarih.Month);
                for (int i = 1; i <= gunSayisi; i++)
                {
                    DateTime gunTarihi = new DateTime(secilenTarih.Year, secilenTarih.Month, i);
                    string baslik = i.ToString() + "\n" + gunTarihi.ToString("ddd", new System.Globalization.CultureInfo("tr-TR"));

                    dgvPuantaj.Columns.Add("day" + i, baslik);
                    dgvPuantaj.Columns[i + 1].Width = 45;

                    if (gunTarihi.DayOfWeek == DayOfWeek.Saturday || gunTarihi.DayOfWeek == DayOfWeek.Sunday)
                    {
                        dgvPuantaj.Columns[i + 1].DefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
                    }
                }
            }

            void VerileriYukle()
            {
                GridOlustur();
                try
                {
                    using (var conn = new MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        string sql = "SELECT p_tc, p_ad_soyad, p_gun_detaylari FROM puantaj";
                        var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn);
                        var dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {
                            int rowIndex = dgvPuantaj.Rows.Add();
                            dgvPuantaj.Rows[rowIndex].Cells[0].Value = dr["p_tc"].ToString();
                            dgvPuantaj.Rows[rowIndex].Cells[1].Value = dr["p_ad_soyad"].ToString();
                            dgvPuantaj.Rows[rowIndex].Tag = dr["p_tc"].ToString();

                            string detay = dr["p_gun_detaylari"].ToString();
                            if (!string.IsNullOrEmpty(detay))
                            {
                                string[] gunler = detay.Split('-');
                                for (int i = 0; i < gunler.Length && i < dgvPuantaj.Columns.Count - 2; i++)
                                {
                                    string val = gunler[i] == "0" ? "" : gunler[i];
                                    var cell = dgvPuantaj.Rows[rowIndex].Cells[i + 2];
                                    cell.Value = val;
                                    if (val == "X") cell.Style.BackColor = Color.LightGreen;
                                    else if (val == "İ") cell.Style.BackColor = Color.LightYellow;
                                    else if (val == "R") cell.Style.BackColor = Color.LightPink;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            }

            bool HaftalikLimitAsildiMi(int rowIndex, int gunSutunIndex)
            {
                int gun = gunSutunIndex - 1;
                DateTime tiklananTarih = new DateTime(secilenTarih.Year, secilenTarih.Month, gun);
                int fark = (int)tiklananTarih.DayOfWeek == 0 ? 6 : (int)tiklananTarih.DayOfWeek - 1;
                DateTime haftaBasi = tiklananTarih.AddDays(-fark);
                DateTime haftaSonu = haftaBasi.AddDays(6);

                int buHaftakiXSayisi = 0;
                int gunSayisi = DateTime.DaysInMonth(secilenTarih.Year, secilenTarih.Month);

                for (int i = 1; i <= gunSayisi; i++)
                {
                    DateTime currentDay = new DateTime(secilenTarih.Year, secilenTarih.Month, i);
                    if (currentDay >= haftaBasi && currentDay <= haftaSonu)
                    {
                        var val = dgvPuantaj.Rows[rowIndex].Cells[i + 1].Value;
                        if (val != null && val.ToString() == "X") buHaftakiXSayisi++;
                    }
                }

                var tiklananHucreDegeri = dgvPuantaj.Rows[rowIndex].Cells[gunSutunIndex].Value;
                bool suAnXDegil = tiklananHucreDegeri == null || tiklananHucreDegeri.ToString() != "X";

                return (buHaftakiXSayisi >= 3 && suAnXDegil);
            }

            // Olaylar
            VerileriYukle();

            dtpDonem.ValueChanged += (s, e) =>
            {
                secilenTarih = dtpDonem.Value;
                VerileriYukle();
            };

            dgvPuantaj.CellClick += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 2)
                {
                    var cell = dgvPuantaj.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    string val = cell.Value?.ToString() ?? "";

                    if (val == "")
                    {
                        if (HaftalikLimitAsildiMi(e.RowIndex, e.ColumnIndex))
                        {
                            MessageBox.Show("Bu hafta için maksimum 3 gün çalışma limiti doldu!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        cell.Value = "X"; cell.Style.BackColor = Color.LightGreen;
                    }
                    else if (val == "X") { cell.Value = "İ"; cell.Style.BackColor = Color.LightYellow; }
                    else if (val == "İ") { cell.Value = "R"; cell.Style.BackColor = Color.LightPink; }
                    else { cell.Value = ""; cell.Style.BackColor = Color.White; }
                }
            };

            btnKaydet.Click += (s, e) =>
            {
                try
                {
                    using (var conn = new MySql.Data.MySqlClient.MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        foreach (DataGridViewRow row in dgvPuantaj.Rows)
                        {
                            string tc = row.Cells[0].Value.ToString();
                            List<string> gunVerileri = new List<string>();
                            int toplamCalisilanGun = 0;

                            for (int i = 2; i < dgvPuantaj.Columns.Count; i++)
                            {
                                string v = row.Cells[i].Value?.ToString();
                                if (string.IsNullOrEmpty(v)) v = "0";
                                if (v == "X") toplamCalisilanGun++;
                                gunVerileri.Add(v);
                            }
                            string detayString = string.Join("-", gunVerileri);

                            string sql = "UPDATE puantaj SET p_gun_detaylari = @detay, p_calistigi_gun_sayisi = @toplam, p_yil_ay = @donem WHERE p_tc = @tc";
                            var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn);
                            cmd.Parameters.AddWithValue("@detay", detayString);
                            cmd.Parameters.AddWithValue("@toplam", toplamCalisilanGun);
                            cmd.Parameters.AddWithValue("@donem", secilenTarih.ToString("yyyy-MM"));
                            cmd.Parameters.AddWithValue("@tc", tc);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    MessageBox.Show("✅ Puantaj başarıyla kaydedildi!");
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            };

            btnExcelExport.Click += (s, e) =>
            {
                try
                {
                    using (var workbook = new ClosedXML.Excel.XLWorkbook())
                    {
                        var ws = workbook.Worksheets.Add("Puantaj");

                        ws.Cell(1, 1).Value = "BURSA ULUDAĞ ÜNİVERSİTESİ";
                        ws.Range(1, 1, 1, dgvPuantaj.Columns.Count).Merge().Style.Font.Bold = true;
                        ws.Cell(1, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

                        ws.Cell(2, 1).Value = dtpDonem.Value.ToString("MMMM yyyy").ToUpper() + " PUANTAJ CETVELİ";
                        ws.Range(2, 1, 2, dgvPuantaj.Columns.Count).Merge().Style.Font.Bold = true;
                        ws.Cell(2, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

                        int colIndex = 1;
                        foreach (DataGridViewColumn col in dgvPuantaj.Columns)
                        {
                            ws.Cell(4, colIndex).Value = col.HeaderText.Replace("\n", " ");
                            colIndex++;
                        }

                        for (int i = 0; i < dgvPuantaj.Rows.Count; i++)
                        {
                            for (int j = 0; j < dgvPuantaj.Columns.Count; j++)
                            {
                                var val = dgvPuantaj.Rows[i].Cells[j].Value?.ToString();
                                ws.Cell(i + 5, j + 1).Value = val;

                                if (val == "X") ws.Cell(i + 5, j + 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightGreen;
                                if (val == "İ") ws.Cell(i + 5, j + 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightYellow;
                                if (val == "R") ws.Cell(i + 5, j + 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightPink;
                            }
                        }
                        ws.Columns().AdjustToContents();

                        SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel Dosyası|*.xlsx", FileName = "Puantaj_Cizelgesi.xlsx" };
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            workbook.SaveAs(sfd.FileName);
                            MessageBox.Show("Excel oluşturuldu!");
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show("Excel Hatası: " + ex.Message); }
            };
        }

        private void LoadAyarlarView()
        {
            Label lblInfo = new Label();
            lblInfo.Text = "⚙️ Ayarlar\n\nSistem ayarları bu bölümde yapılandırılacaktır.";
            lblInfo.Font = new Font("Segoe UI", 14, FontStyle.Regular);
            lblInfo.ForeColor = colorTextPrimary;
            lblInfo.AutoSize = true;
            lblInfo.Location = new Point(30, 30);
            panelContent.Controls.Add(lblInfo);
        }

    }
}
