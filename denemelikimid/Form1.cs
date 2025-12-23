using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace denemelikimid
{
    public partial class Form1 : Form
    {
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
                case "Excel":
                    LoadExcelView();
                    break;
                case "Ayarlar":
                    LoadAyarlarView();
                    break;
            }
        }

        private void LoadDashboardView()
        {
            // İçeriği temizle
            panelContent.Controls.Clear();

            // 1. ÜST KISIM: İstatistik Kartları için
            TableLayoutPanel tlpStats = new TableLayoutPanel();
            tlpStats.Dock = DockStyle.Top;
            tlpStats.Height = 160;
            tlpStats.ColumnCount = 4;
            tlpStats.Padding = new Padding(0, 0, 0, 20);


            tlpStats.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            tlpStats.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            tlpStats.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            tlpStats.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));

            panelContent.Controls.Add(tlpStats);


            tlpStats.Controls.Add(CreateStatCard("👥 Toplam Personel", "247", colorPrimary));
            tlpStats.Controls.Add(CreateStatCard("✅ Aktif Personel", "198", colorSuccess));
            tlpStats.Controls.Add(CreateStatCard("📊 Bu Ay Puantaj", "1,234", colorInfo));
            tlpStats.Controls.Add(CreateStatCard("📄 Bekleyen İşlem", "12", Color.FromArgb(251, 191, 36)));

            // 2. ALT KISIM: Hoş Geldiniz Mesajı
            Panel panelWelcome = new Panel();
            panelWelcome.Dock = DockStyle.Fill;
            panelWelcome.BackColor = Color.White;
            panelWelcome.Padding = new Padding(30);
            panelContent.Controls.Add(panelWelcome);
            panelWelcome.BringToFront();

            Label lblWelcome = new Label();
            lblWelcome.Text = "🎉 Sisteme Hoş Geldiniz!\n\n" +
                              "Sol menüden işlemlerinize devam edebilirsiniz.";
            lblWelcome.Font = new Font("Segoe UI", 14, FontStyle.Regular);
            lblWelcome.ForeColor = colorTextPrimary;
            lblWelcome.AutoSize = true;
            lblWelcome.Location = new Point(30, 30);
            panelWelcome.Controls.Add(lblWelcome);
        }


        private Panel CreateStatCard(string title, string value, Color accentColor)
        {
            Panel card = new Panel();
            card.Dock = DockStyle.Fill;
            card.Margin = new Padding(5);
            card.BackColor = Color.White;

            // Sol taraftaki renkli şerit
            Panel accentBar = new Panel();
            accentBar.Dock = DockStyle.Left;
            accentBar.Width = 6;
            accentBar.BackColor = accentColor;
            card.Controls.Add(accentBar);

            Label lblValue = new Label();
            lblValue.Text = value;
            lblValue.Font = new Font("Segoe UI", 24, FontStyle.Bold);
            lblValue.ForeColor = colorTextPrimary;
            lblValue.AutoSize = true;
            lblValue.Location = new Point(20, 25);
            card.Controls.Add(lblValue);

            Label lblTitle = new Label();
            lblTitle.Text = title;
            lblTitle.Font = new Font("Segoe UI", 10, FontStyle.Regular);
            lblTitle.ForeColor = colorTextSecondary;
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(20, 75);
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
            Label lblInfo = new Label();
            lblInfo.Text = "👥 Personel Listesi\n\nBu bölümde tüm personel bilgileri görüntülenecektir.";
            lblInfo.Font = new Font("Segoe UI", 14, FontStyle.Regular);
            lblInfo.ForeColor = colorTextPrimary;
            lblInfo.AutoSize = true;
            lblInfo.Location = new Point(30, 30);
            panelContent.Controls.Add(lblInfo);
        }



        private void LoadExcelView()
        {

            Panel panelContainer = new Panel();
            panelContainer.Dock = DockStyle.Fill;
            panelContainer.BackColor = Color.Transparent;

            panelContainer.Padding = new Padding(0);
            panelContent.Controls.Add(panelContainer);

            // Araçlar Paneli
            Panel pnlTools = new Panel();
            pnlTools.Dock = DockStyle.Top;
            pnlTools.Height = 80;
            pnlTools.BackColor = Color.Transparent;
            pnlTools.Padding = new Padding(0, 0, 0, 20);
            panelContainer.Controls.Add(pnlTools);
            pnlTools.SendToBack();

            // İçe Aktar Butonu
            Button btnImport = CreateModernButton("📥 Excel'den İçe Aktar", colorSuccess, 0, pnlTools);
            btnImport.Click += (s, e) =>
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel Dosyaları|*.xlsx;*.xls";
                ofd.Title = "Excel Dosyası Seçin";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    MessageBox.Show($"Dosya seçildi: {ofd.FileName}\n\n(İçe aktarma işlevi eklenecek)",
                        "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            pnlTools.Controls.Add(btnImport);

            // Dışa Aktar Butonu
            Button btnExport = CreateModernButton("📤 Excel'e Dışa Aktar", colorInfo, 1, pnlTools);


            btnExport.Click += (s, e) =>
            {



            };
            pnlTools.Controls.Add(btnExport);
            pnlTools.SendToBack();

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
