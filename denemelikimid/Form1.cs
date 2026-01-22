using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using denemelikimid.DataBase;
using MySql.Data.MySqlClient;

namespace denemelikimid
{
    public partial class Form1 : Form
    {
        private DateTime secilenTarih = DateTime.Now;

        // Renkler
        private Color colorPrimary = Color.FromArgb(67, 97, 238);
        private Color colorPrimaryDark = Color.FromArgb(52, 76, 186);
        private Color colorSecondary = Color.FromArgb(255, 255, 255);
        private Color colorSidebar = Color.FromArgb(30, 41, 59);
        private Color colorSidebarHover = Color.FromArgb(51, 65, 85);
        private Color colorHeader = Color.FromArgb(248, 250, 252);
        private Color colorContent = Color.FromArgb(241, 245, 249);
        private Color colorSuccess = Color.FromArgb(34, 197, 94);
        private Color colorInfo = Color.FromArgb(59, 130, 246);
        private Color colorDanger = Color.FromArgb(239, 68, 68);
        private Color colorTextPrimary = Color.FromArgb(15, 23, 42);
        private Color colorTextSecondary = Color.FromArgb(100, 116, 139);

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

            // Varsayılan görünüm
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
                lblTitle.Text = text
                    .Replace("📊 ", "")
                    .Replace("👥 ", "")
                    .Replace("📝 ", "")
                    .Replace("📄 ", "")
                    .Replace("⚙️ ", "");

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

        private Button CreateModernButton(string text, Color backColor, int index, Panel parent)
        {
            Button btn = new Button
            {
                Text = text,
                Size = new Size(220, 50),
                Location = new Point(index * 240, 15),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };

            btn.FlatAppearance.BorderSize = 0;

            Color hoverColor = Color.FromArgb(
                System.Math.Max(0, backColor.R - 20),
                System.Math.Max(0, backColor.G - 20),
                System.Math.Max(0, backColor.B - 20)
            );

            btn.MouseEnter += (s, e) => { btn.BackColor = hoverColor; };
            btn.MouseLeave += (s, e) => { btn.BackColor = backColor; };

            parent.Controls.Add(btn);
            return btn;
        }
    }
}



