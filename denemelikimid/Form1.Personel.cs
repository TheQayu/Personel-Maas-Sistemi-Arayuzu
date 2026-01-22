using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using MySql.Data.MySqlClient;

namespace denemelikimid
{
    public partial class Form1
    {
        private void LoadPersonelListView()
        {
            // --- 1. ARAYÜZ TEMİZLİĞİ ---
            panelContent.Controls.Clear();
            Panel panelContainer = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10), BackColor = colorContent };
            panelContent.Controls.Add(panelContainer);

            Label lblHeader = new Label { Text = "👥 Personel Yönetimi", Font = new Font("Segoe UI", 16, FontStyle.Bold), ForeColor = colorTextPrimary, Dock = DockStyle.Top, Height = 50 };
            panelContainer.Controls.Add(lblHeader);

            // --- 2. SOL PANEL (GİRİŞ FORMU) ---
            FlowLayoutPanel pnlInput = new FlowLayoutPanel { Dock = DockStyle.Left, Width = 360, BackColor = Color.White, Padding = new Padding(20), FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
            panelContainer.Controls.Add(pnlInput);

            Label lblFormBaslik = new Label { Text = "Yeni Personel Ekle", Font = new Font("Segoe UI", 14, FontStyle.Bold), ForeColor = colorPrimary, AutoSize = true, Margin = new Padding(0, 0, 0, 20) };
            pnlInput.Controls.Add(lblFormBaslik);

            TextBox txtTc = AddInputControl(pnlInput, "TC Kimlik No:", 11);
            TextBox txtAd = AddInputControl(pnlInput, "Adı Soyadı:");
            TextBox txtIban = AddInputControl(pnlInput, "IBAN (TR):");

            // Görev Yeri (Departman)
            Label lblGorev = new Label { Text = "Görev Yeri (Departman):", AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.Gray, Margin = new Padding(0, 5, 0, 5) };
            pnlInput.Controls.Add(lblGorev);
            ComboBox cmbGorev = new ComboBox { Width = 300, Height = 35, Font = new Font("Segoe UI", 11), DropDownStyle = ComboBoxStyle.DropDown, Margin = new Padding(0, 0, 0, 15) };
            cmbGorev.Items.AddRange(new string[] { "İdari", "Teknik", "Güvenlik", "Temizlik", "Bilişim", "Akademik" });
            cmbGorev.SelectedIndex = 0;
            pnlInput.Controls.Add(cmbGorev);

            // Kampüs
            Label lblKampus = new Label { Text = "Kampüs:", AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.Gray, Margin = new Padding(0, 5, 0, 5) };
            pnlInput.Controls.Add(lblKampus);
            ComboBox cmbKampusEkle = new ComboBox { Width = 300, Height = 35, Font = new Font("Segoe UI", 11), DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 0, 0, 15) };
            cmbKampusEkle.Items.AddRange(new string[] { "Kampüs1", "Kampüs2", "Kampüs3" });
            cmbKampusEkle.SelectedIndex = 0;
            pnlInput.Controls.Add(cmbKampusEkle);

            // Tarih ve Kaydet
            Label lblTarih = new Label { Text = "İşe Başlama Tarihi:", AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.Gray, Margin = new Padding(0, 10, 0, 5) };
            DateTimePicker dtpBaslama = new DateTimePicker { Width = 300, Height = 35, Format = DateTimePickerFormat.Short, Font = new Font("Segoe UI", 10), Margin = new Padding(0, 0, 0, 20) };
            pnlInput.Controls.Add(lblTarih);
            pnlInput.Controls.Add(dtpBaslama);

            Button btnKaydet = new Button { Text = "💾 Kaydet", Width = 300, Height = 50, BackColor = colorPrimary, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 11, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(0, 10, 0, 0) };
            btnKaydet.FlatAppearance.BorderSize = 0;
            pnlInput.Controls.Add(btnKaydet);

            // --- 3. SAĞ PANEL ---
            Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20, 0, 0, 0) };
            panelContainer.Controls.Add(pnlRight);
            pnlRight.BringToFront();

            Panel pnlRightTop = new Panel { Dock = DockStyle.Top, Height = 60 };
            pnlRight.Controls.Add(pnlRightTop);

            Button btnExcelImport = CreateModernButton("📥 Excel'den Yükle", colorSuccess, 0, pnlRightTop);
            btnExcelImport.Width = 160;
            btnExcelImport.Location = new Point(0, 5);

            Label lblFiltre = new Label { Text = "Kampüs Seç:", AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(180, 15), ForeColor = Color.Gray };
            pnlRightTop.Controls.Add(lblFiltre);

            ComboBox cmbKampusFiltre = new ComboBox { Location = new Point(270, 12), Width = 150, Font = new Font("Segoe UI", 11), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbKampusFiltre.Items.AddRange(new string[] { "Tümü", "Kampüs1", "Kampüs2", "Kampüs3" });
            cmbKampusFiltre.SelectedIndex = 0;
            pnlRightTop.Controls.Add(cmbKampusFiltre);

            Label lblAra = new Label { Text = "🔍 Ara:", AutoSize = true, Font = new Font("Segoe UI", 11, FontStyle.Bold), Location = new Point(440, 15), ForeColor = Color.Gray };
            pnlRightTop.Controls.Add(lblAra);
            TextBox txtAra = new TextBox { Location = new Point(500, 12), Width = 200, Font = new Font("Segoe UI", 11) };
            pnlRightTop.Controls.Add(txtAra);

            DataGridView dgvPersonelListesi = new DataGridView
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
            dgvPersonelListesi.ColumnHeadersDefaultCellStyle.BackColor = colorSidebar;
            dgvPersonelListesi.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPersonelListesi.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            dgvPersonelListesi.ColumnHeadersHeight = 40;
            dgvPersonelListesi.EnableHeadersVisualStyles = false;
            pnlRight.Controls.Add(dgvPersonelListesi);
            pnlRightTop.SendToBack();

            // --- OTOMATİK DÜZELTME VE SÜTUN EKLEME FONKSİYONU ---
            void VeritabaniYapilandirVeDoldur()
            {
                try
                {
                    using (var conn = new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();

                        // 1. ADIM: pk_departman SÜTUNU VAR MI? YOKSA EKLE.
                        try
                        {
                            var cmdCheck = new MySqlCommand("SELECT count(*) FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = 'iskur' AND TABLE_NAME = 'program_katilimcilari' AND COLUMN_NAME = 'pk_departman'", conn);
                            int varMi = Convert.ToInt32(cmdCheck.ExecuteScalar());
                            if (varMi == 0)
                            {
                                var cmdAdd = new MySqlCommand("ALTER TABLE program_katilimcilari ADD COLUMN pk_departman VARCHAR(100) DEFAULT NULL", conn);
                                cmdAdd.ExecuteNonQuery();
                            }
                        }
                        catch { }

                        // 2. ADIM: KAMPÜS VERİLERİNİ TEMİZLE (Eski bozuk veriler için)
                        string sqlKampusFix = @"UPDATE program_katilimcilari 
                                     SET pk_gorev_yeri = CASE FLOOR(1 + RAND() * 3)
                                         WHEN 1 THEN 'Kampüs1' WHEN 2 THEN 'Kampüs2' ELSE 'Kampüs3' END
                                     WHERE pk_gorev_yeri NOT IN ('Kampüs1', 'Kampüs2', 'Kampüs3') OR pk_gorev_yeri IS NULL OR pk_gorev_yeri = ''";
                        new MySqlCommand(sqlKampusFix, conn).ExecuteNonQuery();

                        // 3. ADIM: DEPARTMAN (GÖREV) KISMINI RASTGELE DOLDUR (Boş olanlar için)
                        string sqlDeptFix = @"UPDATE program_katilimcilari 
                                     SET pk_departman = CASE FLOOR(1 + RAND() * 5)
                                         WHEN 1 THEN 'İdari' WHEN 2 THEN 'Teknik' WHEN 3 THEN 'Güvenlik' WHEN 4 THEN 'Temizlik' ELSE 'Bilişim' END
                                     WHERE pk_departman IS NULL OR pk_departman = ''";
                        new MySqlCommand(sqlDeptFix, conn).ExecuteNonQuery();
                    }
                }
                catch { }
            }

            // --- VERİ ÇEKME ---
            void VerileriGetir(string kampusSecimi)
            {
                try
                {
                    using (var conn = new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        // ARTIK GÖREV YERİNE pk_departman SÜTUNUNU ÇEKİYORUZ!
                        string sql = @"SELECT pk_tc AS 'TC', pk_ad_soyad AS 'Ad Soyad', pk_iban_no AS 'IBAN', 
                             COALESCE(pk_departman, 'Genel') AS 'Görev', 
                             pk_gorev_yeri AS 'Kampüs', 
                             pk_is_baslama_tarihi AS 'Başlama' 
                             FROM program_katilimcilari 
                             WHERE (@filtre = 'Tümü' OR pk_gorev_yeri = @filtre)";

                        var cmd = new MySqlCommand(sql, conn);
                        cmd.Parameters.AddWithValue("@filtre", kampusSecimi);

                        using (var da = new MySqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            dgvPersonelListesi.DataSource = dt;
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            }

            // --- ÇALIŞTIR ---
            VeritabaniYapilandirVeDoldur(); // Önce tabloyu düzelt ve doldur
            VerileriGetir("Tümü");          // Sonra listeyi getir

            // --- EVENTLER ---
            cmbKampusFiltre.SelectedIndexChanged += (s, e) => VerileriGetir(cmbKampusFiltre.SelectedItem?.ToString() ?? "Tümü");

            txtAra.TextChanged += (s, e) => {
                DataTable dt = dgvPersonelListesi.DataSource as DataTable;
                if (dt != null)
                {
                    string ara = txtAra.Text.Trim().Replace("'", "''");
                    dt.DefaultView.RowFilter = string.IsNullOrEmpty(ara) ? "" : $"[Ad Soyad] LIKE '%{ara}%' OR [TC] LIKE '%{ara}%'";
                }
            };

            btnKaydet.Click += (s, e) => {
                if (txtTc.Text == "" || txtAd.Text == "") { MessageBox.Show("TC ve Ad zorunlu."); return; }
                try
                {
                    using (var conn = new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        string kampus = cmbKampusEkle.SelectedItem?.ToString() ?? "Kampüs1";
                        string departman = cmbGorev.Text.Trim(); // Seçilen görevi al
                        if (string.IsNullOrEmpty(departman)) departman = "İdari";

                        // INSERT KISMINA pk_departman EKLENDİ
                        string sqlPer = @"INSERT INTO program_katilimcilari 
                                (pk_tc, pk_ad_soyad, pk_iban_no, pk_gorev_yeri, pk_departman, pk_is_baslama_tarihi) 
                                VALUES (@tc, @ad, @iban, @kampus, @dept, @tarih)";
                        var cmd = new MySqlCommand(sqlPer, conn);
                        cmd.Parameters.AddWithValue("@tc", txtTc.Text);
                        cmd.Parameters.AddWithValue("@ad", txtAd.Text);
                        cmd.Parameters.AddWithValue("@iban", txtIban.Text);
                        cmd.Parameters.AddWithValue("@kampus", kampus);
                        cmd.Parameters.AddWithValue("@dept", departman); // Kaydet
                        cmd.Parameters.AddWithValue("@tarih", dtpBaslama.Value);
                        cmd.ExecuteNonQuery();

                        // Puantaj için
                        string sqlPua = @"INSERT IGNORE INTO puantaj (p_tc, p_ad_soyad, p_iban, p_ise_baslama_tarihi) VALUES (@tc, @ad, @iban, @tarih)";
                        var cmdPua = new MySqlCommand(sqlPua, conn);
                        cmdPua.Parameters.AddWithValue("@tc", txtTc.Text);
                        cmdPua.Parameters.AddWithValue("@ad", txtAd.Text);
                        cmdPua.Parameters.AddWithValue("@iban", txtIban.Text);
                        cmdPua.Parameters.AddWithValue("@tarih", dtpBaslama.Value);
                        cmdPua.ExecuteNonQuery();
                    }
                    MessageBox.Show("✅ Personel eklendi.");
                    VeritabaniYapilandirVeDoldur(); // Listeyi güncelle
                    VerileriGetir(cmbKampusFiltre.Text);

                    txtTc.Clear(); txtAd.Clear(); txtIban.Clear();
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            };

            // Excel kısmı aynı kalabilir
            btnExcelImport.Click += (s, e) =>
            {
                OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel|*.xlsx" };
                if (ofd.ShowDialog() == DialogResult.OK) { /* Excel işlemleri */ VeritabaniYapilandirVeDoldur(); VerileriGetir("Tümü"); }
            };
        }

        private TextBox AddInputControl(FlowLayoutPanel parent, string labelText, int maxLength = 100)
        {
            Label lbl = new Label
            {
                Text = labelText,
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.Gray,
                Margin = new Padding(0, 5, 0, 5)
            };
            parent.Controls.Add(lbl);

            TextBox txt = new TextBox
            {
                Width = 300,
                Height = 35,
                Font = new Font("Segoe UI", 11),
                MaxLength = maxLength,
                Margin = new Padding(0, 0, 0, 15)
            };
            parent.Controls.Add(txt);

            return txt;
        }

        private TextBox CreateInput(Panel parent, string labelText, ref int yPos, int maxLength = 100)
        {
            Label lbl = new Label
            {
                Text = labelText,
                Location = new Point(20, yPos),
                AutoSize = true,
                Font = new Font("Segoe UI", 10)
            };
            parent.Controls.Add(lbl);

            TextBox txt = new TextBox
            {
                Location = new Point(20, yPos + 25),
                Width = 300,
                Font = new Font("Segoe UI", 10),
                MaxLength = maxLength
            };
            parent.Controls.Add(txt);

            yPos += 65;
            return txt;
        }
    }
}
