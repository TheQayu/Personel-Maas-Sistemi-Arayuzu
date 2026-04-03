using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using denemelikimid.DataBase;

namespace denemelikimid
{
    public partial class Form1
    {
        private void LoadRaporlarView()
        {
            // --- 1. SAYFA İSKELETİ ---
            panelContent.Controls.Clear();

            // Ana Düzen (TableLayout ile kaymayı önlüyoruz)
            TableLayoutPanel tlpMain = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = colorContent,
                Padding = new Padding(10)
            };
            tlpMain.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // Üst kısım otomatik
            tlpMain.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // Alt kısım %100
            panelContent.Controls.Add(tlpMain);

            // --- 2. ÜST PANEL (BAŞLIK, KUTULAR, BUTONLAR) ---
            Panel pnlTopContainer = new Panel { AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 10) };

            Label lblHeader = new Label
            {
                Text = "📊 Bordro ve Muhtasar İşlemleri",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                Dock = DockStyle.Top,
                Height = 45
            };
            pnlTopContainer.Controls.Add(lblHeader);

            // Araç Çubuğu
            FlowLayoutPanel flowTools = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(0, 10, 0, 0)
            };

            // Günlük Ücret Girişi
            Panel pnlUcret = new Panel { Width = 140, Height = 60, Margin = new Padding(0, 0, 10, 0) };
            Label lblUcret = new Label { Text = "Günlük Net (TL):", Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.Gray };
            NumericUpDown numUcret = new NumericUpDown { Location = new Point(0, 25), Width = 130, Height = 35, Maximum = 10000, DecimalPlaces = 2, Value = 1375.00M, Font = new Font("Segoe UI", 11) };
            pnlUcret.Controls.Add(lblUcret); pnlUcret.Controls.Add(numUcret);
            flowTools.Controls.Add(pnlUcret);

            // KAMPÜS FİLTRELEME KUTUSU (YENİ EKLENDİ)
            Panel pnlFiltre = new Panel { Width = 160, Height = 60, Margin = new Padding(0, 0, 10, 0) };
            Label lblFiltre = new Label { Text = "Kampüs Filtrele:", Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.Gray };
            ComboBox cmbKampusFiltre = new ComboBox { Location = new Point(0, 25), Width = 150, Height = 35, Font = new Font("Segoe UI", 11), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbKampusFiltre.Items.AddRange(new string[] { "Tümü", "Kampüs1", "Kampüs2", "Kampüs3" });
            cmbKampusFiltre.SelectedIndex = 0;
            pnlFiltre.Controls.Add(lblFiltre); pnlFiltre.Controls.Add(cmbKampusFiltre);
            flowTools.Controls.Add(pnlFiltre);

            // AY SEÇİCİ
            Panel pnlDonem = new Panel { Width = 160, Height = 60, Margin = new Padding(0, 0, 10, 0) };
            Label lblDonem = new Label { Text = "Dönem (Ay):", Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.Gray };
            DateTimePicker dtpRaporDonem = new DateTimePicker
            {
                Location = new Point(0, 25),
                Width = 150,
                Height = 35,
                Font = new Font("Segoe UI", 11),
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "MMMM yyyy",
                ShowUpDown = true
            };
            pnlDonem.Controls.Add(lblDonem); pnlDonem.Controls.Add(dtpRaporDonem);
            flowTools.Controls.Add(pnlDonem);

            // Butonlar
            Button btnHesapla = CreateActionButton("⚙️ 1. Hesapla", Color.Orange);
            Button btnMuhtasar = CreateActionButton("📄 2. Muhtasar İndir", colorSuccess);
            Button btnBordro = CreateActionButton("📑 3. Bordro İndir", colorInfo);

            flowTools.Controls.Add(btnHesapla);
            flowTools.Controls.Add(btnMuhtasar);
            flowTools.Controls.Add(btnBordro);

            pnlTopContainer.Controls.Add(flowTools);
            flowTools.BringToFront(); lblHeader.BringToFront();
            tlpMain.Controls.Add(pnlTopContainer, 0, 0);


            // --- 3. ALT TABLO (TEK BİR GRID) ---
            DataGridView dgvBordro = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvBordro.ColumnHeadersDefaultCellStyle.BackColor = colorSidebar;
            dgvBordro.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvBordro.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            dgvBordro.ColumnHeadersHeight = 40;
            dgvBordro.EnableHeadersVisualStyles = false;

            tlpMain.Controls.Add(dgvBordro, 0, 1);

            // --- FONKSİYONLAR ---

            // Listeyi Veritabanından Çek ve Filtrele
            void ListeyiGuncelle()
            {
                string secilenKampus = cmbKampusFiltre.SelectedItem?.ToString() ?? "Tümü";
                string donem = dtpRaporDonem.Value.ToString("yyyy-MM");

                Task.Run(() => FetchBordroData(secilenKampus, donem))
                    .ContinueWith(t =>
                    {
                        if (t.IsFaulted)
                        {
                            var ex = t.Exception?.GetBaseException() ?? t.Exception;
                            MessageBox.Show("Veriler yüklenirken hata: " + ex?.Message);
                            return;
                        }

                        if (dgvBordro.IsDisposed)
                        {
                            return;
                        }

                        dgvBordro.DataSource = t.Result;
                    }, TaskScheduler.FromCurrentSynchronizationContext());
            }

            DataTable FetchBordroData(string secilenKampus, string donem)
            {
                using (var conn = DbConnection.GetConnection())
                {
                    conn.Open();

                    string sql = @"SELECT b_tc AS 'TC', b_ad_soyad AS 'Ad Soyad', 
                         b_gorev_yeri AS 'Kampüs',
                         b_aylik_calisilan_gun AS 'Gün', 
                         b_odenmesi_gereken_net_tutar AS 'NET MAAŞ' 
                         FROM bordro 
                         WHERE (@filtre = 'Tümü' OR b_gorev_yeri = @filtre)
                           AND b_yil_ay = @donem";

                    var cmd = new SqliteCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@filtre", secilenKampus);
                    cmd.Parameters.AddWithValue("@donem", donem);

                    using (var reader = cmd.ExecuteReader())
                    {
                        DataTable dt = new DataTable();
                        dt.Load(reader);
                        return dt;
                    }
                }
            }

            // İlk açılışta listele
            ListeyiGuncelle();

            // Filtre değişince güncelle
            cmbKampusFiltre.SelectedIndexChanged += (s, e) => ListeyiGuncelle();
            dtpRaporDonem.ValueChanged += (s, e) => ListeyiGuncelle();

            // --- HESAPLA BUTONU ---
            btnHesapla.Click += (s, e) => {
                try
                {
                    decimal gunlukBruk = numUcret.Value;
                    using (var conn = DbConnection.GetConnection())
                    {
                        conn.Open();

                        // Çoklu ay desteği için tüm ilgili tablolardaki UNIQUE kısıtlarını kaldır
                        try
                        {
                            var tables = new[] { "puantaj", "bordro", "muhtasar_raporu", "banka_listesi" };
                            foreach (var table in tables)
                            {
                                using (var idxCmd = new SqliteCommand($"PRAGMA index_list('{table}')", conn))
                                using (var reader = idxCmd.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        var isUnique = Convert.ToInt32(reader["unique"]) == 1;
                                        var indexName = reader["name"].ToString();
                                        if (isUnique && !string.IsNullOrWhiteSpace(indexName))
                                        {
                                            var safeName = indexName.Replace("\"", "\"\"");
                                            try { new SqliteCommand($"DROP INDEX IF EXISTS \"{safeName}\"", conn).ExecuteNonQuery(); } catch { }
                                        }
                                    }
                                }
                            }
                        } catch { }

                        // b_yil_ay sütununu ekle (zaten varsa hata yutulur)
                        try { new SqliteCommand("ALTER TABLE bordro ADD COLUMN b_yil_ay TEXT", conn).ExecuteNonQuery(); } catch { }
                        try { new SqliteCommand("ALTER TABLE muhtasar_raporu ADD COLUMN mh_yil_ay TEXT", conn).ExecuteNonQuery(); } catch { }
                        try { new SqliteCommand("ALTER TABLE banka_listesi ADD COLUMN bl_yil_ay TEXT", conn).ExecuteNonQuery(); } catch { }

                        // Seçili ayın verilerini temizle (diğer aylar korunur)
                        string secilenDonem = dtpRaporDonem.Value.ToString("yyyy-MM");
                        using (var delCmd = new SqliteCommand("DELETE FROM bordro WHERE b_yil_ay = @d", conn)) { delCmd.Parameters.AddWithValue("@d", secilenDonem); delCmd.ExecuteNonQuery(); }
                        using (var delCmd = new SqliteCommand("DELETE FROM muhtasar_raporu WHERE mh_yil_ay = @d", conn)) { delCmd.Parameters.AddWithValue("@d", secilenDonem); delCmd.ExecuteNonQuery(); }
                        using (var delCmd = new SqliteCommand("DELETE FROM banka_listesi WHERE bl_yil_ay = @d", conn)) { delCmd.Parameters.AddWithValue("@d", secilenDonem); delCmd.ExecuteNonQuery(); }

                        // Puantajdan verileri al (Personel tablosuyla birleştirip Kampüsü de alıyoruz)
                        string sqlPuantaj = @"SELECT p.*, COALESCE(pk.pk_gorev_yeri, 'Kampüs1') AS guncel_kampus 
                                      FROM puantaj p
                                      LEFT JOIN program_katilimcilari pk ON p.p_tc = pk.pk_tc
                                      WHERE p.p_calistigi_gun_sayisi > 0
                                        AND p.p_yil_ay = @donem";

                        var cmdGet = new SqliteCommand(sqlPuantaj, conn);
                        cmdGet.Parameters.AddWithValue("@donem", secilenDonem);
                        var dr = cmdGet.ExecuteReader();
                        DataTable dtPuantaj = new DataTable();
                        dtPuantaj.Load(dr);

                        foreach (DataRow row in dtPuantaj.Rows)
                        {
                            string tc = row["p_tc"].ToString();
                            string ad = row["p_ad_soyad"].ToString();
                            string iban = row["p_iban"].ToString();
                            int gun = Convert.ToInt32(row["p_calistigi_gun_sayisi"]);
                            string kampus = row["guncel_kampus"].ToString(); // Kampüs bilgisi

                            // Maaş Hesabı
                            decimal netUcret = gun * gunlukBruk;
                            decimal brutUcret = netUcret;
                            decimal sgkPrimi = 0M;
                            decimal damgaVergisi = 0M;
                            decimal gelirVergisiMatrahi = 0M;
                            decimal gelirVergisi = 0M;

                            // Bordroya Ekle (KAMPÜS BİLGİSİYLE BERABER)
                            string sqlBordro = @"INSERT INTO bordro 
                        (b_tc, b_ad_soyad, b_gorev_yeri, b_aylik_calisilan_gun, b_tahakkuk_toplami, b_sosyal_guvenlik_primi, b_gelir_vergisi_kesintisi, b_damga_vergisi_kesintisi, b_odenmesi_gereken_net_tutar, b_yil_ay) 
                        VALUES (@tc, @ad, @kampus, @gun, @brut, @sgk, @gv, @dv, @net, @donem)";

                            using (var cmd = new SqliteCommand(sqlBordro, conn))
                            {
                                cmd.Parameters.AddWithValue("@tc", tc);
                                cmd.Parameters.AddWithValue("@ad", ad);
                                cmd.Parameters.AddWithValue("@kampus", kampus);
                                cmd.Parameters.AddWithValue("@gun", gun);
                                cmd.Parameters.AddWithValue("@brut", brutUcret);
                                cmd.Parameters.AddWithValue("@sgk", sgkPrimi);
                                cmd.Parameters.AddWithValue("@gv", gelirVergisi);
                                cmd.Parameters.AddWithValue("@dv", damgaVergisi);
                                cmd.Parameters.AddWithValue("@net", netUcret);
                                cmd.Parameters.AddWithValue("@donem", secilenDonem);
                                cmd.ExecuteNonQuery();
                            }

                            // Muhtasar ve Banka tablolarına ekleme kısımları aynen devam...
                            string sqlMuhtasar = "INSERT INTO muhtasar_raporu (mh_tc, mh_ad_soyad, mh_prim_odeme_gunu, mh_hak_edilen_ucret, mh_doneme_ait_gelir_vergisi_matrahi, mh_gelir_vergisi_kesintisi, mh_damga_vergisi_kesintisi, mh_yil_ay) VALUES (@tc, @ad, @gun, @brut, @matrah, @gv, @dv, @donem)";
                            using (var cmd = new SqliteCommand(sqlMuhtasar, conn))
                            {
                                cmd.Parameters.AddWithValue("@tc", tc); cmd.Parameters.AddWithValue("@ad", ad); cmd.Parameters.AddWithValue("@gun", gun); cmd.Parameters.AddWithValue("@brut", brutUcret); cmd.Parameters.AddWithValue("@matrah", gelirVergisiMatrahi); cmd.Parameters.AddWithValue("@gv", gelirVergisi); cmd.Parameters.AddWithValue("@dv", damgaVergisi); cmd.Parameters.AddWithValue("@donem", secilenDonem); cmd.ExecuteNonQuery();
                            }

                            string sqlBanka = "INSERT INTO banka_listesi (bl_tc, bl_ad_soyad, bl_iban_no, bl_tutar, bl_yil_ay) VALUES (@tc, @ad, @iban, @net, @donem)";
                            using (var cmd = new SqliteCommand(sqlBanka, conn))
                            {
                                cmd.Parameters.AddWithValue("@tc", tc); cmd.Parameters.AddWithValue("@ad", ad); cmd.Parameters.AddWithValue("@iban", iban); cmd.Parameters.AddWithValue("@net", netUcret); cmd.Parameters.AddWithValue("@donem", secilenDonem); cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    MessageBox.Show("✅ Hesaplama tamamlandı!");
                    ListeyiGuncelle();
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            };

            // --- EXCEL BUTONLARI ---
            btnMuhtasar.Click += (s, e) => ExportTableToExcel("muhtasar_raporu", "Muhtasar_Raporu");

            // Bordro İndir Butonu (Filtreye göre indirebilir)
            btnBordro.Click += (s, e) => ExportTableToExcel("bordro", "Personel_Bordrosu");
        }

        private Button CreateActionButton(string text, Color color)
        {
            var btn = new Button
            {
                Text = text,
                Size = new Size(200, 55),
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 15, 10)
            };
            btn.FlatAppearance.BorderSize = 0;
            ApplyRoundedCorners(btn, 12);
            return btn;
        }

        private void ExportTableToExcel(string tableName, string fileName)
        {
            try
            {
                DataTable dt = new DataTable();
                using (var conn = DbConnection.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand($"SELECT * FROM {tableName}", conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        dt.Load(reader);
                    }
                }

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("Tabloda veri yok. Önce 'Hesapla' butonuna basın.");
                    return;
                }

                using (var workbook = new XLWorkbook())
                {
                    // Bordro tablosu için kampüslere göre ayrı dosyalar (BUÜ formatı)
                    if (tableName == "bordro" && dt.Columns.Contains("b_gorev_yeri"))
                    {
                        string defaultTemplate1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resmi_Sablon.xlsx");
                        string defaultTemplate2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resmi_Sablon.xls");
                        string templatePath = null;
                        if (File.Exists(defaultTemplate1)) templatePath = defaultTemplate1;
                        else if (File.Exists(defaultTemplate2)) templatePath = defaultTemplate2;
                        else
                        {
                            OpenFileDialog ofdTemplate = new OpenFileDialog()
                            {
                                Filter = "Excel Şablonu|*.xlsx;*.xls",
                                Title = "Lütfen resmi şablon Excel dosyasını seçin"
                            };
                            if (ofdTemplate.ShowDialog() != DialogResult.OK)
                            {
                                MessageBox.Show("Şablon seçilmedi. İşlem iptal edildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                            templatePath = ofdTemplate.FileName;
                        }

                        var kampusler = dt.AsEnumerable()
                            .Select(row => row.Field<string>("b_gorev_yeri") ?? "Diğer")
                            .Where(k => !string.IsNullOrEmpty(k) && (k == "Kampüs1" || k == "Kampüs2" || k == "Kampüs3"))
                            .Distinct()
                            .ToList();

                        if (kampusler.Count == 0)
                        {
                            kampusler.AddRange(new string[] { "Kampüs1", "Kampüs2", "Kampüs3" });
                        }

                        foreach (string kampus in kampusler)
                        {
                            string kampusNo = "13376";
                            if (kampus == "Kampüs2") kampusNo = "13377";
                            else if (kampus == "Kampüs3") kampusNo = "13378";

                            string buuFileName = $"{kampusNo} BUÜ BORDRO VE PUANTAJ {DateTime.Now:MMMM yyyy}";

                            SaveFileDialog sfd = new SaveFileDialog
                            {
                                Filter = "Excel Dosyası|*.xlsx",
                                FileName = buuFileName + ".xlsx"
                            };

                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                using (var kampusWorkbook = new XLWorkbook(templatePath))
                                {
                                    var bordroSheet = kampusWorkbook.Worksheets.FirstOrDefault(w => w.Name.Replace(" ", "").ToUpperInvariant().Contains("BORDRO"));
                                    if (bordroSheet != null)
                                    {
                                        string kampusLabel = kampus.Replace("Kampüs", "KAMPÜS ");
                                        string headerText = $"BURSA ULUDAĞ ÜNİVERSİTESİ GENEL SEKRETERLİK ÖZEL KALEM {kampusNo} PORTAL NOLU {kampusLabel} İŞKUR GENÇLİK PROGRAMI ÖDEME BORDROSU";
                                        var headerCell = bordroSheet.Cell(2, 1);
                                        var merged = headerCell.MergedRange();
                                        if (merged != null)
                                            merged.Value = headerText;
                                        else
                                            headerCell.Value = headerText;
                                    }

                                    kampusWorkbook.SaveAs(sfd.FileName);
                                }

                                MessageBox.Show($"✅ {kampus} için dosya kaydedildi: {System.IO.Path.GetFileName(sfd.FileName)}");
                            }
                        }
                    }
                    else
                    {
                        var ws = workbook.Worksheets.Add(tableName == "muhtasar_raporu" ? "Muhtasar Raporu" : "Rapor");
                        ws.Cell(1, 1).InsertTable(dt);
                        ws.Columns().AdjustToContents();

                        SaveFileDialog sfd = new SaveFileDialog
                        {
                            Filter = "Excel Dosyası|*.xlsx",
                            FileName = $"{fileName}_{DateTime.Now:yyyy-MM}.xlsx"
                        };
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            workbook.SaveAs(sfd.FileName);
                            MessageBox.Show("✅ Dosya başarıyla kaydedildi.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel Hatası: " + ex.Message);
            }
        }
    }
}
