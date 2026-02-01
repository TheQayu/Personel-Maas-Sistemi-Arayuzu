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
            Label lblUcret = new Label { Text = "Günlük Brüt (TL):", Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.Gray };
            NumericUpDown numUcret = new NumericUpDown { Location = new Point(0, 25), Width = 130, Height = 35, Maximum = 10000, DecimalPlaces = 2, Value = 666.75M, Font = new Font("Segoe UI", 11) };
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
                try
                {
                    using (var conn = new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();

                        string secilenKampus = cmbKampusFiltre.SelectedItem?.ToString() ?? "Tümü";

                        string sql = @"SELECT b_tc AS 'TC', b_ad_soyad AS 'Ad Soyad', 
                             b_gorev_yeri AS 'Kampüs',
                             b_aylik_calisilan_gun AS 'Gün', 
                             b_tahakkuk_toplami AS 'Brüt', 
                             b_odenmesi_gereken_net_tutar AS 'NET MAAŞ' 
                             FROM bordro 
                             WHERE (@filtre = 'Tümü' OR b_gorev_yeri = @filtre)";

                        var cmd = new MySqlCommand(sql, conn);
                        cmd.Parameters.AddWithValue("@filtre", secilenKampus);

                        using (var da = new MySqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            dgvBordro.DataSource = dt;
                        }
                    }
                }
                catch { }
            }

            // İlk açılışta listele
            ListeyiGuncelle();

            // Filtre değişince güncelle
            cmbKampusFiltre.SelectedIndexChanged += (s, e) => ListeyiGuncelle();

            // --- HESAPLA BUTONU ---
            btnHesapla.Click += (s, e) => {
                try
                {
                    decimal gunlukBruk = numUcret.Value;
                    using (var conn = new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();

                        // Önce temizle
                        new MySqlCommand("TRUNCATE TABLE bordro", conn).ExecuteNonQuery();
                        new MySqlCommand("TRUNCATE TABLE muhtasar_raporu", conn).ExecuteNonQuery();
                        new MySqlCommand("TRUNCATE TABLE banka_listesi", conn).ExecuteNonQuery();

                        // Puantajdan verileri al (Personel tablosuyla birleştirip Kampüsü de alıyoruz)
                        string sqlPuantaj = @"SELECT p.*, COALESCE(pk.pk_gorev_yeri, 'Kampüs1') AS guncel_kampus 
                                      FROM puantaj p
                                      LEFT JOIN program_katilimcilari pk ON p.p_tc = pk.pk_tc
                                      WHERE p.p_calistigi_gun_sayisi > 0";

                        var cmdGet = new MySqlCommand(sqlPuantaj, conn);
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
                            decimal brutUcret = gun * gunlukBruk;
                            decimal sgkPrimi = brutUcret * 0.14M;
                            decimal damgaVergisi = brutUcret * 0.00759M;
                            decimal gelirVergisiMatrahi = brutUcret - sgkPrimi;
                            decimal gelirVergisi = gelirVergisiMatrahi * 0.15M;
                            decimal netUcret = brutUcret - (sgkPrimi + damgaVergisi + gelirVergisi);

                            // Bordroya Ekle (KAMPÜS BİLGİSİYLE BERABER)
                            string sqlBordro = @"INSERT INTO bordro 
                        (b_tc, b_ad_soyad, b_gorev_yeri, b_aylik_calisilan_gun, b_tahakkuk_toplami, b_sosyal_guvenlik_primi, b_gelir_vergisi_kesintisi, b_damga_vergisi_kesintisi, b_odenmesi_gereken_net_tutar) 
                        VALUES (@tc, @ad, @kampus, @gun, @brut, @sgk, @gv, @dv, @net)";

                            using (var cmd = new MySqlCommand(sqlBordro, conn))
                            {
                                cmd.Parameters.AddWithValue("@tc", tc);
                                cmd.Parameters.AddWithValue("@ad", ad);
                                cmd.Parameters.AddWithValue("@kampus", kampus); // <-- Önemli: Kampüsü kaydediyoruz
                                cmd.Parameters.AddWithValue("@gun", gun);
                                cmd.Parameters.AddWithValue("@brut", brutUcret);
                                cmd.Parameters.AddWithValue("@sgk", sgkPrimi);
                                cmd.Parameters.AddWithValue("@gv", gelirVergisi);
                                cmd.Parameters.AddWithValue("@dv", damgaVergisi);
                                cmd.Parameters.AddWithValue("@net", netUcret);
                                cmd.ExecuteNonQuery();
                            }

                            // Muhtasar ve Banka tablolarına ekleme kısımları aynen devam...
                            string sqlMuhtasar = "INSERT INTO muhtasar_raporu (mh_tc, mh_ad_soyad, mh_prim_odeme_gunu, mh_hak_edilen_ucret, mh_doneme_ait_gelir_vergisi_matrahi, mh_gelir_vergisi_kesintisi, mh_damga_vergisi_kesintisi) VALUES (@tc, @ad, @gun, @brut, @matrah, @gv, @dv)";
                            using (var cmd = new MySqlCommand(sqlMuhtasar, conn))
                            {
                                cmd.Parameters.AddWithValue("@tc", tc); cmd.Parameters.AddWithValue("@ad", ad); cmd.Parameters.AddWithValue("@gun", gun); cmd.Parameters.AddWithValue("@brut", brutUcret); cmd.Parameters.AddWithValue("@matrah", gelirVergisiMatrahi); cmd.Parameters.AddWithValue("@gv", gelirVergisi); cmd.Parameters.AddWithValue("@dv", damgaVergisi); cmd.ExecuteNonQuery();
                            }

                            string sqlBanka = "INSERT INTO banka_listesi (bl_tc, bl_ad_soyad, bl_iban_no, bl_tutar) VALUES (@tc, @ad, @iban, @net)";
                            using (var cmd = new MySqlCommand(sqlBanka, conn))
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

            // --- EXCEL BUTONLARI ---
            btnMuhtasar.Click += (s, e) => ExportTableToExcel("muhtasar_raporu", "Muhtasar_Raporu");

            // Bordro İndir Butonu (Filtreye göre indirebilir)
            btnBordro.Click += (s, e) => ExportTableToExcel("bordro", "Personel_Bordrosu");
        }

        private Button CreateActionButton(string text, Color color)
        {
            return new Button
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
        }

        private void ExportTableToExcel(string tableName, string fileName)
        {
            try
            {
                DataTable dt = new DataTable();
                using (var conn = new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                {
                    conn.Open();
                    using (var da = new MySqlDataAdapter($"SELECT * FROM {tableName}", conn))
                    {
                        da.Fill(dt);
                    }
                }

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("Tabloda veri yok. Önce 'Hesapla' butonuna basın.");
                    return;
                }

                using (var workbook = new XLWorkbook())
                {
                    // Bordro tablosu için kampüslere göre ayrı sayfalar, muhtasar için tek sayfa
                    if (tableName == "bordro" && dt.Columns.Contains("b_gorev_yeri"))
                    {
                        // Kampüslere göre ayrı sayfalar oluştur
                        var kampusler = dt.AsEnumerable()
                            .Select(row => row.Field<string>("b_gorev_yeri") ?? "Diğer")
                            .Where(k => !string.IsNullOrEmpty(k))
                            .Distinct()
                            .ToList();

                        if (kampusler.Count == 0)
                        {
                            kampusler.AddRange(new string[] { "Kampüs1", "Kampüs2", "Kampüs3" });
                        }

                        foreach (string kampus in kampusler)
                        {
                            var ws = workbook.Worksheets.Add(kampus);
                            var kampusRows = dt.AsEnumerable()
                                .Where(row => (row.Field<string>("b_gorev_yeri") ?? "Diğer") == kampus)
                                .ToList();
                            
                            if (kampusRows.Count > 0)
                            {
                                var kampusDt = kampusRows.CopyToDataTable();
                                ws.Cell(1, 1).InsertTable(kampusDt);
                                ws.Columns().AdjustToContents();
                            }
                        }

                        SaveFileDialog sfd = new SaveFileDialog
                        {
                            Filter = "Excel Dosyası|*.xlsx",
                            FileName = $"{fileName}_{DateTime.Now:yyyy-MM}.xlsx"
                        };
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            workbook.SaveAs(sfd.FileName);
                            MessageBox.Show("✅ Dosya kaydedildi. Her kampüs için ayrı sayfa oluşturuldu.");
                        }
                    }
                    else
                    {
                        // Muhtasar veya diğer tablolar için tek sayfa
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
