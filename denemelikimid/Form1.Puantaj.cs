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
        private void LoadPuantajView()
        {
            // 1. SAYFA TEMİZLİĞİ
            panelContent.Controls.Clear();

            // ANA DÜZENLEYİCİ
            TableLayoutPanel tlpMain = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = colorContent,
                Padding = new Padding(10)
            };
            tlpMain.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlpMain.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            panelContent.Controls.Add(tlpMain);

            // ÜST KISIM
            Panel pnlTopContainer = new Panel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 10)
            };

            Label lblHeader = new Label
            {
                Text = "📝 Personel Puantaj Girişi",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                Dock = DockStyle.Top,
                Height = 40
            };
            pnlTopContainer.Controls.Add(lblHeader);

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
                Text = "💾 Tümünü Kaydet",
                Size = new Size(160, 45),
                BackColor = colorPrimary,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 10, 5)
            };
            btnKaydet.FlatAppearance.BorderSize = 0;
            flowTools.Controls.Add(btnKaydet);

            pnlTopContainer.Controls.Add(flowTools);

            Label lblInfo = new Label
            {
                Text =
                    "ℹ️ Bilgi: Hücrelere tıklayarak durumu değiştirin (X: Çalıştı, İ: İzinli, R: Raporlu). Haftada maksimum 3 gün çalışılabilir.",
                AutoSize = true,
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 10, FontStyle.Italic),
                Dock = DockStyle.Bottom,
                Padding = new Padding(5, 5, 0, 0)
            };
            pnlTopContainer.Controls.Add(lblInfo);

            tlpMain.Controls.Add(pnlTopContainer, 0, 0);

            // TAB CONTROL - KAMPÜSLERE GÖRE AYRI SEKMELER
            TabControl tabControlKampusler = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10, FontStyle.Regular),
                Appearance = TabAppearance.FlatButtons
            };

            // Kampüsleri ve DataGridView'leri tutacak dictionary
            Dictionary<string, DataGridView> kampusGrids = new Dictionary<string, DataGridView>();

            // FONKSİYONLAR
            void KampusleriYukle()
            {
                tabControlKampusler.TabPages.Clear();
                kampusGrids.Clear();

                try
                {
                    using (var conn = new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        // Farklı kampüsleri çek (pk_gorev_yeri kolonu kampüs bilgisini tutar)
                        string sqlKampusler = @"
                            SELECT DISTINCT COALESCE(pk.pk_gorev_yeri, 'Diğer') AS kampus
                            FROM puantaj p
                            LEFT JOIN program_katilimcilari pk ON p.p_tc = pk.pk_tc
                            ORDER BY kampus";

                        var cmdKampusler = new MySqlCommand(sqlKampusler, conn);
                        var drKampusler = cmdKampusler.ExecuteReader();

                        List<string> kampusListesi = new List<string>();
                        while (drKampusler.Read())
                        {
                            string kampus = drKampusler["kampus"].ToString();
                            if (string.IsNullOrEmpty(kampus)) kampus = "Diğer";
                            kampusListesi.Add(kampus);
                        }
                        drKampusler.Close();

                        // Eğer hiç kampüs yoksa varsayılan kampüsler oluştur
                        if (kampusListesi.Count == 0)
                        {
                            kampusListesi.AddRange(new string[] { "Kampüs1", "Kampüs2", "Kampüs3" });
                        }

                        // Her kampüs için tab ve grid oluştur
                        foreach (string kampus in kampusListesi)
                        {
                            TabPage tabPage = new TabPage
                            {
                                Text = kampus,
                                Padding = new Padding(5),
                                BackColor = Color.White
                            };

                            DataGridView dgvPuantaj = new DataGridView
                            {
                                Dock = DockStyle.Fill,
                                BackgroundColor = Color.White,
                                AllowUserToAddRows = false,
                                RowHeadersVisible = false,
                                BorderStyle = BorderStyle.FixedSingle
                            };
                            dgvPuantaj.ColumnHeadersHeight = 40;
                            dgvPuantaj.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            dgvPuantaj.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            dgvPuantaj.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);

                            tabPage.Controls.Add(dgvPuantaj);
                            tabControlKampusler.TabPages.Add(tabPage);
                            kampusGrids[kampus] = dgvPuantaj;

                            // Grid yapısını oluştur ve verileri yükle
                            GridOlustur(dgvPuantaj);
                            VerileriYukle(dgvPuantaj, kampus);
                            HucreTiklamaOlayiEkle(dgvPuantaj);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Kampüsler yüklenirken hata: " + ex.Message);
                    // Hata durumunda varsayılan kampüsler oluştur
                    string[] varsayilanKampusler = { "Kampüs1", "Kampüs2", "Kampüs3" };
                    foreach (string kampus in varsayilanKampusler)
                    {
                        TabPage tabPage = new TabPage { Text = kampus };
                        DataGridView dgvPuantaj = new DataGridView { Dock = DockStyle.Fill };
                        tabPage.Controls.Add(dgvPuantaj);
                        tabControlKampusler.TabPages.Add(tabPage);
                        kampusGrids[kampus] = dgvPuantaj;
                    }
                }
            }

            void GridOlustur(DataGridView dgvPuantaj)
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
                    string baslik = i + "\n" + gunTarihi.ToString("ddd", new System.Globalization.CultureInfo("tr-TR"));

                    dgvPuantaj.Columns.Add("day" + i, baslik);
                    dgvPuantaj.Columns[i + 1].Width = 45;

                    if (gunTarihi.DayOfWeek == DayOfWeek.Saturday || gunTarihi.DayOfWeek == DayOfWeek.Sunday)
                    {
                        dgvPuantaj.Columns[i + 1].DefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
                    }
                }
            }

            void VerileriYukle(DataGridView dgvPuantaj, string kampus)
            {
                try
                {
                    using (var conn = new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        string sql = @"
                            SELECT p.p_tc, p.p_ad_soyad, p.p_gun_detaylari 
                            FROM puantaj p
                            LEFT JOIN program_katilimcilari pk ON p.p_tc = pk.pk_tc
                            WHERE COALESCE(pk.pk_gorev_yeri, 'Diğer') = @kampus";

                        var cmd = new MySqlCommand(sql, conn);
                        cmd.Parameters.AddWithValue("@kampus", kampus == "Diğer" ? "" : kampus);
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
                catch (Exception ex)
                {
                    MessageBox.Show("Veriler yüklenirken hata: " + ex.Message);
                }
            }

            bool HaftalikLimitAsildiMi(DataGridView dgvPuantaj, int rowIndex, int gunSutunIndex)
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

                return buHaftakiXSayisi >= 3 && suAnXDegil;
            }

            void HucreTiklamaOlayiEkle(DataGridView dgvPuantaj)
            {
                dgvPuantaj.CellClick += (s, e) =>
                {
                    if (e.RowIndex >= 0 && e.ColumnIndex >= 2)
                    {
                        var cell = dgvPuantaj.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        string val = cell.Value?.ToString() ?? "";

                        if (val == "")
                        {
                            if (HaftalikLimitAsildiMi(dgvPuantaj, e.RowIndex, e.ColumnIndex))
                            {
                                MessageBox.Show("Bu hafta için maksimum 3 gün çalışma limiti doldu!", "Uyarı",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                            cell.Value = "X";
                            cell.Style.BackColor = Color.LightGreen;
                        }
                        else if (val == "X")
                        {
                            cell.Value = "İ";
                            cell.Style.BackColor = Color.LightYellow;
                        }
                        else if (val == "İ")
                        {
                            cell.Value = "R";
                            cell.Style.BackColor = Color.LightPink;
                        }
                        else
                        {
                            cell.Value = "";
                            cell.Style.BackColor = Color.White;
                        }
                    }
                };
            }

            // OLAYLAR
            KampusleriYukle();

            dtpDonem.ValueChanged += (s, e) =>
            {
                secilenTarih = dtpDonem.Value;
                KampusleriYukle();
            };

            btnKaydet.Click += (s, e) =>
            {
                try
                {
                    using (var conn = new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        foreach (var kvp in kampusGrids)
                        {
                            DataGridView dgvPuantaj = kvp.Value;
                            foreach (DataGridViewRow row in dgvPuantaj.Rows)
                            {
                                string tc = row.Cells[0].Value?.ToString();
                                if (string.IsNullOrEmpty(tc)) continue;

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

                                string sql =
                                    "UPDATE puantaj SET p_gun_detaylari = @detay, p_calistigi_gun_sayisi = @toplam, p_yil_ay = @donem WHERE p_tc = @tc";
                                var cmd = new MySqlCommand(sql, conn);
                                cmd.Parameters.AddWithValue("@detay", detayString);
                                cmd.Parameters.AddWithValue("@toplam", toplamCalisilanGun);
                                cmd.Parameters.AddWithValue("@donem", secilenTarih.ToString("yyyy-MM"));
                                cmd.Parameters.AddWithValue("@tc", tc);
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    MessageBox.Show("✅ Tüm puantajlar başarıyla kaydedildi!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hata: " + ex.Message);
                }
            };

            btnExcelExport.Click += (s, e) =>
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        foreach (var kvp in kampusGrids)
                        {
                            string kampus = kvp.Key;
                            DataGridView dgvPuantaj = kvp.Value;

                            var ws = workbook.Worksheets.Add(kampus);

                            ws.Cell(1, 1).Value = "BURSA ULUDAĞ ÜNİVERSİTESİ";
                            ws.Range(1, 1, 1, dgvPuantaj.Columns.Count).Merge().Style.Font.Bold = true;
                            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ws.Cell(2, 1).Value = dtpDonem.Value.ToString("MMMM yyyy").ToUpper() + " PUANTAJ CETVELİ - " + kampus.ToUpper();
                            ws.Range(2, 1, 2, dgvPuantaj.Columns.Count).Merge().Style.Font.Bold = true;
                            ws.Cell(2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

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

                                    if (val == "X")
                                        ws.Cell(i + 5, j + 1).Style.Fill.BackgroundColor = XLColor.LightGreen;
                                    if (val == "İ")
                                        ws.Cell(i + 5, j + 1).Style.Fill.BackgroundColor = XLColor.LightYellow;
                                    if (val == "R")
                                        ws.Cell(i + 5, j + 1).Style.Fill.BackgroundColor = XLColor.LightPink;
                                }
                            }
                            ws.Columns().AdjustToContents();
                        }

                        SaveFileDialog sfd = new SaveFileDialog
                        {
                            Filter = "Excel Dosyası|*.xlsx",
                            FileName = "Puantaj_Cizelgesi_" + dtpDonem.Value.ToString("yyyy_MM") + ".xlsx"
                        };
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            workbook.SaveAs(sfd.FileName);
                            MessageBox.Show("Excel oluşturuldu! Her kampüs için ayrı sayfa oluşturuldu.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Excel Hatası: " + ex.Message);
                }
            };

            tlpMain.Controls.Add(tabControlKampusler, 0, 1);
        }
    }
}
