using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using MySql.Data.MySqlClient;
using denemelikimid.DataBase;
using System.Globalization;
using System.Threading.Tasks;
using System.IO;

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
                Padding = new Padding(0)
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
            ApplyRoundedCorners(btnExcelExport, 12);
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
            ApplyRoundedCorners(btnKaydet, 12);
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
                    List<string> kampusListesi = new List<string>();
                    using (var conn = NotDbConnection.GetConnection())
                    {
                        // Farklı kampüsleri çek (pk_gorev_yeri kolonu kampüs bilgisini tutar)
                        string sqlKampusler = @"
                            SELECT DISTINCT COALESCE(pk.pk_gorev_yeri, 'Diğer') AS kampus
                            FROM puantaj p
                            LEFT JOIN program_katilimcilari pk ON p.p_tc = pk.pk_tc
                            ORDER BY kampus";

                        var cmdKampusler = new MySqlCommand(sqlKampusler, conn);
                        using (var drKampusler = cmdKampusler.ExecuteReader())
                        {
                            while (drKampusler.Read())
                            {
                                string kampus = drKampusler["kampus"].ToString();
                                if (string.IsNullOrEmpty(kampus)) kampus = "Diğer";
                                kampusListesi.Add(kampus);
                            }
                        }

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

                            // Daha iyi performans için DoubleBuffered
                            typeof(DataGridView).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(dgvPuantaj, true, null);

                            // Only allow changes via program (click handler) to prevent manual typing
                            dgvPuantaj.EditMode = DataGridViewEditMode.EditProgrammatically;
                            dgvPuantaj.SelectionMode = DataGridViewSelectionMode.CellSelect;
                            dgvPuantaj.ColumnHeadersHeight = 66;
                            dgvPuantaj.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            dgvPuantaj.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            dgvPuantaj.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                            dgvPuantaj.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
                            dgvPuantaj.EnableHeadersVisualStyles = false;
                            dgvPuantaj.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(238, 244, 255);
                            dgvPuantaj.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(50, 50, 50);
                            dgvPuantaj.RowTemplate.Height = 32;
                            dgvPuantaj.AllowUserToResizeRows = false;
                            dgvPuantaj.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;

                            // Özel çizim ve satır vurgulama
                            dgvPuantaj.CellEnter += (s, e) => {
                                dgvPuantaj.InvalidateRow(e.RowIndex);
                            };
                            dgvPuantaj.CellLeave += (s, e) => {
                                dgvPuantaj.InvalidateRow(e.RowIndex);
                            };

                            dgvPuantaj.CellPainting += (s, e) => {
                                if (e.RowIndex >= 0 && e.ColumnIndex >= 2)
                                {
                                    bool isSelectedRow = (dgvPuantaj.CurrentCell != null && dgvPuantaj.CurrentCell.RowIndex == e.RowIndex);
                                    Color backColor = e.CellStyle.BackColor;

                                    // Sadece boyanmamış (beyaz/boş veya haftasonu gri) hücrelerde seçili satır efekti yap
                                    bool isDefaultColor = (backColor.ToArgb() == Color.White.ToArgb() || backColor.ToArgb() == Color.Empty.ToArgb() || backColor.ToArgb() == Color.FromArgb(245, 245, 245).ToArgb());

                                    if (isSelectedRow && isDefaultColor)
                                    {
                                        backColor = Color.FromArgb(220, 235, 255); // hafif mavi 
                                    }

                                    using (SolidBrush bgBrush = new SolidBrush(backColor))
                                    {
                                        e.Graphics.FillRectangle(bgBrush, e.CellBounds);
                                    }

                                    // Gün numarasını filigran (watermark) olarak çiz
                                    string dayNum = (e.ColumnIndex - 1).ToString();
                                    using (Font wFont = new Font("Segoe UI", 12, FontStyle.Bold))
                                    {
                                        SizeF size = e.Graphics.MeasureString(dayNum, wFont);
                                        PointF pt = new PointF(
                                            e.CellBounds.Left + (e.CellBounds.Width - size.Width) / 2,
                                            e.CellBounds.Top + (e.CellBounds.Height - size.Height) / 2
                                        );
                                        using (SolidBrush wBrush = new SolidBrush(Color.FromArgb(30, 0, 0, 0))) // yarı saydam
                                        {
                                            e.Graphics.DrawString(dayNum, wFont, wBrush, pt);
                                        }
                                    }

                                    // Borderları ve içeriği (X, İ, R vb) normal çizmeye devam et (arka planı çizmeden)
                                    e.Paint(e.CellBounds, DataGridViewPaintParts.ContentForeground | DataGridViewPaintParts.Border);
                                    e.Handled = true;
                                }
                                else if (e.RowIndex >= 0 && (e.ColumnIndex == 0 || e.ColumnIndex == 1))
                                {
                                    // Sol taraftaki isim ve TC kısmı için satır seçiliyse vurgula
                                    bool isSelectedRow = (dgvPuantaj.CurrentCell != null && dgvPuantaj.CurrentCell.RowIndex == e.RowIndex);
                                    if (isSelectedRow)
                                    {
                                        e.CellStyle.BackColor = Color.FromArgb(220, 235, 255);
                                    }
                                    else
                                    {
                                        e.CellStyle.BackColor = Color.White;
                                    }
                                }
                            };

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
                        string gunAdi = gunTarihi.ToString("ddd", new System.Globalization.CultureInfo("tr-TR")).ToLower(new System.Globalization.CultureInfo("tr-TR"));
                        string dikeyGunAdi = string.Join("\n", gunAdi.ToCharArray());
                        string baslik = i + "\n" + dikeyGunAdi;

                        dgvPuantaj.Columns.Add("day" + i, baslik);
                        dgvPuantaj.Columns[i + 1].Width = 32;

                        // Gün başlıklarının arka planını hafif renklendir
                        dgvPuantaj.Columns[i + 1].HeaderCell.Style.BackColor = Color.FromArgb(238, 244, 255);
                        dgvPuantaj.Columns[i + 1].HeaderCell.Style.ForeColor = Color.FromArgb(50, 50, 50);

                        if (gunTarihi.DayOfWeek == DayOfWeek.Saturday || gunTarihi.DayOfWeek == DayOfWeek.Sunday)
                        {
                            dgvPuantaj.Columns[i + 1].HeaderCell.Style.BackColor = Color.FromArgb(255, 236, 236);
                            dgvPuantaj.Columns[i + 1].DefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
                        }
                    }
            }

            void VerileriYukle(DataGridView dgvPuantaj, string kampus)
            {
                try
                {
                    using (var conn = NotDbConnection.GetConnection())
                    {
                        string donem = secilenTarih.ToString("yyyy-MM");
                        string sql = @"
                            SELECT p.p_tc, p.p_ad_soyad,
                                   MAX(CASE WHEN p.p_yil_ay = @donem THEN p.p_gun_detaylari ELSE NULL END) AS p_gun_detaylari
                            FROM puantaj p
                            LEFT JOIN program_katilimcilari pk ON p.p_tc = pk.pk_tc
                            WHERE COALESCE(pk.pk_gorev_yeri, 'Diğer') = @kampus
                            GROUP BY p.p_tc, p.p_ad_soyad";

                        var cmd = new MySqlCommand(sql, conn);
                        cmd.Parameters.AddWithValue("@kampus", kampus);
                        cmd.Parameters.AddWithValue("@donem", donem);
                        using (var dr = cmd.ExecuteReader())
                        {
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

            // Server-side weekly limit check for a full month data array (1-based days)
            bool IsWeeklyLimitExceeded(List<string> gunVerileri)
            {
                if (gunVerileri == null) return false;
                int gunSayisi = DateTime.DaysInMonth(secilenTarih.Year, secilenTarih.Month);
                var weekCounts = new Dictionary<DateTime, int>();

                for (int i = 1; i <= Math.Min(gunSayisi, gunVerileri.Count); i++)
                {
                    DateTime currentDay = new DateTime(secilenTarih.Year, secilenTarih.Month, i);
                    int dow = (int)currentDay.DayOfWeek; // 0 = Sunday
                    int offset = dow == 0 ? 6 : dow - 1; // make Monday = 0
                    DateTime weekStart = currentDay.AddDays(-offset).Date;

                    string val = gunVerileri[i - 1];
                    if (string.IsNullOrEmpty(val) || val == "0") continue;
                    if (val == "X")
                    {
                        if (!weekCounts.ContainsKey(weekStart)) weekCounts[weekStart] = 0;
                        weekCounts[weekStart]++;
                        if (weekCounts[weekStart] > 3) return true;
                    }
                }

                return false;
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
                    using (var conn = NotDbConnection.GetConnection())
                    {
                        // Çoklu ay desteği için puantaj tablosundaki tüm UNIQUE kısıtlarını kaldır
                        try
                        {
                            var dtIdx = new DataTable();
                            using (var da = new MySqlDataAdapter("SELECT INDEX_NAME FROM INFORMATION_SCHEMA.STATISTICS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'puantaj' AND NON_UNIQUE = 0 AND INDEX_NAME <> 'PRIMARY' GROUP BY INDEX_NAME", conn))
                                da.Fill(dtIdx);
                            foreach (DataRow dr in dtIdx.Rows)
                            {
                                try { new MySqlCommand($"ALTER TABLE puantaj DROP INDEX `{dr["INDEX_NAME"]}`", conn).ExecuteNonQuery(); } catch { }
                            }
                        } catch { }

                        using (var transaction = conn.BeginTransaction())
                        {
                            try
                            {
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

                                        // Server-side weekly validation
                                        if (IsWeeklyLimitExceeded(gunVerileri))
                                        {
                                            SimpleLogger.Log($"Haftalık limit aşıldı. Kaydetme atlandı: TC={tc}, Donem={secilenTarih:yyyy-MM}");
                                            continue; // skip saving this row
                                        }

                                        string detayString = string.Join("-", gunVerileri);
                                        string donem = secilenTarih.ToString("yyyy-MM");

                                        // Seçili ay için kayıt var mı kontrol et
                                        string checkSql = "SELECT COUNT(*) FROM puantaj WHERE p_tc = @tc AND p_yil_ay = @donem";
                                        int mevcutKayit;
                                        using (var checkCmd = new MySqlCommand(checkSql, conn, transaction))
                                        {
                                            checkCmd.Parameters.AddWithValue("@tc", tc);
                                            checkCmd.Parameters.AddWithValue("@donem", donem);
                                            mevcutKayit = Convert.ToInt32(checkCmd.ExecuteScalar());
                                        }

                                        if (mevcutKayit > 0)
                                        {
                                            string sqlUpdate = "UPDATE puantaj SET p_gun_detaylari = @detay, p_calistigi_gun_sayisi = @toplam WHERE p_tc = @tc AND p_yil_ay = @donem";
                                            using (var cmd = new MySqlCommand(sqlUpdate, conn, transaction))
                                            {
                                                cmd.Parameters.AddWithValue("@detay", detayString);
                                                cmd.Parameters.AddWithValue("@toplam", toplamCalisilanGun);
                                                cmd.Parameters.AddWithValue("@tc", tc);
                                                cmd.Parameters.AddWithValue("@donem", donem);
                                                cmd.ExecuteNonQuery();
                                            }
                                        }
                                        else
                                        {
                                            string sqlInsert = @"INSERT INTO puantaj (p_tc, p_ad_soyad, p_iban, p_gun_detaylari, p_calistigi_gun_sayisi, p_yil_ay,
                                                                   p_ise_baslama_tarihi, p_isten_ayrilma_tarihi, p_devamsizlik, p_yillik_izin, p_ad, p_soyad)
                                                                 SELECT p_tc, p_ad_soyad, p_iban, @detay, @toplam, @donem,
                                                                        p_ise_baslama_tarihi, p_isten_ayrilma_tarihi, p_devamsizlik, p_yillik_izin, p_ad, p_soyad
                                                                 FROM puantaj WHERE p_tc = @tc LIMIT 1";
                                            using (var cmd = new MySqlCommand(sqlInsert, conn, transaction))
                                            {
                                                cmd.Parameters.AddWithValue("@detay", detayString);
                                                cmd.Parameters.AddWithValue("@toplam", toplamCalisilanGun);
                                                cmd.Parameters.AddWithValue("@donem", donem);
                                                cmd.Parameters.AddWithValue("@tc", tc);
                                                cmd.ExecuteNonQuery();
                                            }
                                        }
                                    }
                                }

                                transaction.Commit();
                            }
                            catch (Exception txEx)
                            {
                                try { transaction.Rollback(); } catch { }
                                SimpleLogger.Log("Kaydetme işlemi hata: " + txEx);
                                throw;
                            }
                        }
                    }

                    MessageBox.Show("✅ Tüm puantajlar başarıyla kaydedildi!");
                }
                catch (Exception ex)
                {
                    SimpleLogger.Log("Kaydetme genel hata: " + ex);
                    MessageBox.Show("Hata: " + ex.Message);
                }
            };

            btnExcelExport.Click += (s, e) =>
            {
                try
                {
                    // Şablon dosyasını proje debug dizininde "Resmi_Sablon.xlsx" olarak arayın; yoksa kullanıcıdan isteyin
                    string defaultTemplate1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resmi_Sablon.xlsx");
                    string defaultTemplate2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resmi_Sablon.xls");
                    string templatePath = null;
                    if (File.Exists(defaultTemplate1)) templatePath = defaultTemplate1;
                    else if (File.Exists(defaultTemplate2)) templatePath = defaultTemplate2;
                    else
                    {
                        // Kullanıcıdan şablon dosyasını seçmesini iste
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
                    // Kaydetme yeri seç
                    string ilkKampus = kampusGrids.Keys.FirstOrDefault() ?? "Kampüs1";
                    Dictionary<string, string> kampusNumaralari = new Dictionary<string, string>
                    {
                        { "Kampüs1", "13376" },
                        { "Kampüs2", "13377" },
                        { "Kampüs3", "13378" }
                    };
                    string dosyaKampusNo = kampusNumaralari.ContainsKey(ilkKampus) ? kampusNumaralari[ilkKampus] : "13376";
                    string dosyaAyAdi = dtpDonem.Value.ToString("MMMM", new CultureInfo("tr-TR"));
                    string dosyaAdi = $"{dosyaKampusNo} BUÜ BORDRO VE PUANTAJ {dosyaAyAdi} {dtpDonem.Value.Year}";

                    SaveFileDialog sfd = new SaveFileDialog
                    {
                        Filter = "Excel Dosyası|*.xlsx",
                        FileName = dosyaAdi + ".xlsx"
                    };
                    if (sfd.ShowDialog() != DialogResult.OK) return;

                    string savePath = sfd.FileName;

                    // Kopyala ve arkaplanda doldur
                    File.Copy(templatePath, savePath, true);

                    int donemYear = dtpDonem.Value.Year;
                    int donemMonth = dtpDonem.Value.Month;

                    Task.Run(() =>
                    {
                        try
                        {
                            using (var workbook = new XLWorkbook(savePath))
                            {
                                // Şablondaki "Puantaj" sayfasını 3 kampüs için kopyala ve her birini ayrı doldur
                                if (workbook.Worksheets.Any(w => w.Name.Equals("PUANTAJ", StringComparison.OrdinalIgnoreCase)))
                                {
                                    var templateWs = workbook.Worksheets.First(w => w.Name.Equals("PUANTAJ", StringComparison.OrdinalIgnoreCase));
                                    string origSheetName = templateWs.Name;
                                    var kampusListesi = kampusGrids.Keys.ToArray();
                                    var kampusSheets = new List<IXLWorksheet>();
                                    foreach (var kName in kampusListesi)
                                    {
                                        kampusSheets.Add(templateWs.CopyTo(kName));
                                    }
                                    workbook.Worksheet(origSheetName).Visibility = XLWorksheetVisibility.VeryHidden;
                                    for (int i = 0; i < kampusSheets.Count; i++)
                                        kampusSheets[i].Position = i + 1;

                                    foreach (var ws in kampusSheets)
                                    {
                                        string currentKampus = ws.Name;

                                    // Bul header satırını ve ilgili sütun indekslerini tespit et
                                    int headerRow = -1;
                                    int maxSearchRow = 20;
                                    int maxSearchCol = 40;
                                    int tcCol = -1, adCol = -1, soyadCol = -1, ibanCol = -1;

                                    for (int r = 1; r <= maxSearchRow; r++)
                                    {
                                        for (int c = 1; c <= maxSearchCol; c++)
                                        {
                                            string hv = ws.Cell(r, c).GetString()?.Trim();
                                            if (string.IsNullOrEmpty(hv)) continue;
                                            string hvUp = hv.ToUpperInvariant();
                                            // More flexible matching: accept variants like 'TC KIMLIK', 'T.C.', 'TC'
                                            if ((hvUp.Contains("TC") || hvUp.Contains("T.C") || hvUp.Contains("TCKIMLIK") || hvUp.Contains("TC KIMLIK") || hvUp.Contains("TC KİMLİK")) && headerRow == -1)
                                            {
                                                headerRow = r; tcCol = c;
                                            }
                                            if ((hvUp.Contains("AD SOYAD") || hvUp.Contains("AD") && hvUp.Contains("SOYAD")) && adCol == -1)
                                            {
                                                adCol = c;
                                            }
                                            else if (hvUp.Contains("SOYAD") && soyadCol == -1) { soyadCol = c; }
                                            else if (hvUp.Contains("IBAN") && ibanCol == -1) { ibanCol = c; }
                                            else if (hvUp.Contains("SIRA") && hvUp.Length < 10) { /* possible SIRA column, ignore for headerRow detection */ }
                                        }
                                        if (headerRow != -1 && (adCol != -1 || soyadCol != -1)) break;
                                    }

                                    // Log header row detection for debugging
                                    try
                                    {
                                        var dbgHdr = new System.Text.StringBuilder();
                                        dbgHdr.AppendLine($"Detected headerRow={headerRow}, tcCol={tcCol}, adCol={adCol}, soyadCol={soyadCol}, ibanCol={ibanCol}");
                                        for (int c = 1; c <= Math.Min(maxSearchCol, 60); c++)
                                        {
                                            var cellv = ws.Cell(headerRow, c).GetString();
                                            if (!string.IsNullOrEmpty(cellv)) dbgHdr.AppendLine($"col{c}: '{cellv}'");
                                        }
                                        SimpleLogger.Log(dbgHdr.ToString());
                                    }
                                    catch { }

                                    if (headerRow == -1 || tcCol == -1)
                                    {
                                        // Eğer beklenen başlıklar yoksa kullanıcıyı bilgilendir
                                        SimpleLogger.Log("PUANTAJ sayfasında 'TC' başlığı bulunamadı. Şablon uyumsuz.");
                                    }
                                    else
                                    {
                                        // Önce şablondaki başlık penceresini tara ve eğer '#YOK' yer tutucuları varsa bunları
                                        // seçilen dönemin gün başlıklarıyla değiştir. Bunu, veri satırlarını temizlemeden önce yapıyoruz
                                        // çünkü bazı formüller veri bölgesine referans verebilir.
                                        int gunSayisi_p = DateTime.DaysInMonth(donemYear, donemMonth);
                                        int headerSearchTop_p = Math.Max(1, headerRow - 2);
                                        int headerSearchBottom_p = headerRow + 2;
                                        int dayStartPlaceholderCol_p = -1;
                                        int dayStartPlaceholderRow_p = -1;
                                        for (int r = headerSearchTop_p; r <= headerSearchBottom_p; r++)
                                        {
                                            for (int c = 1; c <= maxSearchCol; c++)
                                            {
                                                var hvTest = ws.Cell(r, c).GetString();
                                                if (string.IsNullOrEmpty(hvTest)) continue;
                                                if (hvTest.Trim().Equals("#YOK", StringComparison.OrdinalIgnoreCase) || hvTest.ToUpperInvariant().Contains("YOK"))
                                                {
                                                    dayStartPlaceholderCol_p = c;
                                                    dayStartPlaceholderRow_p = r;
                                                    break;
                                                }
                                            }
                                            if (dayStartPlaceholderCol_p > 0) break;
                                        }
                                        if (dayStartPlaceholderCol_p > 0)
                                        {
                                        for (int d = 1; d <= gunSayisi_p; d++)
                                        {
                                            DateTime dayDate = new DateTime(donemYear, donemMonth, d);
                                            // Use full day name in Turkish and convert to lowercase (e.g. "cumartesi")
                                            string shortName = dayDate.ToString("dddd", new System.Globalization.CultureInfo("tr-TR")).ToLower(new System.Globalization.CultureInfo("tr-TR"));
                                            ws.Cell(dayStartPlaceholderRow_p, dayStartPlaceholderCol_p + d - 1).Value = $"{d} {shortName}";
                                        }
                                            SimpleLogger.Log($"(Pre-clear) Replaced #YOK starting at row {dayStartPlaceholderRow_p}, col {dayStartPlaceholderCol_p}");
                                        }

                                        int writeRow = headerRow + 1;
                                        int weekdayRow = -1;
                                        // Detect if the row after headers is a WEEKDAY helper row (used by conditional formatting)
                                        try
                                        {
                                            for (int c = 1; c <= Math.Min(maxSearchCol, 20); c++)
                                            {
                                                var cell = ws.Cell(writeRow, c);
                                                if (cell.HasFormula)
                                                {
                                                    string f = (cell.FormulaA1 ?? "").ToUpperInvariant();
                                                    if (f.Contains("WEEKDAY"))
                                                    {
                                                        weekdayRow = writeRow;
                                                        writeRow++;
                                                        SimpleLogger.Log($"Detected WEEKDAY helper row at {weekdayRow}, data starts at {writeRow}");
                                                        break;
                                                    }
                                                }
                                            }
                                            // Fallback: check for small integers (1-7) with empty TC cell
                                            if (weekdayRow == -1)
                                            {
                                                bool tcEmpty = string.IsNullOrEmpty(ws.Cell(writeRow, tcCol).GetString()?.Trim());
                                                if (tcEmpty)
                                                {
                                                    for (int c = 5; c <= Math.Min(15, maxSearchCol); c++)
                                                    {
                                                        var v = ws.Cell(writeRow, c).GetString()?.Trim();
                                                        if (string.IsNullOrEmpty(v)) continue;
                                                        int num;
                                                        if (int.TryParse(v, out num) && num >= 1 && num <= 7)
                                                        {
                                                            weekdayRow = writeRow;
                                                            writeRow++;
                                                            SimpleLogger.Log($"Detected WEEKDAY helper row (fallback) at {weekdayRow}, data starts at {writeRow}");
                                                        }
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                        catch { }

                                        // Detect template cells for Year/Month display and hire/exit date columns
                                        int yilLabelRow = -1, yilLabelCol = -1, ayLabelRow = -1, ayLabelCol = -1;
                                        int iseBaslamaCol = -1, iseAyrilmaCol = -1;
                                        try
                                        {
                                            // First pass: look for explicit labels 'YIL' and 'AY' in a reasonable header area
                                            for (int r = 1; r <= Math.Min(15, maxSearchRow); r++)
                                            {
                                                for (int c = 1; c <= Math.Min(80, maxSearchCol); c++)
                                                {
                                                    var hv = ws.Cell(r, c).GetString();
                                                    if (string.IsNullOrWhiteSpace(hv)) continue;
                                                    var hvUp = hv.ToUpperInvariant().Trim();
                                                    if (hvUp.Contains("YIL") && yilLabelRow == -1) { yilLabelRow = r; yilLabelCol = c; }
                                                    if (hvUp.Contains(" AY") || hvUp == "AY" || hvUp.Contains("AY ") && ayLabelRow == -1) { ayLabelRow = r; ayLabelCol = c; }

                                                    // hire/exit headers detection (various variants)
                                                    if ((hvUp.Contains("İŞE GİRİŞ") || hvUp.Contains("ISE GIRIS") || hvUp.Contains("İŞE GİRİŞ TARİH") || hvUp.Contains("GİRİŞ TARİH")) && iseBaslamaCol <= 0)
                                                    {
                                                        iseBaslamaCol = c;
                                                    }
                                                    if ((hvUp.Contains("İŞTEN ÇIKIŞ") || hvUp.Contains("İŞTEN AYRILMA") || hvUp.Contains("AYRILMA TARİH") || hvUp.Contains("ÇIKIŞ TARİH")) && iseAyrilmaCol <= 0)
                                                    {
                                                        iseAyrilmaCol = c;
                                                    }
                                                }
                                            }

                                            // If explicit labels not found, try to find an existing year cell (e.g. '2023') in top area and use its column
                                            if (yilLabelRow == -1)
                                            {
                                                for (int r = 1; r <= Math.Min(15, maxSearchRow) && yilLabelRow == -1; r++)
                                                {
                                                    for (int c = 1; c <= Math.Min(80, maxSearchCol); c++)
                                                    {
                                                        var hv = ws.Cell(r, c).GetString();
                                                        if (string.IsNullOrWhiteSpace(hv)) continue;
                                                        // look for 4-digit year
                                                        if (hv.Trim().Length == 4 && hv.Trim().All(char.IsDigit) && hv.Trim().StartsWith("20"))
                                                        {
                                                            // prefer to treat this as year value cell; assume label is left of it if exists
                                                            yilLabelRow = r;
                                                            yilLabelCol = c - 1 > 0 ? c - 1 : c;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }

                                            // Write year/month values into adjacent cells if labels found or year cell located
                                            if (yilLabelRow > 0 && yilLabelCol > 0)
                                            {
                                                // prefer writing into the cell that currently contains a year or the cell right to label
                                                var candidateCell = ws.Cell(yilLabelRow, yilLabelCol + 1);
                                                try
                                                {
                                                    // if candidate contains a 4-digit year or is empty, write year there
                                                    var cur = candidateCell.GetString();
                                                    if (string.IsNullOrWhiteSpace(cur) || (cur.Trim().Length == 4 && cur.Trim().All(char.IsDigit)))
                                                        candidateCell.Value = donemYear;
                                                    else
                                                    {
                                                        // fallback: write into label cell's right anyway
                                                        candidateCell.Value = donemYear;
                                                    }
                                                }
                                                catch { try { ws.Cell(yilLabelRow, yilLabelCol + 1).Value = donemYear; } catch { } }
                                            }
                                            if (ayLabelRow > 0 && ayLabelCol > 0)
                                            {
                                                var candidateCell = ws.Cell(ayLabelRow, ayLabelCol + 1);
                                                try
                                                {
                                                    var cur = candidateCell.GetString();
                                                    if (string.IsNullOrWhiteSpace(cur) || cur.Trim().Any())
                                                        candidateCell.Value = dtpDonem.Value.ToString("MMMM", new CultureInfo("tr-TR"));
                                                }
                                                catch { try { ws.Cell(ayLabelRow, ayLabelCol + 1).Value = dtpDonem.Value.ToString("MMMM", new CultureInfo("tr-TR")); } catch { } }
                                            }
                                        }
                                        catch { }

                                        // Temizle: mevcut veri satırlarını kaldır (headerRow altındaki tüm dolu satırları)
                                        int lastRow = ws.LastRowUsed()?.RowNumber() ?? writeRow - 1;
                                        if (lastRow >= writeRow)
                                        {
                                            ws.Rows(writeRow, lastRow).Clear(XLClearOptions.Contents);
                                        }

                                        // Fix: CopyTo may break date values and WEEKDAY formulas on copied sheets.
                                        // Pre-write fresh date serials and WEEKDAY values so day column detection works.
                                        int preDayCol = Math.Max(Math.Max(tcCol, adCol), Math.Max(soyadCol > 0 ? soyadCol : 0, ibanCol)) + 1;
                                        if (preDayCol < 6) preDayCol = 6;
                                        int preDateRow = headerRow > 1 ? headerRow - 1 : -1;
                                        int preGunSayisi = DateTime.DaysInMonth(donemYear, donemMonth);
                                        if (preDateRow > 0)
                                        {
                                            try
                                            {
                                                for (int d = 1; d <= 31; d++)
                                                {
                                                    int col = preDayCol + d - 1;
                                                    if (d <= preGunSayisi)
                                                    {
                                                        DateTime dayDate = new DateTime(donemYear, donemMonth, d);
                                                        ws.Cell(preDateRow, col).Value = dayDate.ToOADate();
                                                    }
                                                    else
                                                    {
                                                        ws.Cell(preDateRow, col).Value = "";
                                                    }
                                                }
                                                SimpleLogger.Log($"Pre-wrote date serials to row {preDateRow}, startCol={preDayCol}");
                                            }
                                            catch { }
                                        }
                                        if (weekdayRow > 0 && preDayCol > 0)
                                        {
                                            try
                                            {
                                                for (int d = 1; d <= 31; d++)
                                                {
                                                    int col = preDayCol + d - 1;
                                                    if (d <= preGunSayisi)
                                                    {
                                                        DateTime dayDate = new DateTime(donemYear, donemMonth, d);
                                                        int dow = (int)dayDate.DayOfWeek;
                                                        int weekdayVal = dow == 0 ? 7 : dow;
                                                        ws.Cell(weekdayRow, col).Value = weekdayVal;
                                                    }
                                                    else
                                                    {
                                                        ws.Cell(weekdayRow, col).Value = "";
                                                    }
                                                }
                                                SimpleLogger.Log($"Pre-wrote WEEKDAY values to row {weekdayRow}, startCol={preDayCol}");
                                            }
                                            catch { }
                                        }

                                        // DB'den kampüse ait katılımcıları çek ve sayfaya sırayla yaz
                                        using (var conn = NotDbConnection.GetConnection())
                                        {
                                            string sqlAll = @"SELECT p.p_iban, p.p_tc, p.p_ad_soyad, COALESCE(p.p_ad, '') AS p_ad, COALESCE(p.p_soyad, '') AS p_soyad, p.p_gun_detaylari, p.p_ise_baslama_tarihi AS p_ise_baslama_tarihi, p.p_isten_ayrilma_tarihi AS p_isten_ayrilma_tarihi,
                                                              p.p_calistigi_gun_sayisi, p.p_devamsizlik, p.p_yillik_izin,
                                                              pk.pk_is_baslama_tarihi AS pk_ise_baslama, pk.pk_isten_ayrilma_tarihi AS pk_isten_ayrilma
                                                              FROM puantaj p
                                                              LEFT JOIN program_katilimcilari pk ON p.p_tc = pk.pk_tc
                                                              WHERE COALESCE(pk.pk_gorev_yeri, 'Diğer') = @kampus
                                                                AND p.p_yil_ay = @donem
                                                              ORDER BY COALESCE(p.p_ad, p.p_ad_soyad)";
                                            using (var cmdAll = new MySqlCommand(sqlAll, conn))
                                            {
                                            cmdAll.Parameters.AddWithValue("@kampus", currentKampus);
                                            cmdAll.Parameters.AddWithValue("@donem", $"{donemYear:D4}-{donemMonth:D2}");
                                            using (var drAll = cmdAll.ExecuteReader())
                                            {
                                                int siraNo = 1;
                                                // If template uses placeholders like "#YOK" for day headers, replace them with actual headers for the selected month
                                                int gunSayisi = DateTime.DaysInMonth(donemYear, donemMonth);
                                                int dayStartPlaceholderCol = -1;
                                                int dayStartPlaceholderRow = -1;
                                                // search a small header window around headerRow because template may place day placeholders on adjacent row
                                                int headerSearchTop = Math.Max(1, headerRow - 2);
                                                int headerSearchBottom = headerRow + 2;

                                                // Log header window cells to help debugging
                                                try
                                                {
                                                    var dbgWin = new System.Text.StringBuilder();
                                                    dbgWin.AppendLine($"Scanning header window rows {headerSearchTop}..{headerSearchBottom}");
                                                    for (int r = headerSearchTop; r <= headerSearchBottom; r++)
                                                    {
                                                        for (int c = 1; c <= maxSearchCol; c++)
                                                        {
                                                            var cellv = ws.Cell(r, c).GetString();
                                                            if (!string.IsNullOrEmpty(cellv)) dbgWin.AppendLine($"r{r}c{c}: '{cellv}'");
                                                        }
                                                    }
                                                    SimpleLogger.Log(dbgWin.ToString());
                                                }
                                                catch { }

                                                for (int r = headerSearchTop; r <= headerSearchBottom; r++)
                                                {
                                                    for (int c = 1; c <= maxSearchCol; c++)
                                                    {
                                                        var hvTest = ws.Cell(r, c).GetString();
                                                        if (string.IsNullOrEmpty(hvTest)) continue;
                                                        // check for '#YOK' or cell containing 'YOK' (trim and ignore case)
                                                        if (hvTest.Trim().Equals("#YOK", StringComparison.OrdinalIgnoreCase) || hvTest.ToUpperInvariant().Contains("YOK"))
                                                        {
                                                            dayStartPlaceholderCol = c;
                                                            dayStartPlaceholderRow = r;
                                                            break;
                                                        }
                                                    }
                                                    if (dayStartPlaceholderCol > 0) break;
                                                }

                                                if (dayStartPlaceholderCol > 0)
                                                {
                                                        for (int d = 1; d <= gunSayisi; d++)
                                                        {
                                                            DateTime dayDate = new DateTime(donemYear, donemMonth, d);
                                                            string shortName = dayDate.ToString("dddd", new System.Globalization.CultureInfo("tr-TR")).ToLower(new System.Globalization.CultureInfo("tr-TR"));
                                                            ws.Cell(dayStartPlaceholderRow, dayStartPlaceholderCol + d - 1).Value = $"{d} {shortName}";
                                                        }
                                                    SimpleLogger.Log($"Replaced #YOK starting at row {dayStartPlaceholderRow}, col {dayStartPlaceholderCol}");
                                                }

                                                // Detect day columns based on header rows in the header window
                                                var dayColByNumber = new Dictionary<int, int>();
                                                int headerRowForDays = -1;
                                                for (int r = headerSearchTop; r <= headerSearchBottom; r++)
                                                {
                                                    if (r == headerRow || r == weekdayRow) continue; // Skip header and WEEKDAY rows
                                                    for (int c = 1; c <= maxSearchCol; c++)
                                                    {
                                                        try
                                                        {
                                                            var hv = ws.Cell(r, c).GetString()?.Trim();
                                                            if (string.IsNullOrEmpty(hv)) continue;
                                                            var token = hv.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0];
                                                            int dayNum;
                                                            if (int.TryParse(token, out dayNum) && dayNum >= 1 && dayNum <= 31)
                                                            {
                                                                if (!dayColByNumber.ContainsKey(dayNum))
                                                                {
                                                                    dayColByNumber[dayNum] = c;
                                                                    if (headerRowForDays == -1) headerRowForDays = r;
                                                                }
                                                            }
                                                        }
                                                        catch
                                                        {
                                                            // ignore parse issues
                                                        }
                                                    }
                                                }

                                                // If still no explicit day headers, try fallback: find first header cell with any digit in the header window
                                                if (dayColByNumber.Count == 0)
                                                {
                                                    int firstDayCol = -1;
                                                    int firstDayRow = -1;
                                                    for (int r = headerSearchTop; r <= headerSearchBottom; r++)
                                                    {
                                                        if (r == headerRow || r == weekdayRow) continue; // Skip header and WEEKDAY rows
                                                        for (int c = 1; c <= maxSearchCol; c++)
                                                        {
                                                            var hv = ws.Cell(r, c).GetString();
                                                            if (string.IsNullOrEmpty(hv)) continue;
                                                            if (hv.Any(char.IsDigit)) { firstDayCol = c; firstDayRow = r; break; }
                                                        }
                                                        if (firstDayCol > 0) break;
                                                    }
                                                    if (firstDayCol > 0)
                                                    {
                                                        for (int d = 1; d <= gunSayisi; d++) dayColByNumber[d] = firstDayCol + d - 1;
                                                        SimpleLogger.Log($"Fallback day start at row {firstDayRow}, col {firstDayCol}");
                                                        if (headerRowForDays == -1) headerRowForDays = firstDayRow;
                                                    }
                                                }

                                                // If template has 'AD SOYAD' combined header (merged D:E) but no separate 'Soyad' column,
                                                // use the next column for soyad data. Don't insert columns to preserve template layout.
                                                if (adCol > 0 && soyadCol <= 0)
                                                {
                                                    soyadCol = adCol + 1;
                                                    SimpleLogger.Log($"Using column {soyadCol} for Soyad (adjacent to AD SOYAD at {adCol})");
                                                }

                                                // Log detection result to help debug when nothing is written
                                                try
                                                {
                                                    var dbg = new System.Text.StringBuilder();
                                                    dbg.AppendLine($"PUANTAJ detection: headerRow={headerRow}, tcCol={tcCol}, adCol={adCol}, soyadCol={soyadCol}, ibanCol={ibanCol}");
                                                    dbg.AppendLine($"Detected day columns count: {dayColByNumber.Count}");
                                                    if (dayColByNumber.Count > 0)
                                                    {
                                                        foreach (var kv in dayColByNumber.OrderBy(k => k.Key)) dbg.AppendLine($"day{kv.Key} -> col{kv.Value}");
                                                    }
                                                    SimpleLogger.Log(dbg.ToString());
                                                    // If no day columns found, show message to user to help fix template
                                                    if (dayColByNumber.Count == 0)
                                                    {
                                                        this.Invoke(new Action(() =>
                                                        {
                                                            MessageBox.Show("PUANTAJ sayfasında gün sütunları tespit edilemedi. Lütfen şablon başlıklarını kontrol edin (ör. '1', '2', ... veya '1 Pazartesi'). Log kaydına bakın.", "Şablon Uyarısı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                                        }));
                                                    }
                                                }
                                                catch { }

                                                // Update date values in the date header row for the selected month.
                                                // Write date serial numbers (not text) to preserve the template's date format and conditional formatting.
                                                // The template's style already formats dates as "1 cumartesi" etc.
                                                // WEEKDAY formulas in the helper row reference these dates and will recalculate in Excel.
                                                try
                                                {
                                                    int dateRow = -1;
                                                    if (headerRowForDays > 0 && headerRowForDays < headerRow) dateRow = headerRowForDays;
                                                    else if (headerRow > 1) dateRow = headerRow - 1;

                                                    if (dateRow > 0 && dayColByNumber.Count > 0)
                                                    {
                                                        int firstDayCol = dayColByNumber.Values.Min();
                                                        // Write date serial values for each day; clear extra columns for shorter months
                                                        for (int d = 1; d <= 31; d++)
                                                        {
                                                            int col = firstDayCol + d - 1;
                                                            if (d <= gunSayisi)
                                                            {
                                                                DateTime dayDate = new DateTime(donemYear, donemMonth, d);
                                                                ws.Cell(dateRow, col).Value = dayDate.ToOADate();
                                                            }
                                                            else
                                                            {
                                                                ws.Cell(dateRow, col).Value = "";
                                                            }
                                                        }
                                                        // Ensure dayColByNumber covers all days of the selected month
                                                        for (int d = 1; d <= gunSayisi; d++)
                                                        {
                                                            if (!dayColByNumber.ContainsKey(d))
                                                                dayColByNumber[d] = firstDayCol + d - 1;
                                                        }
                                                        SimpleLogger.Log($"Wrote date serial values to row {dateRow} for {donemMonth}/{donemYear}, firstDayCol={firstDayCol}");
                                                    }
                                                }
                                                catch { }

                                                while (drAll.Read())
                                                {
                                                    string tc = drAll["p_tc"]?.ToString() ?? "";
                                                    if (string.IsNullOrEmpty(tc)) continue;

                                                    if (ibanCol > 0)
                                                    {
                                                        var c = ws.Cell(writeRow, ibanCol);
                                                        c.Value = drAll["p_iban"]?.ToString() ?? "";
                                                        try { c.Style.Font.FontName = "Segoe UI"; c.Style.Font.FontSize = 9; c.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; c.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; c.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; c.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; } catch { }
                                                    }
                                                    if (tcCol > 0)
                                                    {
                                                        var c = ws.Cell(writeRow, tcCol);
                                                        c.Value = tc;
                                                        try { c.Style.Font.FontName = "Segoe UI"; c.Style.Font.FontSize = 9; c.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; c.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; c.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; c.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; } catch { }
                                                    }

                                                    // DB'den ad ve soyad ayrı sütunlardan oku
                                                    string dbAd = drAll["p_ad"]?.ToString() ?? "";
                                                    string dbSoyad = drAll["p_soyad"]?.ToString() ?? "";

                                                    // Eğer p_ad boşsa eski p_ad_soyad'dan split et (geriye uyum)
                                                    if (string.IsNullOrEmpty(dbAd))
                                                    {
                                                        string adSoyad = drAll["p_ad_soyad"]?.ToString() ?? "";
                                                        var parts = adSoyad.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                                        if (parts.Length > 0) dbAd = parts[0];
                                                        if (parts.Length > 1) dbSoyad = string.Join(" ", parts.Skip(1));
                                                    }

                                                    if (adCol > 0 && soyadCol > 0)
                                                    {
                                                        var cad = ws.Cell(writeRow, adCol);
                                                        var csoy = ws.Cell(writeRow, soyadCol);
                                                        cad.Value = dbAd;
                                                        csoy.Value = dbSoyad;
                                                        try { cad.Style.Font.FontName = "Segoe UI"; cad.Style.Font.FontSize = 9; cad.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; cad.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; cad.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; cad.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; } catch { }
                                                        try { csoy.Style.Font.FontName = "Segoe UI"; csoy.Style.Font.FontSize = 9; csoy.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; csoy.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; csoy.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; csoy.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; } catch { }
                                                    }
                                                    else if (adCol > 0)
                                                    {
                                                        var cad = ws.Cell(writeRow, adCol);
                                                        cad.Value = dbAd;
                                                        try { cad.Style.Font.FontName = "Segoe UI"; cad.Style.Font.FontSize = 9; cad.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; cad.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; cad.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; cad.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; } catch { }
                                                        try
                                                        {
                                                            var cnext = ws.Cell(writeRow, adCol + 1);
                                                            cnext.Value = dbSoyad;
                                                            try { cnext.Style.Font.FontName = "Segoe UI"; cnext.Style.Font.FontSize = 9; cnext.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; cnext.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; cnext.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; cnext.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; } catch { }
                                                        }
                                                        catch { }
                                                    }

                                                    // Gün detaylarını yaz (p_gun_detaylari -> '-' ile ayrılmış)
                                                    try
                                                    {
                                                        string gunDetaylari = drAll["p_gun_detaylari"]?.ToString() ?? "";
                                                        if (!string.IsNullOrEmpty(gunDetaylari))
                                                        {
                                                            string[] gunler = gunDetaylari.Split('-');
                                                            for (int gun = 1; gun <= gunler.Length; gun++)
                                                            {
                                                                string gunDegeri = gunler[gun - 1];
                                                                if (gunDegeri == "0" || string.IsNullOrEmpty(gunDegeri)) gunDegeri = "";

                                                                    int targetCol;
                                                                    if (dayColByNumber.TryGetValue(gun, out targetCol))
                                                                    {
                                                                        var dayCell = ws.Cell(writeRow, targetCol);
                                                                        dayCell.Value = gunDegeri;
                                                                        try { dayCell.Style.Font.FontName = "Segoe UI"; dayCell.Style.Font.FontSize = 9; dayCell.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center; dayCell.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; dayCell.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; dayCell.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; dayCell.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; } catch { }
                                                                    }
                                                            }
                                                        }
                                                    }
                                                    catch { }

                                                    // Diğer özet hücreler: çalışılan gün sayısını COUNTIF formülü ile yaz
                                                    if (writeRow > headerRow)
                                                    {
                                                        try
                                                        {
                                                            // Şablonda çalışılan gün/özet sütunu başlığı varsa tespit et (header window içinde ara)
                                                            int ozetCol = -1;
                                                            for (int r = headerSearchTop; r <= headerSearchBottom && ozetCol == -1; r++)
                                                            {
                                                                for (int c = 1; c <= maxSearchCol; c++)
                                                                {
                                                                    var hv = ws.Cell(r, c).GetString();
                                                                    if (string.IsNullOrEmpty(hv)) continue;
                                                                    var hvUp = hv.ToUpperInvariant();
                                                                    if (hvUp.Contains("ÇALIŞ") || hvUp.Contains("CALIS") || hvUp.Contains("TOPLAM") || hvUp.Contains("GÜN") || hvUp.Contains("GUN"))
                                                                    {
                                                                        ozetCol = c;
                                                                        break;
                                                                    }
                                                                }
                                                            }

                                                            if (ozetCol > 0 && dayColByNumber.Count > 0)
                                                            {
                                                                int minCol = dayColByNumber.Values.Min();
                                                                int maxCol = dayColByNumber.Values.Max();
                                                                string firstColLetter = XLHelper.GetColumnLetterFromNumber(minCol);
                                                                string lastColLetter = XLHelper.GetColumnLetterFromNumber(maxCol);
                                                                ws.Cell(writeRow, ozetCol).FormulaA1 = $"COUNTIF({firstColLetter}{writeRow}:{lastColLetter}{writeRow},\"X\")";
                                                            }
                                                        }
                                                        catch { }
                                                    }

                                                    // Write hire/exit dates per row if template columns detected
                                                    try
                                                    {
                                                        if (iseBaslamaCol > 0)
                                                        {
                                                            bool wrote = false;
                                                            try
                                                            {
                                                                int ordPk = drAll.GetOrdinal("pk_ise_baslama");
                                                                if (!drAll.IsDBNull(ordPk))
                                                                {
                                                                    DateTime iseBaslama = Convert.ToDateTime(drAll.GetValue(ordPk));
                                                                    ws.Cell(writeRow, iseBaslamaCol).Value = iseBaslama.ToString("dd.MM.yyyy");
                                                                    wrote = true;
                                                                }
                                                            }
                                                            catch (IndexOutOfRangeException) { }

                                                            if (!wrote)
                                                            {
                                                                try
                                                                {
                                                                    int ordP = drAll.GetOrdinal("p_ise_baslama_tarihi");
                                                                    if (!drAll.IsDBNull(ordP))
                                                                    {
                                                                        DateTime iseBaslama = Convert.ToDateTime(drAll.GetValue(ordP));
                                                                        ws.Cell(writeRow, iseBaslamaCol).Value = iseBaslama.ToString("dd.MM.yyyy");
                                                                    }
                                                                }
                                                                catch (IndexOutOfRangeException) { }
                                                            }
                                                        }

                                                        if (iseAyrilmaCol > 0)
                                                        {
                                                            bool wrote2 = false;
                                                            try
                                                            {
                                                                int ordPk2 = drAll.GetOrdinal("pk_isten_ayrilma");
                                                                if (!drAll.IsDBNull(ordPk2))
                                                                {
                                                                    DateTime iseAyrilma = Convert.ToDateTime(drAll.GetValue(ordPk2));
                                                                    ws.Cell(writeRow, iseAyrilmaCol).Value = iseAyrilma.ToString("dd.MM.yyyy");
                                                                    wrote2 = true;
                                                                }
                                                            }
                                                            catch (IndexOutOfRangeException) { }

                                                            if (!wrote2)
                                                            {
                                                                try
                                                                {
                                                                    int ordP2 = drAll.GetOrdinal("p_isten_ayrilma_tarihi");
                                                                    if (!drAll.IsDBNull(ordP2))
                                                                    {
                                                                        DateTime iseAyrilma = Convert.ToDateTime(drAll.GetValue(ordP2));
                                                                        ws.Cell(writeRow, iseAyrilmaCol).Value = iseAyrilma.ToString("dd.MM.yyyy");
                                                                    }
                                                                }
                                                                catch (IndexOutOfRangeException) { }
                                                            }
                                                        }
                                                    }
                                                    catch { }

                                                    siraNo++;
                                                    writeRow++;
                                                }
                                            }
                                            }
                                        }
                                    }
                                    }
                                }
                                else
                                {
                                    // Eğer PUANTAJ sayfası yoksa önceki davranış: her kampüs için ayrı sayfa doldur
                                    foreach (var kvp in kampusGrids)
                                    {
                                        string kampus = kvp.Key;
                                        string kampusNo = kampusNumaralari.ContainsKey(kampus) ? kampusNumaralari[kampus] : "13376";

                                        IXLWorksheet ws;
                                        if (workbook.Worksheets.Any(w => w.Name.Equals(kampus, StringComparison.OrdinalIgnoreCase)))
                                        {
                                            ws = workbook.Worksheet(kampus);
                                        }
                                        else
                                        {
                                            ws = workbook.AddWorksheet(kampus);
                                        }

                                        int gunSayisi = DateTime.DaysInMonth(donemYear, donemMonth);
                                        int gunKolonBaslangic = 6; // varsayılan
                                        int ozetKolonBaslangic = gunKolonBaslangic + gunSayisi;

                                        int dataStartRow = 6;
                                        for (int r = 1; r <= 10; r++)
                                        {
                                            for (int c = 1; c <= 10; c++)
                                            {
                                                var hv = ws.Cell(r, c).GetString();
                                                if (!string.IsNullOrWhiteSpace(hv) && hv.Trim().Equals("TC", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    dataStartRow = r + 1;
                                                    break;
                                                }
                                            }
                                            if (dataStartRow != 6) break;
                                        }

                                        using (var conn = NotDbConnection.GetConnection())
                                        {
                                            string sql = @"SELECT p.p_iban, p.p_tc, p.p_ad_soyad, COALESCE(p.p_ad, '') AS p_ad, COALESCE(p.p_soyad, '') AS p_soyad, p.p_gun_detaylari, p.p_ise_baslama_tarihi AS p_ise_baslama_tarihi, p.p_isten_ayrilma_tarihi AS p_isten_ayrilma_tarihi, 
                                                           p.p_calistigi_gun_sayisi, p.p_devamsizlik, p.p_yillik_izin,
                                                           pk.pk_is_baslama_tarihi AS pk_ise_baslama, pk.pk_isten_ayrilma_tarihi AS pk_isten_ayrilma
                                                           FROM puantaj p
                                                           LEFT JOIN program_katilimcilari pk ON p.p_tc = pk.pk_tc
                                                           WHERE COALESCE(pk.pk_gorev_yeri, 'Diğer') = @kampus";
                                            using (var cmd = new MySqlCommand(sql, conn))
                                            {
                                                cmd.Parameters.AddWithValue("@kampus", kampus);
                                                using (var dr = cmd.ExecuteReader())
                                                {
                                                    int siraNo = 1;
                                                     int writeRow = dataStartRow;

                                                     // Gün başlık hücrelerini yaz ve hafta sonu renklendir
                                                     int headerRowKampus = dataStartRow - 1;
                                                     var weekendColorKampus = ClosedXML.Excel.XLColor.FromArgb(198, 224, 255);
                                                     for (int d = 1; d <= gunSayisi; d++)
                                                     {
                                                         DateTime dayDate = new DateTime(donemYear, donemMonth, d);
                                                         string dayName = dayDate.ToString("dddd", new System.Globalization.CultureInfo("tr-TR")).ToLower(new System.Globalization.CultureInfo("tr-TR"));
                                                         var hdrCell = ws.Cell(headerRowKampus, gunKolonBaslangic + d - 1);
                                                         hdrCell.Value = $"{d} {dayName}";
                                                         hdrCell.Style.NumberFormat.Format = "@";
                                                         hdrCell.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                                                         try { hdrCell.Style.Font.Bold = true; hdrCell.Style.Font.FontName = "Segoe UI"; hdrCell.Style.Font.FontSize = 9; } catch { }
                                                         try { hdrCell.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; hdrCell.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; hdrCell.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; hdrCell.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin; } catch { }
                                                         if (dayDate.DayOfWeek == DayOfWeek.Saturday || dayDate.DayOfWeek == DayOfWeek.Sunday)
                                                         {
                                                             hdrCell.Style.Fill.BackgroundColor = weekendColorKampus;
                                                         }
                                                     }

                                                     while (dr.Read())
                                                    {
                                                        string tc = dr["p_tc"]?.ToString() ?? "";
                                                        if (string.IsNullOrEmpty(tc)) continue;

                                                        ws.Cell(writeRow, 1).Value = dr["p_iban"]?.ToString() ?? "";
                                                        ws.Cell(writeRow, 2).Value = siraNo;
                                                        ws.Cell(writeRow, 3).Value = tc;

                                                        string dbAd = dr["p_ad"]?.ToString() ?? "";
                                                        string dbSoyad = dr["p_soyad"]?.ToString() ?? "";
                                                        if (string.IsNullOrEmpty(dbAd))
                                                        {
                                                            string adSoyad = dr["p_ad_soyad"]?.ToString() ?? "";
                                                            var parcalar = adSoyad.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                                            if (parcalar.Length > 0) dbAd = parcalar[0];
                                                            if (parcalar.Length > 1) dbSoyad = string.Join(" ", parcalar.Skip(1));
                                                        }
                                                        ws.Cell(writeRow, 4).Value = dbAd;
                                                        ws.Cell(writeRow, 5).Value = dbSoyad;

                                                        string gunDetaylari = dr["p_gun_detaylari"]?.ToString() ?? "";
                                                        string[] gunler = gunDetaylari.Split('-');
                                                        int calisilanGunSayisi = 0;
                                                        int raporluGunSayisi = 0;
                                                        for (int gun = 1; gun <= gunSayisi && gun <= gunler.Length; gun++)
                                                        {
                                                            string gunDegeri = gunler[gun - 1];
                                                            if (gunDegeri == "0" || string.IsNullOrEmpty(gunDegeri))
                                                                gunDegeri = "";
                                                            else if (gunDegeri == "X")
                                                                calisilanGunSayisi++;
                                                            else if (gunDegeri == "İ" || gunDegeri == "i")
                                                                raporluGunSayisi++;

                                                            ws.Cell(writeRow, gunKolonBaslangic + gun - 1).Value = gunDegeri;
                                                        }

                                                        string firstDayCol = XLHelper.GetColumnLetterFromNumber(gunKolonBaslangic);
                                                        string lastDayCol = XLHelper.GetColumnLetterFromNumber(gunKolonBaslangic + gunSayisi - 1);
                                                        ws.Cell(writeRow, ozetKolonBaslangic).FormulaA1 = $"COUNTIF({firstDayCol}{writeRow}:{lastDayCol}{writeRow},\"X\")";
                                                        int toplamDevamsizlik = (dr["p_devamsizlik"] == DBNull.Value) ? 0 : Convert.ToInt32(dr["p_devamsizlik"]);
                                                        ws.Cell(writeRow, ozetKolonBaslangic + 1).Value = toplamDevamsizlik > 0 ? toplamDevamsizlik : raporluGunSayisi;
                                                        int yillikIzin = (dr["p_yillik_izin"] == DBNull.Value) ? 0 : Convert.ToInt32(dr["p_yillik_izin"]);
                                                        ws.Cell(writeRow, ozetKolonBaslangic + 2).Value = yillikIzin;
                                                        // Prefer program_katilimcilari dates (pk_*) if available, otherwise fall back to puantaj p_* dates
                                                        try
                                                        {
                                                            bool wrotePkStart = false;
                                                            try
                                                            {
                                                                int ordPkStart = dr.GetOrdinal("pk_ise_baslama");
                                                                if (!dr.IsDBNull(ordPkStart))
                                                                {
                                                                    DateTime iseBaslama = Convert.ToDateTime(dr.GetValue(ordPkStart));
                                                                    ws.Cell(writeRow, ozetKolonBaslangic + 3).Value = iseBaslama.ToString("dd.MM.yyyy");
                                                                    wrotePkStart = true;
                                                                }
                                                            }
                                                            catch (IndexOutOfRangeException) { }

                                                            if (!wrotePkStart)
                                                            {
                                                                try
                                                                {
                                                                    int ordPStart = dr.GetOrdinal("p_ise_baslama_tarihi");
                                                                    if (!dr.IsDBNull(ordPStart))
                                                                    {
                                                                        DateTime iseBaslama = Convert.ToDateTime(dr.GetValue(ordPStart));
                                                                        ws.Cell(writeRow, ozetKolonBaslangic + 3).Value = iseBaslama.ToString("dd.MM.yyyy");
                                                                    }
                                                                }
                                                                catch (IndexOutOfRangeException) { }
                                                            }
                                                        }
                                                        catch { }

                                                        try
                                                        {
                                                            bool wrotePkEnd = false;
                                                            try
                                                            {
                                                                int ordPkEnd = dr.GetOrdinal("pk_isten_ayrilma");
                                                                if (!dr.IsDBNull(ordPkEnd))
                                                                {
                                                                    DateTime istAyrilma = Convert.ToDateTime(dr.GetValue(ordPkEnd));
                                                                    ws.Cell(writeRow, ozetKolonBaslangic + 4).Value = istAyrilma.ToString("dd.MM.yyyy");
                                                                    wrotePkEnd = true;
                                                                }
                                                            }
                                                            catch (IndexOutOfRangeException) { }

                                                            if (!wrotePkEnd)
                                                            {
                                                                try
                                                                {
                                                                    int ordPEnd = dr.GetOrdinal("p_isten_ayrilma_tarihi");
                                                                    if (!dr.IsDBNull(ordPEnd))
                                                                    {
                                                                        DateTime istAyrilma = Convert.ToDateTime(dr.GetValue(ordPEnd));
                                                                        ws.Cell(writeRow, ozetKolonBaslangic + 4).Value = istAyrilma.ToString("dd.MM.yyyy");
                                                                    }
                                                                }
                                                                catch (IndexOutOfRangeException) { }
                                                            }
                                                        }
                                                        catch { }

                                                        siraNo++;
                                                        writeRow++;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                // Save to temporary file first to avoid leaving partial file if target is locked
                                string tempPath = Path.Combine(Path.GetDirectoryName(savePath), Path.GetFileNameWithoutExtension(savePath) + ".tmp" + Path.GetExtension(savePath));
                                workbook.SaveAs(tempPath);

                                bool replaced = false;
                                int attempts = 0;
                                // Try to copy/replace the target file. If target is locked (open in Excel), prompt user to close and retry.
                                while (!replaced)
                                {
                                    try
                                    {
                                        attempts++;
                                        // Overwrite target
                                        File.Copy(tempPath, savePath, true);
                                        // remove temp
                                        try { File.Delete(tempPath); } catch { }
                                        replaced = true;
                                    }
                                    catch (IOException)
                                    {
                                        // If we've tried several times, ask user to close the file or cancel
                                        if (attempts >= 5)
                                        {
                                            var dr = System.Windows.Forms.DialogResult.Cancel;
                                            try
                                            {
                                                this.Invoke(new Action(() =>
                                                {
                                                    dr = MessageBox.Show("Hedef Excel dosyası başka bir uygulama tarafından açık. Lütfen dosyayı kapatın ve 'Tekrar' ile yeniden deneyin. 'İptal' seçerseniz işlem durdurulacaktır.", "Dosya Kilitli", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);
                                                }));
                                            }
                                            catch
                                            {
                                                dr = System.Windows.Forms.DialogResult.Cancel;
                                            }

                                            if (dr == System.Windows.Forms.DialogResult.Retry)
                                            {
                                                attempts = 0; // reset attempts and try again
                                                System.Threading.Thread.Sleep(500);
                                                continue;
                                            }
                                            else
                                            {
                                                // user cancelled -> rethrow to outer catch and notify
                                                try { File.Delete(tempPath); } catch { }
                                                throw new IOException("Hedef dosya kilitli, kullanıcı iptal etti.");
                                            }
                                        }

                                        // short delay then retry
                                        System.Threading.Thread.Sleep(500);
                                    }
                                }
                            }

                            this.Invoke(new Action(() =>
                            {
                                MessageBox.Show($"✅ Şablon kullanılarak Excel oluşturuldu!\n\nDosya: {Path.GetFileName(savePath)}", "Tamamlandı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }));
                        }
                        catch (Exception ex)
                        {
                            this.Invoke(new Action(() =>
                            {
                                MessageBox.Show("Excel şablon doldurma hatası: " + ex.Message + "\n\nDetay: " + ex.StackTrace, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }));
                        }
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Excel Hatası: " + ex.Message + "\n\nDetay: " + ex.StackTrace, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            tlpMain.Controls.Add(tabControlKampusler, 0, 1);
        }
    }
}
