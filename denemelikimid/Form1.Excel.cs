using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using MySql.Data.MySqlClient;
using denemelikimid.Validations;

namespace denemelikimid
{
    public partial class Form1
    {
        private void LoadExcelView()
        {
            panelContent.Controls.Clear();

            // ARAYÜZ OLUŞTURMA KISMI
            Panel panelContainer = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Padding = new Padding(30)
            };
            panelContent.Controls.Add(panelContainer);

            // Başlık
            Label lblHeader = new Label
            {
                Text = "📄 İŞKUR Puantaj ve Banka Entegrasyonu",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = colorTextPrimary,
                AutoSize = true,
                Dock = DockStyle.Top
            };
            panelContainer.Controls.Add(lblHeader);

            // Araçlar Paneli
            Panel pnlTools = new Panel
            {
                Dock = DockStyle.Top,
                Height = 120,
                BackColor = Color.Transparent,
                Padding = new Padding(0, 20, 0, 0)
            };
            panelContainer.Controls.Add(pnlTools);
            pnlTools.BringToFront();

            // 1. BUTON: EXCEL'DEN PUANTAJ YÜKLE
            Button btnImport = CreateModernButton("📥 1. Puantaj Yükle", colorSuccess, 0, pnlTools);
            btnImport.Click += (s, e) =>
            {
                Action<string[]> processFiles = files =>
                {
                    int toplamAktarilan = 0;
                    int toplamAtlanan = 0;
                    int basarisizDosya = 0;
                    List<string> dosyaOzetleri = new List<string>();

                    foreach (var file in files)
                    {
                        try
                        {
                            var result = ImportBuuExcelFile(file, false);
                            toplamAktarilan += result.ImportedCount;
                            toplamAtlanan += result.SkippedCount;
                            dosyaOzetleri.Add($"• {System.IO.Path.GetFileName(file)} → {result.Kampus} ({result.ImportedCount} kayıt)");
                        }
                        catch (Exception ex)
                        {
                            basarisizDosya++;
                            string err = ex.Message;
                            if (err.Length > 80) err = err.Substring(0, 80) + "...";
                            dosyaOzetleri.Add($"• {System.IO.Path.GetFileName(file)} → ❌ {err}");
                        }
                    }

                    MessageBox.Show(
                        $"✅ Toplu Excel içe aktarma tamamlandı!\n\n" +
                        $"📁 İşlenen dosya: {files.Length}\n" +
                        $"✅ Toplam içe aktarılan: {toplamAktarilan}\n" +
                        $"⏭️ Toplam atlanan: {toplamAtlanan}\n" +
                        $"❌ Hatalı dosya: {basarisizDosya}\n\n" +
                        string.Join("\n", dosyaOzetleri),
                        "Toplu İçe Aktarma Sonucu",
                        MessageBoxButtons.OK,
                        basarisizDosya > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                };

                string sabitExcelKlasoru = @"C:\Users\mehme\source\repos\Personel-Maas-Sistemi-Arayuzu\denemelikimid";
                if (System.IO.Directory.Exists(sabitExcelKlasoru))
                {
                    var klasorDosyalari = System.IO.Directory
                        .GetFiles(sabitExcelKlasoru, "*.xlsx")
                        .Concat(System.IO.Directory.GetFiles(sabitExcelKlasoru, "*.xls"))
                        .Where(f => !System.IO.Path.GetFileName(f).StartsWith("~$"))
                        .ToArray();

                    if (klasorDosyalari.Length > 0)
                    {
                        var secim = MessageBox.Show(
                            $"Belirttiğiniz klasörde {klasorDosyalari.Length} Excel dosyası bulundu.\n\n" +
                            "Bu dosyaların tamamı içe aktarılsın mı?",
                            "Klasörden Otomatik Yükleme",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Question);

                        if (secim == DialogResult.Yes)
                        {
                            processFiles(klasorDosyalari);
                            return;
                        }

                        if (secim == DialogResult.Cancel)
                            return;
                    }
                }

                OpenFileDialog ofd = new OpenFileDialog
                {
                    Filter = "Excel Dosyaları|*.xlsx;*.xls",
                    Title = "Puantaj Dosyalarını Seçin (Kampüs1, Kampüs2, Kampüs3)",
                    Multiselect = true,
                    InitialDirectory = sabitExcelKlasoru
                };

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        processFiles(ofd.FileNames);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Hata: " + ex.Message + "\n\nDetay: " + ex.StackTrace, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            };
            pnlTools.Controls.Add(btnImport);

            // 2. BUTON: MAAŞ HESAPLA
            Button btnCalculate = CreateModernButton("💰 2. Maaş Hesapla", Color.Orange, 1, pnlTools);
            btnCalculate.Click += (s, e) =>
            {
                try
                {
                    float gunlukUcret = 500.0f;

                    using (var conn =
                           new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        new MySqlCommand("TRUNCATE TABLE banka_listesi", conn).ExecuteNonQuery();

                        string sql = @"INSERT INTO banka_listesi (bl_tc, bl_ad_soyad, bl_iban_no, bl_tutar)
                               SELECT p_tc, p_ad_soyad, p_iban, (p_calistigi_gun_sayisi * @ucret) 
                               FROM puantaj 
                               WHERE p_calistigi_gun_sayisi > 0";

                        using (var cmd = new MySqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@ucret", gunlukUcret);
                            int sayi = cmd.ExecuteNonQuery();
                            MessageBox.Show(
                                $"✅ {sayi} kişinin maaşı hesaplandı ve banka listesine yazıldı.",
                                "Tamamlandı");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hesaplama Hatası: " + ex.Message);
                }
            };
            pnlTools.Controls.Add(btnCalculate);

            // 3. BUTON: BANKA LİSTESİ İNDİR
            Button btnExport = CreateModernButton("📤 3. Banka Listesi İndir", colorInfo, 2, pnlTools);
            btnExport.Click += (s, e) =>
            {
                try
                {
                    DataTable dt = new DataTable();
                    using (var conn =
                           new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        string sql =
                            "SELECT bl_ad_soyad AS 'Ad Soyad', bl_iban_no AS 'IBAN', bl_tutar AS 'Tutar' FROM banka_listesi";
                        using (var da = new MySqlDataAdapter(sql, conn))
                        {
                            da.Fill(dt);
                        }
                    }

                    SaveFileDialog sfd = new SaveFileDialog
                    {
                        Filter = "Excel Dosyası|*.xlsx",
                        FileName = $"Banka_Odeme_Listesi_{DateTime.Now:yyyy-MM-dd}.xlsx"
                    };

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Banka Listesi");
                            worksheet.Cell(1, 1).InsertTable(dt);
                            worksheet.Columns().AdjustToContents();
                            workbook.SaveAs(sfd.FileName);
                        }
                        MessageBox.Show("✅ Banka listesi Excel dosyası olarak kaydedildi!",
                            "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            Label lblInfo = new Label
            {
                Text =
                    "ℹ️ Sistem 'iskur' veritabanına bağlıdır. Excel dosyanızda sırasıyla: TC, Ad Soyad, IBAN ve Gün Sayısı olmalıdır.",
                Font = new Font("Segoe UI", 10, FontStyle.Italic),
                ForeColor = colorTextSecondary,
                Dock = DockStyle.Bottom,
                Padding = new Padding(0, 20, 0, 0),
                Height = 100
            };
            panelContainer.Controls.Add(lblInfo);
        }

        /// <summary>
        /// BUÜ formatındaki Excel dosyasını import eder (Bordro ve Puantaj bilgilerini içerir)
        /// </summary>
        private (int ImportedCount, int SkippedCount, string Kampus) ImportBuuExcelFile(string filePath, bool showMessage = true)
        {
            try
            {
                string ext = System.IO.Path.GetExtension(filePath)?.ToLowerInvariant() ?? "";
                if (ext == ".xls")
                {
                    // .xls için OleDb fallback
                    return ImportBuuExcelFileWithOleDb(filePath, showMessage);
                }

                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var usedRange = worksheet.RangeUsed();
                    
                    if (usedRange == null)
                    {
                        MessageBox.Show("Dosya boş görünüyor.");
                        return (0, 0, "Bilinmiyor");
                    }

                    // Dosya adından kampüs bilgisini çıkar (13376 = Kampüs1, 13377 = Kampüs2, 13378 = Kampüs3)
                    string fileName = System.IO.Path.GetFileName(filePath);
                    string kampus = "Kampüs1";
                    if (fileName.Contains("13377")) kampus = "Kampüs2";
                    else if (fileName.Contains("13378")) kampus = "Kampüs3";

                    int firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
                    int lastRow = usedRange.RangeAddress.LastAddress.RowNumber;
                    int firstCol = usedRange.RangeAddress.FirstAddress.ColumnNumber;
                    int lastCol = usedRange.RangeAddress.LastAddress.ColumnNumber;

                    string NormalizeHeader(string s)
                    {
                        if (string.IsNullOrWhiteSpace(s)) return "";
                        return s.Trim().ToUpperInvariant()
                            .Replace("İ", "I")
                            .Replace("Ğ", "G")
                            .Replace("Ü", "U")
                            .Replace("Ş", "S")
                            .Replace("Ö", "O")
                            .Replace("Ç", "C");
                    }

                    string CleanDigits(string s) => new string((s ?? "").Where(char.IsDigit).ToArray());
                    bool HasLetter(string s) => !string.IsNullOrWhiteSpace(s) && s.Any(char.IsLetter);

                    // 1) Önce başlıklardan kolonları bulmaya çalış
                    int headerRow = -1;
                    int tcCol = -1, adSoyadCol = -1, adCol = -1, soyadCol = -1, ibanCol = -1, gunCol = -1;

                    for (int row = firstRow; row <= Math.Min(lastRow, firstRow + 25); row++)
                    {
                        for (int col = firstCol; col <= lastCol; col++)
                        {
                            string h = NormalizeHeader(worksheet.Cell(row, col).GetValue<string>());
                            if (string.IsNullOrEmpty(h)) continue;

                            if ((h.Contains("TC") || h.Contains("KIMLIK") || h.Contains("T.C")) && tcCol == -1)
                            {
                                tcCol = col;
                                if (headerRow == -1) headerRow = row;
                            }
                            if ((h.Contains("AD") && h.Contains("SOYAD")) || h.Contains("ADSOYAD") || h.Contains("ADI SOYADI"))
                            {
                                if (adSoyadCol == -1) adSoyadCol = col;
                            }
                            else if (h == "AD" || h.Contains("ADI") || h.Contains("KATILIMCI"))
                            {
                                if (adCol == -1) adCol = col;
                            }
                            else if (h.Contains("SOYAD") || h.Contains("BILGILERI") || h.Contains("BILGILERI"))
                            {
                                if (soyadCol == -1) soyadCol = col;
                            }
                            if (h.Contains("IBAN") && ibanCol == -1) ibanCol = col;
                            if ((h.Contains("GUN") || h.Contains("CALISTIGI GUN")) && gunCol == -1) gunCol = col;
                        }
                    }

                    // 2) Başlıktan TC bulunamazsa, satırlardaki geçerli TC yoğunluğundan kolon tahmin et
                    if (tcCol == -1)
                    {
                        int bestCol = -1;
                        int bestScore = 0;
                        for (int col = firstCol; col <= lastCol; col++)
                        {
                            int score = 0;
                            for (int row = firstRow; row <= Math.Min(lastRow, firstRow + 600); row++)
                            {
                                string tcTry = CleanDigits(worksheet.Cell(row, col).GetValue<string>());
                                if (InputValidator.IsValidTCNumber(tcTry)) score++;
                            }
                            if (score > bestScore)
                            {
                                bestScore = score;
                                bestCol = col;
                            }
                        }
                        if (bestScore > 0)
                            tcCol = bestCol;
                    }

                    if (tcCol == -1)
                        throw new Exception("TC kolonu tespit edilemedi.");

                    // 3) Yardımcı kolon tahminleri
                    if (adSoyadCol == -1)
                    {
                        if (adCol > 0 && soyadCol > 0)
                        {
                            // ayrı alanlardan birleştirilecek
                        }
                        else if (tcCol - 2 >= firstCol)
                        {
                            adCol = tcCol - 2;
                            soyadCol = tcCol - 1;
                        }
                        else if (tcCol - 1 >= firstCol)
                        {
                            adSoyadCol = tcCol - 1;
                        }
                    }

                    if (ibanCol == -1)
                    {
                        int bestIbanCol = -1;
                        int bestIbanScore = 0;
                        for (int col = firstCol; col <= lastCol; col++)
                        {
                            int score = 0;
                            for (int row = firstRow; row <= Math.Min(lastRow, firstRow + 600); row++)
                            {
                                string v = (worksheet.Cell(row, col).GetValue<string>() ?? "").Replace(" ", "").ToUpperInvariant();
                                if (v.StartsWith("TR") && v.Length >= 26) score++;
                            }
                            if (score > bestIbanScore)
                            {
                                bestIbanScore = score;
                                bestIbanCol = col;
                            }
                        }
                        if (bestIbanScore > 0) ibanCol = bestIbanCol;
                    }

                    // 4) Gün kolonlarını tespit et (1..31 başlıkları)
                    int ExtractLeadingDay(string text)
                    {
                        if (string.IsNullOrWhiteSpace(text)) return -1;
                        text = text.Trim();
                        string digits = new string(text.TakeWhile(char.IsDigit).ToArray());
                        int d;
                        if (int.TryParse(digits, out d) && d >= 1 && d <= 31) return d;

                        var token = text.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                        if (int.TryParse(token, out d) && d >= 1 && d <= 31) return d;
                        return -1;
                    }

                    Dictionary<int, int> dayColumnMap = new Dictionary<int, int>();
                    for (int col = firstCol; col <= lastCol; col++)
                    {
                        for (int row = firstRow; row <= Math.Min(lastRow, firstRow + 35); row++)
                        {
                            string h = worksheet.Cell(row, col).GetValue<string>();
                            int dayNo = ExtractLeadingDay(h);
                            if (dayNo > 0)
                            {
                                if (!dayColumnMap.ContainsKey(dayNo))
                                    dayColumnMap[dayNo] = col;
                                break;
                            }
                        }
                    }

                    using (var conn = new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                    {
                        conn.Open();
                        
                        int importedCount = 0;
                        int skippedCount = 0;

                        // Veri satırlarını işle
                        for (int row = firstRow; row <= lastRow; row++)
                        {
                            try
                            {
                                string tcRaw = worksheet.Cell(row, tcCol).GetValue<string>()?.Trim() ?? "";
                                string tc = new string(tcRaw.Where(char.IsDigit).ToArray());
                                string adSoyad = "";
                                string adHucre = adCol > 0 ? (worksheet.Cell(row, adCol).GetValue<string>()?.Trim() ?? "") : "";
                                string soyadHucre = soyadCol > 0 ? (worksheet.Cell(row, soyadCol).GetValue<string>()?.Trim() ?? "") : "";

                                if (!string.IsNullOrWhiteSpace(adHucre) && !string.IsNullOrWhiteSpace(soyadHucre))
                                {
                                    adSoyad = (adHucre + " " + soyadHucre).Trim();
                                }

                                // Ad/Soyad yanlış kolondan (ör. SIRA) geldiyse satırdaki metin alanlarından toparla
                                if (!HasLetter(adSoyad))
                                {
                                    var metinHucreleri = new List<string>();
                                    for (int c = firstCol; c <= lastCol; c++)
                                    {
                                        if (c == tcCol || c == ibanCol) continue;
                                        if (dayColumnMap.Values.Contains(c)) continue;

                                        string v = (worksheet.Cell(row, c).GetValue<string>() ?? "").Trim();
                                        if (string.IsNullOrWhiteSpace(v)) continue;

                                        string vUpper = v.ToUpperInvariant();
                                        if (vUpper.Contains("SIRA") || vUpper.Contains("KATILIMCI") || vUpper.Contains("BILGI") || vUpper.Contains("IBAN") || vUpper.Contains("TC"))
                                            continue;

                                        if (v.All(char.IsDigit)) continue;
                                        if (v.Replace(" ", "").ToUpperInvariant().StartsWith("TR") && v.Replace(" ", "").Length >= 26) continue;
                                        if (InputValidator.IsValidTCNumber(CleanDigits(v))) continue;

                                        if (HasLetter(v))
                                            metinHucreleri.Add(v);
                                    }

                                    if (metinHucreleri.Count >= 2)
                                        adSoyad = (metinHucreleri[0] + " " + metinHucreleri[1]).Trim();
                                    else if (metinHucreleri.Count == 1)
                                        adSoyad = metinHucreleri[0].Trim();
                                }
                                else if (adSoyadCol > 0)
                                {
                                    adSoyad = worksheet.Cell(row, adSoyadCol).GetValue<string>()?.Trim() ?? "";
                                }
                                else
                                {
                                    adSoyad = (adHucre + " " + soyadHucre).Trim();
                                }

                                string iban = ibanCol > 0 ? (worksheet.Cell(row, ibanCol).GetValue<string>()?.Trim() ?? "") : "";
                                
                                // Gün sayısını hesapla (eğer kolon varsa)
                                int gunSayisi = 0;
                                if (gunCol > 0)
                                {
                                    var gunValue = worksheet.Cell(row, gunCol).GetValue<string>();
                                    if (string.IsNullOrEmpty(gunValue))
                                    {
                                        gunValue = worksheet.Cell(row, gunCol).GetValue<double>().ToString();
                                    }
                                    int.TryParse(gunValue, out gunSayisi);
                                }

                                // Gün detaylarını (X/İ/R) al
                                string gunDetaylari = null;
                                if (dayColumnMap.Count > 0)
                                {
                                    int maxDay = dayColumnMap.Keys.Max();
                                    List<string> gunler = new List<string>();
                                    int xCount = 0;
                                    for (int d = 1; d <= maxDay; d++)
                                    {
                                        string v = "0";
                                        int dayCol;
                                        if (dayColumnMap.TryGetValue(d, out dayCol))
                                        {
                                            string raw = (worksheet.Cell(row, dayCol).GetValue<string>() ?? "").Trim().ToUpperInvariant();
                                            if (raw == "X")
                                            {
                                                v = "X";
                                                xCount++;
                                            }
                                            else if (raw == "İ" || raw == "I")
                                            {
                                                v = "İ";
                                            }
                                            else if (raw == "R")
                                            {
                                                v = "R";
                                            }
                                        }
                                        gunler.Add(v);
                                    }
                                    gunDetaylari = string.Join("-", gunler);
                                    gunSayisi = xCount;
                                }

                                // Boş satırları atla
                                if (string.IsNullOrEmpty(tc) && string.IsNullOrEmpty(adSoyad))
                                {
                                    skippedCount++;
                                    continue;
                                }

                                // Geçersiz/başlık satırlarını atla
                                string adSoyadUpper = (adSoyad ?? "").ToUpperInvariant();
                                if (!InputValidator.IsValidTCNumber(tc) ||
                                    adSoyadUpper.Contains("AD SOYAD") ||
                                    adSoyadUpper.Contains("KATILIMCI") ||
                                    adSoyadUpper.Contains("BİLGİ") ||
                                    adSoyadUpper.Contains("BILGI") ||
                                    adSoyadUpper.Contains("SIRA") ||
                                    adSoyadUpper.Contains("IBAN") ||
                                    !HasLetter(adSoyad))
                                {
                                    skippedCount++;
                                    continue;
                                }

                                // Puantaj tablosuna ekle/güncelle
                                string queryPuantaj = @"INSERT INTO puantaj 
                                    (p_tc, p_ad_soyad, p_ad, p_soyad, p_iban, p_calistigi_gun_sayisi, p_ise_baslama_tarihi) 
                                    VALUES (@tc, @adsoy, @ad, @soy, @iban, @gun, CURDATE())
                                    ON DUPLICATE KEY UPDATE 
                                    p_ad_soyad = @adsoy, 
                                    p_ad = @ad, 
                                    p_soyad = @soy, 
                                    p_iban = @iban, 
                                    p_calistigi_gun_sayisi = @gun";

                                // isim/soyisim: son kelime soyisim, geri kalanı ad
                                string firstName = "";
                                string lastName = "";
                                if (!string.IsNullOrWhiteSpace(adSoyad))
                                {
                                    var np = adSoyad.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                    if (np.Length == 1)
                                    {
                                        firstName = np[0];
                                    }
                                    else if (np.Length > 1)
                                    {
                                        lastName = np[np.Length - 1];
                                        firstName = string.Join(" ", np.Take(np.Length - 1));
                                    }
                                }

                                adSoyad = (firstName + " " + lastName).Trim();

                                using (var cmd = new MySqlCommand(queryPuantaj, conn))
                                {
                                    cmd.Parameters.AddWithValue("@tc", tc);
                                    cmd.Parameters.AddWithValue("@adsoy", adSoyad);
                                    cmd.Parameters.AddWithValue("@ad", firstName);
                                    cmd.Parameters.AddWithValue("@soy", lastName);
                                    cmd.Parameters.AddWithValue("@iban", iban);
                                    cmd.Parameters.AddWithValue("@detay", (object)gunDetaylari ?? DBNull.Value);
                                    cmd.Parameters.AddWithValue("@gun", gunSayisi);
                                    cmd.ExecuteNonQuery();
                                }

                                // Program katılımcıları tablosuna ekle/güncelle (kampüs bilgisiyle)
                                string queryKatilimci = @"INSERT INTO program_katilimcilari 
                                    (pk_tc, pk_ad_soyad, pk_ad, pk_soyad, pk_iban_no, pk_gorev_yeri, pk_is_baslama_tarihi) 
                                    VALUES (@tc, @adsoy, @ad, @soy, @iban, @kampus, CURDATE())
                                    ON DUPLICATE KEY UPDATE 
                                    pk_ad_soyad = @adsoy, 
                                    pk_ad = @ad, 
                                    pk_soyad = @soy, 
                                    pk_iban_no = @iban, 
                                    pk_gorev_yeri = @kampus";

                                using (var cmd = new MySqlCommand(queryKatilimci, conn))
                                {
                                    cmd.Parameters.AddWithValue("@tc", tc);
                                    cmd.Parameters.AddWithValue("@adsoy", adSoyad);
                                    cmd.Parameters.AddWithValue("@ad", firstName);
                                    cmd.Parameters.AddWithValue("@soy", lastName);
                                    cmd.Parameters.AddWithValue("@iban", iban);
                                    cmd.Parameters.AddWithValue("@kampus", kampus);
                                    cmd.ExecuteNonQuery();
                                }

                                importedCount++;
                            }
                            catch
                            {
                                skippedCount++;
                                // Hata olan satırları atla ve devam et
                                continue;
                            }
                        }

                        if (showMessage)
                        {
                            MessageBox.Show($"✅ İşlem tamamlandı!\n\n" +
                                $"✅ İçe aktarılan: {importedCount} kayıt\n" +
                                $"⏭️ Atlanan: {skippedCount} satır\n" +
                                $"📍 Kampüs: {kampus}",
                                "İşlem Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                        return (importedCount, skippedCount, kampus);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Excel dosyası okunurken hata oluştu: {ex.Message}", ex);
            }
        }

        private (int ImportedCount, int SkippedCount, string Kampus) ImportBuuExcelFileWithOleDb(string filePath, bool showMessage)
        {
            string fileName = System.IO.Path.GetFileName(filePath);
            string kampus = "Kampüs1";
            if (fileName.Contains("13377")) kampus = "Kampüs2";
            else if (fileName.Contains("13378")) kampus = "Kampüs3";

            DataTable dt = new DataTable();

            string[] connStrings = new[]
            {
                $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={filePath};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\";",
                $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\";",
                $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\";"
            };

            Exception lastEx = null;
            bool loaded = false;

            foreach (var connStr in connStrings)
            {
                try
                {
                    using (var conn = new OleDbConnection(connStr))
                    {
                        conn.Open();
                        var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        if (schema == null || schema.Rows.Count == 0)
                            throw new Exception("Excel sayfası bulunamadı.");

                        string sheetName = schema.Rows
                            .Cast<DataRow>()
                            .Select(r => r["TABLE_NAME"].ToString())
                            .FirstOrDefault(n => n.EndsWith("$") || n.EndsWith("$'"));

                        if (string.IsNullOrWhiteSpace(sheetName))
                            throw new Exception("Okunabilir Excel sayfası bulunamadı.");

                        sheetName = sheetName.Trim('\'');

                        using (var da = new OleDbDataAdapter($"SELECT * FROM [{sheetName}]", conn))
                        {
                            da.Fill(dt);
                        }

                        loaded = true;
                        break;
                    }
                }
                catch (Exception ex)
                {
                    lastEx = ex;
                }
            }

            if (!loaded)
            {
                throw new Exception(
                    ".xls dosyası okunamadı. Sisteminizde uygun Excel OLE DB sağlayıcısı bulunmuyor olabilir. " +
                    "Lütfen dosyayı Excel'de 'Farklı Kaydet' ile .xlsx formatına çevirip tekrar deneyin. " +
                    "Detay: " + (lastEx != null ? lastEx.Message : "Bilinmeyen hata"),
                    lastEx);
            }

            if (dt.Rows.Count == 0)
                return (0, 0, kampus);

            int tcCol = -1, adSoyadCol = -1, ibanCol = -1, gunCol = -1;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                string col = (dt.Columns[i].ColumnName ?? "").ToUpperInvariant();
                if (tcCol == -1 && (col.Contains("TC") || col.Contains("KIMLIK") || col.Contains("KİMLİK") || col.Contains("T.C"))) tcCol = i;
                else if (adSoyadCol == -1 && col.Contains("AD") && col.Contains("SOYAD")) adSoyadCol = i;
                else if (ibanCol == -1 && col.Contains("IBAN")) ibanCol = i;
                else if (gunCol == -1 && (col.Contains("GUN") || col.Contains("GÜN"))) gunCol = i;
            }

            if (tcCol == -1) tcCol = 0;
            if (adSoyadCol == -1) adSoyadCol = Math.Min(1, dt.Columns.Count - 1);
            if (ibanCol == -1) ibanCol = Math.Min(2, dt.Columns.Count - 1);
            if (gunCol == -1) gunCol = Math.Min(3, dt.Columns.Count - 1);

            int importedCount = 0;
            int skippedCount = 0;
            bool HasLetter(string s) => !string.IsNullOrWhiteSpace(s) && s.Any(char.IsLetter);

            using (var conn = new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
            {
                conn.Open();

                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        string tc = (row[tcCol]?.ToString() ?? "").Trim();
                        tc = new string(tc.Where(char.IsDigit).ToArray());

                        string adSoyad = (row[adSoyadCol]?.ToString() ?? "").Trim();
                        string iban = (row[ibanCol]?.ToString() ?? "").Trim();

                        if (!HasLetter(adSoyad))
                        {
                            var metinler = new List<string>();
                            for (int c = 0; c < dt.Columns.Count; c++)
                            {
                                if (c == tcCol || c == ibanCol || c == gunCol) continue;
                                string v = (row[c]?.ToString() ?? "").Trim();
                                if (string.IsNullOrWhiteSpace(v)) continue;
                                string vu = v.ToUpperInvariant();
                                if (vu.Contains("SIRA") || vu.Contains("KATILIMCI") || vu.Contains("BILGI") || vu.Contains("IBAN") || vu.Contains("TC")) continue;
                                if (v.All(char.IsDigit)) continue;
                                if (HasLetter(v)) metinler.Add(v);
                            }

                            if (metinler.Count >= 2) adSoyad = (metinler[0] + " " + metinler[1]).Trim();
                            else if (metinler.Count == 1) adSoyad = metinler[0];
                        }

                        if (tc.ToUpperInvariant().Contains("TC") || adSoyad.ToUpperInvariant().Contains("AD SOYAD"))
                        {
                            skippedCount++;
                            continue;
                        }

                        // Geçersiz/başlık satırlarını atla
                        string adSoyadUpper = (adSoyad ?? "").ToUpperInvariant();
                        if (!InputValidator.IsValidTCNumber(tc) ||
                            adSoyadUpper.Contains("KATILIMCI") ||
                            adSoyadUpper.Contains("BİLGİ") ||
                            adSoyadUpper.Contains("BILGI") ||
                            adSoyadUpper.Contains("SIRA") ||
                            adSoyadUpper.Contains("IBAN") ||
                            !HasLetter(adSoyad))
                        {
                            skippedCount++;
                            continue;
                        }

                        if (string.IsNullOrWhiteSpace(tc) && string.IsNullOrWhiteSpace(adSoyad))
                        {
                            skippedCount++;
                            continue;
                        }

                        int gunSayisi = 0;
                        var gv = row[gunCol];
                        if (gv != null && gv != DBNull.Value)
                        {
                            if (gv is double d) gunSayisi = Convert.ToInt32(Math.Round(d));
                            else int.TryParse(gv.ToString(), out gunSayisi);
                        }

                        string firstName = "";
                        string lastName = "";
                        if (!string.IsNullOrEmpty(adSoyad))
                        {
                            var np = adSoyad.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            if (np.Length == 1)
                            {
                                firstName = np[0];
                            }
                            else if (np.Length > 1)
                            {
                                lastName = np[np.Length - 1];
                                firstName = string.Join(" ", np.Take(np.Length - 1));
                            }
                        }

                        string queryPuantaj = @"INSERT INTO puantaj 
                            (p_tc, p_ad_soyad, p_ad, p_soyad, p_iban, p_calistigi_gun_sayisi, p_ise_baslama_tarihi) 
                            VALUES (@tc, @adsoy, @ad, @soy, @iban, @gun, CURDATE())
                            ON DUPLICATE KEY UPDATE 
                            p_ad_soyad = @adsoy, p_ad = @ad, p_soyad = @soy, p_iban = @iban, p_calistigi_gun_sayisi = @gun";

                        using (var cmd = new MySqlCommand(queryPuantaj, conn))
                        {
                            cmd.Parameters.AddWithValue("@tc", tc);
                            cmd.Parameters.AddWithValue("@adsoy", adSoyad);
                            cmd.Parameters.AddWithValue("@ad", firstName);
                            cmd.Parameters.AddWithValue("@soy", lastName);
                            cmd.Parameters.AddWithValue("@iban", iban);
                            cmd.Parameters.AddWithValue("@gun", gunSayisi);
                            cmd.ExecuteNonQuery();
                        }

                        string queryKatilimci = @"INSERT INTO program_katilimcilari 
                            (pk_tc, pk_ad_soyad, pk_ad, pk_soyad, pk_iban_no, pk_gorev_yeri, pk_is_baslama_tarihi) 
                            VALUES (@tc, @adsoy, @ad, @soy, @iban, @kampus, CURDATE())
                            ON DUPLICATE KEY UPDATE 
                            pk_ad_soyad = @adsoy, pk_ad = @ad, pk_soyad = @soy, pk_iban_no = @iban, pk_gorev_yeri = @kampus";

                        using (var cmd = new MySqlCommand(queryKatilimci, conn))
                        {
                            cmd.Parameters.AddWithValue("@tc", tc);
                            cmd.Parameters.AddWithValue("@adsoy", adSoyad);
                            cmd.Parameters.AddWithValue("@ad", firstName);
                            cmd.Parameters.AddWithValue("@soy", lastName);
                            cmd.Parameters.AddWithValue("@iban", iban);
                            cmd.Parameters.AddWithValue("@kampus", kampus);
                            cmd.ExecuteNonQuery();
                        }

                        importedCount++;
                    }
                    catch
                    {
                        skippedCount++;
                    }
                }
            }

            if (showMessage)
            {
                MessageBox.Show($"✅ İşlem tamamlandı!\n\n✅ İçe aktarılan: {importedCount} kayıt\n⏭️ Atlanan: {skippedCount} satır\n📍 Kampüs: {kampus}",
                    "İşlem Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return (importedCount, skippedCount, kampus);
        }
    }
}


