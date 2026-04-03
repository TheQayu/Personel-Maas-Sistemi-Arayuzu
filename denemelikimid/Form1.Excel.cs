using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using denemelikimid.DataBase;
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
                    float gunlukUcret = 1375.00f;

                    using (var conn = DbConnection.GetConnection())
                    {
                        conn.Open();
                        new SqliteCommand("DELETE FROM banka_listesi", conn).ExecuteNonQuery();

                        string sql = @"INSERT INTO banka_listesi (bl_tc, bl_ad_soyad, bl_iban_no, bl_tutar)
                               SELECT p_tc, p_ad_soyad, p_iban, (p_calistigi_gun_sayisi * @ucret) 
                               FROM puantaj 
                               WHERE p_calistigi_gun_sayisi > 0";

                        using (var cmd = new SqliteCommand(sql, conn))
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
                    using (var conn = DbConnection.GetConnection())
                    {
                        conn.Open();
                        string sql =
                            "SELECT bl_ad_soyad AS 'Ad Soyad', bl_iban_no AS 'IBAN', bl_tutar AS 'Tutar' FROM banka_listesi";
                        using (var cmd = new SqliteCommand(sql, conn))
                        using (var reader = cmd.ExecuteReader())
                        {
                            dt.Load(reader);
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

        private bool IsExcelHeaderLike(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;

            string normalized = text.Trim().ToUpperInvariant()
                .Replace("İ", "I")
                .Replace("Ğ", "G")
                .Replace("Ü", "U")
                .Replace("Ş", "S")
                .Replace("Ö", "O")
                .Replace("Ç", "C");

            return normalized == "SIRA"
                   || normalized == "KATILIMCI"
                   || normalized == "KATILIMCI BILGILERI"
                   || normalized == "BILGILERI"
                   || normalized == "IBAN"
                   || normalized == "TC"
                   || normalized == "AD SOYAD";
        }

        private static void SplitFullName(string fullName, out string first, out string last)
        {
            first = "";
            last = "";
            if (string.IsNullOrWhiteSpace(fullName)) return;
            var parts = fullName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 1)
            {
                first = parts[0];
            }
            else if (parts.Length > 1)
            {
                last = parts[parts.Length - 1];
                first = string.Join(" ", parts.Take(parts.Length - 1));
            }
        }

        private static string NormalizeNameValue(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            return text.Trim().ToUpperInvariant()
                .Replace("İ", "I")
                .Replace("Ğ", "G")
                .Replace("Ü", "U")
                .Replace("Ş", "S")
                .Replace("Ö", "O")
                .Replace("Ç", "C");
        }

        private static bool IsMonthLike(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            string normalized = NormalizeNameValue(text);
            string token = new string(normalized.Where(char.IsLetter).ToArray());
            if (string.IsNullOrWhiteSpace(token) || token.Length < 3) return false;
            string[] months =
            {
                "OCAK", "SUBAT", "MART", "NISAN", "MAYIS", "HAZIRAN",
                "TEMMUZ", "AGUSTOS", "EYLUL", "EKIM", "KASIM", "ARALIK"
            };
            return months.Any(m => m.StartsWith(token) || token.StartsWith(m));
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
                    bool IsLikelyTc(string s) => !string.IsNullOrWhiteSpace(s) && s.Length == 11 && s.All(char.IsDigit) && s[0] != '0';
                    bool IsNameNoise(string text)
                    {
                        if (string.IsNullOrWhiteSpace(text)) return true;
                        string t = NormalizeHeader(text);
                        return IsExcelHeaderLike(t)
                               || t.Contains("KAMPUS")
                               || t == "GUVENLIK" || t == "TEMIZLIK" || t == "IDARI" || t == "BILISIM" || t == "TEKNIK" || t == "AKADEMIK"
                               || IsMonthLike(t)
                               || t == "X" || t == "I" || t == "R";
                    }

                    string ReadTcFromCell(IXLCell cell)
                    {
                        string raw = cell.GetValue<string>()?.Trim() ?? "";
                        string digits = CleanDigits(raw);
                        if (digits.Length == 11) return digits;

                        try
                        {
                            if (cell.DataType == XLDataType.Number)
                            {
                                string n = Convert.ToInt64(Math.Round(cell.GetDouble())).ToString();
                                if (n.Length == 11) return n;
                            }
                        }
                        catch { }

                        return digits;
                    }

                    // 1) Önce başlıklardan kolonları bulmaya çalış
                    int headerRow = -1;
                    int tcCol = -1, adSoyadCol = -1, adCol = -1, soyadCol = -1, ibanCol = -1, gunCol = -1, iseGirisCol = -1;

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
                            else if (h == "AD" || h == "ADI")
                            {
                                if (adCol == -1) adCol = col;
                            }
                            else if (h == "SOYAD" || h == "SOYADI")
                            {
                                if (soyadCol == -1) soyadCol = col;
                            }
                            if (h.Contains("IBAN") && ibanCol == -1) ibanCol = col;
                            if ((h.Contains("GUN") || h.Contains("CALISTIGI GUN")) && gunCol == -1) gunCol = col;
                            if ((h.Contains("ISE GIRIS") || h.Contains("İŞE GİRİŞ") || h.Contains("BASLAMA TARIHI") || h.Contains("BAŞLAMA TARİHİ")) && iseGirisCol == -1) iseGirisCol = col;
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

                    // 3) Yardımcı kolon tahminleri (Sadece adSoyadCol eksikse, adCol/soyadCol varsa kullanacağız, rastgele ezme yok.)
                    if (adSoyadCol == -1 && adCol > 0 && soyadCol > 0)
                    {
                        // ayrı alanlardan birleştirilecek, bu yüzden adSoyadCol'a dokunmuyoruz.
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

                    bool IsNameCandidate(string value)
                    {
                        if (string.IsNullOrWhiteSpace(value)) return false;
                        string v = value.Trim();
                        if (IsNameNoise(v)) return false;
                        if (v.All(char.IsDigit)) return false;
                        if (v.Replace(" ", "").ToUpperInvariant().StartsWith("TR") && v.Replace(" ", "").Length >= 26) return false;
                        if (InputValidator.IsValidTCNumber(CleanDigits(v))) return false;
                        return HasLetter(v);
                    }

                    if ((adCol == -1 || soyadCol == -1) && tcCol > 0)
                    {
                        var nameScores = new Dictionary<int, int>();
                        for (int col = tcCol + 1; col <= lastCol; col++)
                        {
                            if (col == tcCol || col == ibanCol || col == gunCol || col == adSoyadCol) continue;
                            if (dayColumnMap.Values.Contains(col)) continue;
                            int score = 0;
                            for (int row = firstRow; row <= Math.Min(lastRow, firstRow + 200); row++)
                            {
                                string v = (worksheet.Cell(row, col).GetValue<string>() ?? "").Trim();
                                if (IsNameCandidate(v)) score++;
                            }
                            nameScores[col] = score;
                        }

                        int bestPairScore = 0;
                        int bestAdCol = -1;
                        int bestSoyadCol = -1;
                        for (int col = tcCol + 1; col < lastCol; col++)
                        {
                            int leftScore;
                            int rightScore;
                            if (!nameScores.TryGetValue(col, out leftScore)) continue;
                            if (!nameScores.TryGetValue(col + 1, out rightScore)) continue;
                            if (leftScore == 0 || rightScore == 0) continue;
                            int pairScore = leftScore + rightScore;
                            if (pairScore > bestPairScore)
                            {
                                bestPairScore = pairScore;
                                bestAdCol = col;
                                bestSoyadCol = col + 1;
                            }
                        }

                        if (bestPairScore > 0)
                        {
                            if (adCol == -1) adCol = bestAdCol;
                            if (soyadCol == -1) soyadCol = bestSoyadCol;
                        }
                    }

                    using (var conn = DbConnection.GetConnection())
                    {
                        conn.Open();
                        
                        int importedCount = 0;
                        int skippedCount = 0;
                        var gorulenTcListesi = new HashSet<string>();

                        // Veri satırlarını işle
                        for (int row = firstRow; row <= lastRow; row++)
                        {
                            try
                            {
                                string tc = ReadTcFromCell(worksheet.Cell(row, tcCol));
                                if (IsLikelyTc(tc))
                                    gorulenTcListesi.Add(tc);

                                // Öncelik 1: Ayrı Ad ve Soyad kolonları tespit edilmişse ve hücreleri doluysa
                                string secilenAd = adCol > 0 ? (worksheet.Cell(row, adCol).GetValue<string>()?.Trim() ?? "") : "";
                                string secilenSoyad = soyadCol > 0 ? (worksheet.Cell(row, soyadCol).GetValue<string>()?.Trim() ?? "") : "";
                                string secilenAdSoyadBirlikte = adSoyadCol > 0 ? (worksheet.Cell(row, adSoyadCol).GetValue<string>()?.Trim() ?? "") : "";

                                string firstName = "";
                                string lastName = "";
                                bool basariliBulundu = false;

                                if (HasLetter(secilenAd) && HasLetter(secilenSoyad) && !IsNameNoise(secilenAd) && !IsNameNoise(secilenSoyad))
                                {
                                    firstName = secilenAd;
                                    lastName = secilenSoyad;
                                    basariliBulundu = true;
                                }
                                else if (HasLetter(secilenAd) && !IsNameNoise(secilenAd) && (IsNameNoise(secilenSoyad) || !HasLetter(secilenSoyad)))
                                {
                                    SplitFullName(secilenAd, out firstName, out lastName);
                                    basariliBulundu = HasLetter(firstName);
                                }

                                // Öncelik 2: Başlık yanlış yakalansa bile TC'nin solundaki iki hücreyi dene
                                if (!basariliBulundu)
                                {
                                    string adNearTc = tcCol - 2 >= firstCol ? (worksheet.Cell(row, tcCol - 2).GetValue<string>() ?? "").Trim() : "";
                                    string soyadNearTc = tcCol - 1 >= firstCol ? (worksheet.Cell(row, tcCol - 1).GetValue<string>() ?? "").Trim() : "";

                                    if (HasLetter(adNearTc) && HasLetter(soyadNearTc) && !IsNameNoise(adNearTc) && !IsNameNoise(soyadNearTc))
                                    {
                                        firstName = adNearTc;
                                        lastName = soyadNearTc;
                                        basariliBulundu = true;
                                    }
                                }

                                // Öncelik 3: Ad Soyad birleşik kolon tespit edilmişse
                                if (!basariliBulundu && HasLetter(secilenAdSoyadBirlikte))
                                {
                                    // BUÜ şablonunda AD SOYAD kolonunun hemen sağında gerçek soyad olabiliyor.
                                    string soyadSagHucre = "";
                                    if (adSoyadCol > 0 && adSoyadCol + 1 <= lastCol)
                                    {
                                        int rightCol = adSoyadCol + 1;
                                        if (rightCol != tcCol && rightCol != ibanCol && rightCol != gunCol && !dayColumnMap.Values.Contains(rightCol))
                                        {
                                            soyadSagHucre = (worksheet.Cell(row, rightCol).GetValue<string>() ?? "").Trim();
                                        }
                                    }

                                    bool rightIsIban = soyadSagHucre.Replace(" ", "").ToUpperInvariant().StartsWith("TR")
                                                       && soyadSagHucre.Replace(" ", "").Length >= 26;
                                    if (HasLetter(soyadSagHucre) && !IsNameNoise(soyadSagHucre) && !rightIsIban && !IsLikelyTc(CleanDigits(soyadSagHucre)))
                                    {
                                        firstName = secilenAdSoyadBirlikte;
                                        lastName = soyadSagHucre;
                                    }
                                    else
                                    {
                                        SplitFullName(secilenAdSoyadBirlikte, out firstName, out lastName);
                                    }
                                    basariliBulundu = true;
                                }

                                // Öncelik 4: Kolonlardan bulamadıysak metin hücrelerini analiz et (Fallback)
                                if (!basariliBulundu)
                                {
                                    var metinHucreleri = new List<string>();
                                    for (int c = firstCol; c <= lastCol; c++)
                                    {
                                        if (c == tcCol || c == ibanCol || c == gunCol || c == adCol || c == soyadCol || c == adSoyadCol) continue;
                                        if (dayColumnMap.Values.Contains(c)) continue;

                                        string v = (worksheet.Cell(row, c).GetValue<string>() ?? "").Trim();
                                        if (string.IsNullOrWhiteSpace(v)) continue;

                                        if (IsNameNoise(v)) continue;

                                        if (v.All(char.IsDigit)) continue;
                                        if (v.Replace(" ", "").ToUpperInvariant().StartsWith("TR") && v.Replace(" ", "").Length >= 26) continue;
                                        if (InputValidator.IsValidTCNumber(CleanDigits(v))) continue;

                                        if (HasLetter(v))
                                            metinHucreleri.Add(v);
                                    }

                                    if (metinHucreleri.Count > 0)
                                    {
                                        string adSoyadMetin = string.Join(" ", metinHucreleri).Trim();
                                        var np = adSoyadMetin.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
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
                                }

                                if (IsMonthLike(NormalizeNameValue(lastName)))
                                {
                                    string tempFirst;
                                    string tempLast;
                                    SplitFullName(firstName, out tempFirst, out tempLast);
                                    if (!string.IsNullOrWhiteSpace(tempLast) && !IsMonthLike(NormalizeNameValue(tempLast)))
                                    {
                                        firstName = tempFirst;
                                        lastName = tempLast;
                                    }
                                    else
                                    {
                                        lastName = "";
                                    }
                                }

                                string adSoyad = (firstName + " " + lastName).Trim();

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

                                DateTime? iseGirisTarihi = null;
                                if (iseGirisCol > 0)
                                {
                                    var dateCell = worksheet.Cell(row, iseGirisCol);
                                    if (dateCell.TryGetValue(out DateTime d))
                                    {
                                        iseGirisTarihi = d;
                                    }
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
                                if (!IsLikelyTc(tc) ||
                                    IsExcelHeaderLike(adSoyadUpper) ||
                                    !HasLetter(adSoyad))
                                {
                                    skippedCount++;
                                    continue;
                                }

                                // Puantaj tablosuna ekle/güncelle
                                string queryPuantaj = @"INSERT OR REPLACE INTO puantaj 
                                    (p_tc, p_ad_soyad, p_ad, p_soyad, p_iban, p_gun_detaylari, p_calistigi_gun_sayisi, p_ise_baslama_tarihi) 
                                    VALUES (@tc, @adsoy, @ad, @soy, @iban, @detay, @gun, COALESCE(@iseGiris, DATE('now')))";

                                using (var rowTx = conn.BeginTransaction())
                                {
                                    try
                                    {
                                        using (var cmd = new SqliteCommand(queryPuantaj, conn, rowTx))
                                        {
                                            cmd.Parameters.AddWithValue("@tc", tc);
                                            cmd.Parameters.AddWithValue("@adsoy", adSoyad);
                                            cmd.Parameters.AddWithValue("@ad", firstName);
                                            cmd.Parameters.AddWithValue("@soy", lastName);
                                            cmd.Parameters.AddWithValue("@iban", iban);
                                            cmd.Parameters.AddWithValue("@detay", (object)gunDetaylari ?? DBNull.Value);
                                            cmd.Parameters.AddWithValue("@gun", gunSayisi);
                                            cmd.Parameters.AddWithValue("@iseGiris", (object)iseGirisTarihi ?? DBNull.Value);
                                            cmd.ExecuteNonQuery();
                                        }

                                        // Program katılımcıları tablosuna ekle/güncelle (kampüs bilgisiyle)
                                        string queryKatilimci = @"INSERT OR REPLACE INTO program_katilimcilari 
                                            (pk_tc, pk_ad_soyad, pk_ad, pk_soyad, pk_iban_no, pk_gorev_yeri, pk_is_baslama_tarihi) 
                                            VALUES (@tc, @adsoy, @ad, @soy, @iban, @kampus, COALESCE(@iseGiris, DATE('now')))";

                                        using (var cmd = new SqliteCommand(queryKatilimci, conn, rowTx))
                                        {
                                            cmd.Parameters.AddWithValue("@tc", tc);
                                            cmd.Parameters.AddWithValue("@adsoy", adSoyad);
                                            cmd.Parameters.AddWithValue("@ad", firstName);
                                            cmd.Parameters.AddWithValue("@soy", lastName);
                                            cmd.Parameters.AddWithValue("@iban", iban);
                                            cmd.Parameters.AddWithValue("@kampus", kampus);
                                            cmd.Parameters.AddWithValue("@iseGiris", (object)iseGirisTarihi ?? DBNull.Value);
                                            cmd.ExecuteNonQuery();
                                        }

                                        rowTx.Commit();
                                    }
                                    catch
                                    {
                                        rowTx.Rollback();
                                        throw;
                                    }
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

                        // Bu dosyada görülen tüm geçerli TC'ler için kampüs eşleşmesini zorunlu olarak düzelt
                        foreach (var tcFix in gorulenTcListesi)
                        {
                            string fixKatilimciSql = @"INSERT OR REPLACE INTO program_katilimcilari
                                (pk_tc, pk_ad_soyad, pk_ad, pk_soyad, pk_iban_no, pk_gorev_yeri, pk_is_baslama_tarihi)
                                SELECT p.p_tc, p.p_ad_soyad, p.p_ad, p.p_soyad, p.p_iban, @kampus, p.p_ise_baslama_tarihi
                                FROM puantaj p
                                WHERE p.p_tc = @tc";

                            using (var fixCmd = new SqliteCommand(fixKatilimciSql, conn))
                            {
                                fixCmd.Parameters.AddWithValue("@tc", tcFix);
                                fixCmd.Parameters.AddWithValue("@kampus", kampus);
                                fixCmd.ExecuteNonQuery();
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

            int tcCol = -1, adSoyadCol = -1, adCol = -1, soyadCol = -1, ibanCol = -1, gunCol = -1, iseGirisCol = -1;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                string col = (dt.Columns[i].ColumnName ?? "").ToUpperInvariant();
                if (tcCol == -1 && (col.Contains("TC") || col.Contains("KIMLIK") || col.Contains("KİMLİK") || col.Contains("T.C"))) tcCol = i;
                else if (adSoyadCol == -1 && ((col.Contains("AD") && col.Contains("SOYAD")) || col.Contains("ADSOYAD") || col.Contains("ADI SOYADI"))) adSoyadCol = i;
                else if (adCol == -1 && (col == "AD" || col.Contains("ADI") || col.Contains("KATILIMCI"))) adCol = i;
                else if (soyadCol == -1 && (col.Contains("SOYAD") || col.Contains("BILGILERI") || col.Contains("BİLGİLERİ"))) soyadCol = i;
                else if (ibanCol == -1 && col.Contains("IBAN")) ibanCol = i;
                else if (gunCol == -1 && (col.Contains("GUN") || col.Contains("GÜN"))) gunCol = i;
                else if (iseGirisCol == -1 && (col.Contains("ISE GIRIS") || col.Contains("İŞE GİRİŞ") || col.Contains("BASLAMA TARIHI") || col.Contains("BAŞLAMA TARİHİ"))) iseGirisCol = i;
            }

            if (tcCol == -1) tcCol = 0;
            // Removed adSoyadCol default guessing to avoid blindly picking "AD" column
            if (ibanCol == -1) ibanCol = Math.Min(2, dt.Columns.Count - 1);
            if (gunCol == -1) gunCol = Math.Min(3, dt.Columns.Count - 1);

            int importedCount = 0;
            int skippedCount = 0;
            bool HasLetter(string s) => !string.IsNullOrWhiteSpace(s) && s.Any(char.IsLetter);
            bool IsLikelyTc(string s) => !string.IsNullOrWhiteSpace(s) && s.Length == 11 && s.All(char.IsDigit) && s[0] != '0';
            string NormalizeOleHeader(string s)
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
            bool IsNameNoise(string text)
            {
                if (string.IsNullOrWhiteSpace(text)) return true;
                string t = NormalizeOleHeader(text);
                return IsExcelHeaderLike(t)
                       || t.Contains("KAMPUS")
                       || t == "GUVENLIK" || t == "TEMIZLIK" || t == "IDARI" || t == "BILISIM" || t == "TEKNIK" || t == "AKADEMIK"
                       || IsMonthLike(t)
                       || t == "X" || t == "I" || t == "R";
            }

            bool IsNameCandidate(string value)
            {
                if (string.IsNullOrWhiteSpace(value)) return false;
                string v = value.Trim();
                if (IsNameNoise(v)) return false;
                if (v.All(char.IsDigit)) return false;
                if (v.Replace(" ", "").ToUpperInvariant().StartsWith("TR") && v.Replace(" ", "").Length >= 26) return false;
                if (IsLikelyTc(new string(v.Where(char.IsDigit).ToArray()))) return false;
                return HasLetter(v);
            }

            if ((adCol == -1 || soyadCol == -1) && tcCol > -1)
            {
                var nameScores = new Dictionary<int, int>();
                for (int col = tcCol + 1; col < dt.Columns.Count; col++)
                {
                    if (col == tcCol || col == ibanCol || col == gunCol || col == adSoyadCol) continue;
                    int score = 0;
                    for (int row = 0; row < Math.Min(dt.Rows.Count, 200); row++)
                    {
                        string v = (dt.Rows[row][col]?.ToString() ?? "").Trim();
                        if (IsNameCandidate(v)) score++;
                    }
                    nameScores[col] = score;
                }

                int bestPairScore = 0;
                int bestAdCol = -1;
                int bestSoyadCol = -1;
                for (int col = tcCol + 1; col < dt.Columns.Count - 1; col++)
                {
                    int leftScore;
                    int rightScore;
                    if (!nameScores.TryGetValue(col, out leftScore)) continue;
                    if (!nameScores.TryGetValue(col + 1, out rightScore)) continue;
                    if (leftScore == 0 || rightScore == 0) continue;
                    int pairScore = leftScore + rightScore;
                    if (pairScore > bestPairScore)
                    {
                        bestPairScore = pairScore;
                        bestAdCol = col;
                        bestSoyadCol = col + 1;
                    }
                }

                if (bestPairScore > 0)
                {
                    if (adCol == -1) adCol = bestAdCol;
                    if (soyadCol == -1) soyadCol = bestSoyadCol;
                }
            }

            using (var conn = DbConnection.GetConnection())
            {
                conn.Open();
                var gorulenTcListesi = new HashSet<string>();

                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        string tc = (row[tcCol]?.ToString() ?? "").Trim();
                        tc = new string(tc.Where(char.IsDigit).ToArray());
                        if (IsLikelyTc(tc))
                            gorulenTcListesi.Add(tc);

                        string iban = (row[ibanCol]?.ToString() ?? "").Trim();

                        // Öncelik 1: Ayrı Ad ve Soyad kolonları tespit edilmişse
                        string secilenAd = adCol > -1 ? (row[adCol]?.ToString() ?? "").Trim() : "";
                        string secilenSoyad = soyadCol > -1 ? (row[soyadCol]?.ToString() ?? "").Trim() : "";
                        string secilenAdSoyadBirlikte = adSoyadCol > -1 ? (row[adSoyadCol]?.ToString() ?? "").Trim() : "";

                        string firstName = "";
                        string lastName = "";
                        bool basariliBulundu = false;

                        if (HasLetter(secilenAd) && HasLetter(secilenSoyad) && !IsNameNoise(secilenAd) && !IsNameNoise(secilenSoyad))
                        {
                            firstName = secilenAd;
                            lastName = secilenSoyad;
                            basariliBulundu = true;
                        }
                        else if (HasLetter(secilenAd) && !IsNameNoise(secilenAd) && (IsNameNoise(secilenSoyad) || !HasLetter(secilenSoyad)))
                        {
                            SplitFullName(secilenAd, out firstName, out lastName);
                            basariliBulundu = HasLetter(firstName);
                        }

                        // Öncelik 2: Başlık yanlış yakalansa bile TC'nin solundaki iki hücreyi dene
                        if (!basariliBulundu)
                        {
                            string adNearTc = tcCol - 2 >= 0 ? (row[tcCol - 2]?.ToString() ?? "").Trim() : "";
                            string soyadNearTc = tcCol - 1 >= 0 ? (row[tcCol - 1]?.ToString() ?? "").Trim() : "";
                            if (HasLetter(adNearTc) && HasLetter(soyadNearTc) && !IsNameNoise(adNearTc) && !IsNameNoise(soyadNearTc))
                            {
                                firstName = adNearTc;
                                lastName = soyadNearTc;
                                basariliBulundu = true;
                            }
                        }

                        // Öncelik 3: Ad Soyad birleşik kolon tespit edilmişse
                        if (!basariliBulundu && HasLetter(secilenAdSoyadBirlikte))
                        {
                            // BUÜ şablonunda AD SOYAD kolonunun hemen sağında gerçek soyad olabiliyor.
                            string soyadSagHucre = "";
                            if (adSoyadCol > -1 && adSoyadCol + 1 < dt.Columns.Count)
                            {
                                int rightCol = adSoyadCol + 1;
                                if (rightCol != tcCol && rightCol != ibanCol && rightCol != gunCol)
                                {
                                    soyadSagHucre = (row[rightCol]?.ToString() ?? "").Trim();
                                }
                            }

                            bool rightIsIban = soyadSagHucre.Replace(" ", "").ToUpperInvariant().StartsWith("TR")
                                               && soyadSagHucre.Replace(" ", "").Length >= 26;
                            if (HasLetter(soyadSagHucre) && !IsNameNoise(soyadSagHucre) && !rightIsIban)
                            {
                                firstName = secilenAdSoyadBirlikte;
                                lastName = soyadSagHucre;
                            }
                            else
                            {
                                SplitFullName(secilenAdSoyadBirlikte, out firstName, out lastName);
                            }
                            basariliBulundu = true;
                        }

                        // Öncelik 4: Fallback metin hücrelerini analiz et
                        if (!basariliBulundu)
                        {
                            var metinler = new List<string>();
                            for (int c = 0; c < dt.Columns.Count; c++)
                            {
                                if (c == tcCol || c == ibanCol || c == gunCol || c == adCol || c == soyadCol || c == adSoyadCol) continue;
                                string v = (row[c]?.ToString() ?? "").Trim();
                                if (string.IsNullOrWhiteSpace(v)) continue;
                                if (IsNameNoise(v)) continue;
                                if (v.All(char.IsDigit)) continue;
                                if (HasLetter(v)) metinler.Add(v);
                            }

                            if (metinler.Count > 0)
                            {
                                string adSoyadMetin = string.Join(" ", metinler).Trim();
                                var np = adSoyadMetin.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
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
                        }

                        if (IsMonthLike(NormalizeNameValue(lastName)))
                        {
                            string tempFirst;
                            string tempLast;
                            SplitFullName(firstName, out tempFirst, out tempLast);
                            if (!string.IsNullOrWhiteSpace(tempLast) && !IsMonthLike(NormalizeNameValue(tempLast)))
                            {
                                firstName = tempFirst;
                                lastName = tempLast;
                            }
                            else
                            {
                                lastName = "";
                            }
                        }

                        string adSoyad = (firstName + " " + lastName).Trim();

                        if (tc.ToUpperInvariant().Contains("TC") || adSoyad.ToUpperInvariant().Contains("AD SOYAD"))
                        {
                            skippedCount++;
                            continue;
                        }

                        // Geçersiz/başlık satırlarını atla
                        string adSoyadUpper = (adSoyad ?? "").ToUpperInvariant();
                        if (!InputValidator.IsValidTCNumber(tc) ||
                            IsExcelHeaderLike(adSoyadUpper) ||
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

                        DateTime? iseGirisTarihi = null;
                        if (iseGirisCol > -1)
                        {
                            var tv = row[iseGirisCol];
                            if (tv != null && tv != DBNull.Value)
                            {
                                if (tv is DateTime dtVal) iseGirisTarihi = dtVal;
                                else if (DateTime.TryParse(tv.ToString(), out DateTime parsedDate)) iseGirisTarihi = parsedDate;
                            }
                        }

                        string queryPuantaj = @"INSERT OR REPLACE INTO puantaj 
                            (p_tc, p_ad_soyad, p_ad, p_soyad, p_iban, p_calistigi_gun_sayisi, p_ise_baslama_tarihi) 
                            VALUES (@tc, @adsoy, @ad, @soy, @iban, @gun, COALESCE(@iseGiris, DATE('now')))";

                        using (var rowTx = conn.BeginTransaction())
                        {
                            try
                            {
                                using (var cmd = new SqliteCommand(queryPuantaj, conn, rowTx))
                                {
                                    cmd.Parameters.AddWithValue("@tc", tc);
                                    cmd.Parameters.AddWithValue("@adsoy", adSoyad);
                                    cmd.Parameters.AddWithValue("@ad", firstName);
                                    cmd.Parameters.AddWithValue("@soy", lastName);
                                    cmd.Parameters.AddWithValue("@iban", iban);
                                    cmd.Parameters.AddWithValue("@gun", gunSayisi);
                                    cmd.Parameters.AddWithValue("@iseGiris", (object)iseGirisTarihi ?? DBNull.Value);
                                    cmd.ExecuteNonQuery();
                                }

                                string queryKatilimci = @"INSERT OR REPLACE INTO program_katilimcilari 
                                    (pk_tc, pk_ad_soyad, pk_ad, pk_soyad, pk_iban_no, pk_gorev_yeri, pk_is_baslama_tarihi) 
                                    VALUES (@tc, @adsoy, @ad, @soy, @iban, @kampus, COALESCE(@iseGiris, DATE('now')))";

                                using (var cmd = new SqliteCommand(queryKatilimci, conn, rowTx))
                                {
                                    cmd.Parameters.AddWithValue("@tc", tc);
                                    cmd.Parameters.AddWithValue("@adsoy", adSoyad);
                                    cmd.Parameters.AddWithValue("@ad", firstName);
                                    cmd.Parameters.AddWithValue("@soy", lastName);
                                    cmd.Parameters.AddWithValue("@iban", iban);
                                    cmd.Parameters.AddWithValue("@kampus", kampus);
                                    cmd.Parameters.AddWithValue("@iseGiris", (object)iseGirisTarihi ?? DBNull.Value);
                                    cmd.ExecuteNonQuery();
                                }

                                rowTx.Commit();
                            }
                            catch
                            {
                                rowTx.Rollback();
                                throw;
                            }
                        }

                        importedCount++;
                    }
                    catch
                    {
                        skippedCount++;
                    }
                }

                // Bu dosyada görülen tüm geçerli TC'ler için kampüs eşleşmesini zorunlu olarak düzelt
                foreach (var tcFix in gorulenTcListesi)
                {
                    string fixKatilimciSql = @"INSERT OR REPLACE INTO program_katilimcilari
                        (pk_tc, pk_ad_soyad, pk_ad, pk_soyad, pk_iban_no, pk_gorev_yeri, pk_is_baslama_tarihi)
                        SELECT p.p_tc, p.p_ad_soyad, p.p_ad, p.p_soyad, p.p_iban, @kampus, p.p_ise_baslama_tarihi
                        FROM puantaj p
                        WHERE p.p_tc = @tc";

                    using (var fixCmd = new SqliteCommand(fixKatilimciSql, conn))
                    {
                        fixCmd.Parameters.AddWithValue("@tc", tcFix);
                        fixCmd.Parameters.AddWithValue("@kampus", kampus);
                        fixCmd.ExecuteNonQuery();
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


