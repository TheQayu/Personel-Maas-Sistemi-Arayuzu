using System;
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
                OpenFileDialog ofd = new OpenFileDialog
                {
                    Filter = "Excel Dosyaları|*.xlsx;*.xls",
                    Title = "Puantaj Dosyasını Seçin"
                };

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook(ofd.FileName))
                        {
                            var worksheet = workbook.Worksheet(1);
                            var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

                            using (var conn =
                                   new MySqlConnection("Server=localhost;Database=iskur;Uid=yeniAdmin;Pwd=1234;"))
                            {
                                conn.Open();
                                new MySqlCommand("TRUNCATE TABLE puantaj", conn).ExecuteNonQuery();

                                foreach (var row in rows)
                                {
                                    string tc = row.Cell(1).GetValue<string>();
                                    string adSoyad = row.Cell(2).GetValue<string>();
                                    string iban = row.Cell(3).GetValue<string>();
                                    int gunSayisi = 0;
                                    int.TryParse(row.Cell(4).GetValue<string>(), out gunSayisi);

                                    string query = @"INSERT INTO puantaj 
                                           (p_tc, p_ad_soyad, p_iban, p_calistigi_gun_sayisi, p_ise_baslama_tarihi) 
                                           VALUES (@tc, @ad, @iban, @gun, CURDATE())";

                                    using (var cmd = new MySqlCommand(query, conn))
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
                        MessageBox.Show("✅ Puantaj listesi başarıyla yüklendi!", "İşlem Başarılı",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Hata: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
    }
}


