using System;
using System.Linq;
using ClosedXML.Excel;
using System.Windows.Forms;

namespace denemelikimid
{
    public class ExcelAnalyzer
    {
        public static void AnalyzeExcelFile(string filePath)
        {
            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var usedRange = worksheet.RangeUsed();
                    
                    if (usedRange == null)
                    {
                        MessageBox.Show("Dosya boş görünüyor.");
                        return;
                    }

                    string info = $"Dosya: {System.IO.Path.GetFileName(filePath)}\n\n";
                    info += $"Toplam Satır: {usedRange.RowCount()}\n";
                    info += $"Toplam Kolon: {usedRange.ColumnCount()}\n\n";
                    
                    // İlk 10 satırı göster
                    info += "İlk 10 Satır:\n";
                    info += "=".PadRight(80, '=') + "\n";
                    
                    for (int row = 1; row <= Math.Min(10, usedRange.RowCount()); row++)
                    {
                        info += $"Satır {row}: ";
                        for (int col = 1; col <= Math.Min(15, usedRange.ColumnCount()); col++)
                        {
                            var cellValue = worksheet.Cell(row, col).GetValue<string>();
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                info += $"[{col}] {cellValue.Substring(0, Math.Min(20, cellValue.Length))} | ";
                            }
                        }
                        info += "\n";
                    }
                    
                    MessageBox.Show(info, "Excel Dosya Analizi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hata: {ex.Message}\n\nDetay: {ex.StackTrace}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
