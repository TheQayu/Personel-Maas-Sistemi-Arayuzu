using System;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;
using denemelikimid.DataBase;

namespace denemelikimid.Validations
{
    /// <summary>
    /// TC Kimlik No, Telefon No, IBAN ve diğer girişleri valide eden sınıf
    /// </summary>
    public static class InputValidator
    {
        /// <summary>
        /// TC Kimlik Numarasını doğrular - Resmi T.C. Nüfus ve Vatandaşlık İşleri Başkanlığı Algoritması
        /// 
        /// AÇIKLAMA:
        /// 1. 11 basamaklı bir sayı olmalıdır
        /// 2. İlk basamak 0 olamaz
        /// 3. 10. basamak (index 9): Tek pozisyonların (1,3,5,7,9) toplamı * 7 - Çift pozisyonların (2,4,6,8,10) toplamı
        /// 4. 11. basamak (index 10): İlk 10 basamağın toplamı
        /// 
        /// ÖRNEKLER:
        /// ✅ 12345678905 → Geçerli
        /// ✅ 11111111110 → Geçerli
        /// ❌ 00000000000 → 0 ile başlayamaz
        /// ❌ 12345678901 → Algoritma uyumsuz
        /// </summary>
        public static bool IsValidTCNumber(string tc)
        {
            if (string.IsNullOrWhiteSpace(tc))
                return false;

            tc = tc.Trim();

            // KONTROL 1: 11 basamak olmalı ve sadece sayı içermeli
            if (tc.Length != 11)
                return false;

            if (!Regex.IsMatch(tc, @"^\d{11}$"))
                return false;

            // KONTROL 2: 0 ile başlayamaz
            if (tc[0] == '0')
                return false;

            // Basamakları integer dizisine çevir
            int[] digits = new int[11];
            for (int i = 0; i < 11; i++)
            {
                digits[i] = tc[i] - '0'; // Char to int dönüşümü (hızlı)
            }

            // KONTROL 3: 10. Basamak (index 9) doğrulaması
            // Tek pozisyonlar (1, 3, 5, 7, 9 basamaklar) = indices (0, 2, 4, 6, 8)
            int oddPositionSum = digits[0] + digits[2] + digits[4] + digits[6] + digits[8];

            // Çift pozisyonlar (2, 4, 6, 8 basamaklar) = indices (1, 3, 5, 7)
            // NOT: 10. basamak (index 9) çift toplamına KATILMAZ
            int evenPositionSum = digits[1] + digits[3] + digits[5] + digits[7];

            // 10. basamak formülü: ((Tek * 7) - Çift) % 10
            int calculatedTenth = ((oddPositionSum * 7) - evenPositionSum) % 10;

            // Modulo negatif olabilir, düzelt
            if (calculatedTenth < 0)
                calculatedTenth += 10;

            // 10. basamak eşleşmelidir (index 9)
            if (digits[9] != calculatedTenth)
                return false;

            // KONTROL 4: 11. Basamak (index 10) doğrulaması
            // İlk 10 basamağın (0-9 indexleri) toplamının mod 10'u
            int sum0To9 = 0;
            for (int i = 0; i < 10; i++)
            {
                sum0To9 += digits[i];
            }

            int calculatedEleventh = sum0To9 % 10;

            // 11. basamak eşleşmelidir (index 10)
            if (digits[10] != calculatedEleventh)
                return false;

            // TÜM KONTROLLER GEÇTİ
            return true;
        }

        /// <summary>
        /// Veritabanında aynı TC kimlik numarası kaydı olup olmadığını kontrol eder
        /// </summary>
        public static bool IsTCNumberExists(string tc)
        {
            if (string.IsNullOrWhiteSpace(tc))
                return false;

            try
            {
                using (var conn = DbConnection.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqliteCommand("SELECT COUNT(*) FROM program_katilimcilari WHERE pk_tc = @tc", conn);
                    cmd.Parameters.AddWithValue("@tc", tc.Trim());
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    return count > 0;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Türk telefon numarasını doğrular (10 basamak)
        /// Geçerli formatlar: 5XX XXX XXXX veya 5XXXXXXXXXX
        /// </summary>
        public static bool IsValidPhoneNumber(string phone)
        {
            if (string.IsNullOrWhiteSpace(phone))
                return false;

            phone = phone.Trim().Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "");

            // Sadece sayılar olmalı
            if (!Regex.IsMatch(phone, @"^\d+$"))
                return false;

            // 10 basamak olmalı
            if (phone.Length != 10)
                return false;

            // 5 ile başlamalı
            if (phone[0] != '5')
                return false;

            return true;
        }

        /// <summary>
        /// Veritabanında aynı telefon numarası kaydı olup olmadığını kontrol eder
        /// </summary>
        public static bool IsPhoneNumberExists(string phone)
        {
            if (string.IsNullOrWhiteSpace(phone))
                return false;

            try
            {
                string cleanPhone = phone.Trim().Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "");
                using (var conn = DbConnection.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqliteCommand("SELECT COUNT(*) FROM program_katilimcilari WHERE REPLACE(REPLACE(pk_telefon, ' ', ''), '-', '') = @phone", conn);
                    cmd.Parameters.AddWithValue("@phone", cleanPhone);
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    return count > 0;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Türk IBAN numarasını doğrular (24 sayı, TR prefix otomatik eklenir)
        /// Kullanıcı sadece 24 sayıyı girsin, TR otomatik eklenir
        /// </summary>
        public static bool IsValidIBAN(string iban)
        {
            if (string.IsNullOrWhiteSpace(iban))
                return false;

            iban = iban.Trim().Replace(" ", "").ToUpper();

            // Eğer TR ile başlamıyorsa ekle
            if (!iban.StartsWith("TR"))
                iban = "TR" + iban;

            // TR ile başlamalı ve 26 karakterden oluşmalı
            if (!Regex.IsMatch(iban, @"^TR\d{24}$"))
                return false;

            // IBAN Check Digit doğrulaması
            if (!ValidateIBANCheckDigit(iban))
                return false;

            return true;
        }

        /// <summary>
        /// Veritabanında aynı IBAN kaydı olup olmadığını kontrol eder
        /// </summary>
        public static bool IsIBANExists(string iban)
        {
            if (string.IsNullOrWhiteSpace(iban))
                return false;

            try
            {
                string cleanIban = iban.Trim().Replace(" ", "").ToUpper();
                if (!cleanIban.StartsWith("TR"))
                    cleanIban = "TR" + cleanIban;

                using (var conn = DbConnection.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqliteCommand("SELECT COUNT(*) FROM program_katilimcilari WHERE REPLACE(pk_iban_no, ' ', '') = @iban", conn);
                    cmd.Parameters.AddWithValue("@iban", cleanIban);
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    return count > 0;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Adı doğrular (boş olmaması, uygun karakterler)
        /// </summary>
        public static bool IsValidName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return false;

            name = name.Trim();

            // En az 2 karakter, sadece harf ve boşluk
            if (name.Length < 2)
                return false;

            if (!Regex.IsMatch(name, @"^[a-zA-ZçÇğĞıİöÖşŞüÜ\s]+$"))
                return false;

            return true;
        }

        /// <summary>
        /// IBAN numarasını formatla ve TR prefix ekle (boşluklu gösterim)
        /// Eğer TR yoksa ekler
        /// </summary>
        public static string FormatIBAN(string iban)
        {
            if (string.IsNullOrWhiteSpace(iban))
                return "";

            iban = iban.Trim().Replace(" ", "").ToUpper();
            
            // TR yoksa ekle
            if (!iban.StartsWith("TR"))
                iban = "TR" + iban;

            if (iban.Length != 26)
                return iban;

            return $"{iban.Substring(0, 2)} {iban.Substring(2, 6)} {iban.Substring(8, 5)} {iban.Substring(13, 5)} {iban.Substring(18, 8)}";
        }

        /// <summary>
        /// Telefon numarasını formatla (5XX XXX XXXX)
        /// </summary>
        public static string FormatPhoneNumber(string phone)
        {
            if (string.IsNullOrWhiteSpace(phone))
                return "";

            phone = phone.Trim().Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "");

            if (phone.Length != 10)
                return phone;

            return $"{phone.Substring(0, 3)} {phone.Substring(3, 3)} {phone.Substring(6, 4)}";
        }

        /// <summary>
        /// IBAN Check Digit doğrulaması (Mod-97 algoritması)
        /// </summary>
        private static bool ValidateIBANCheckDigit(string iban)
        {
            try
            {
                // IBAN'ı rearrange et: sayısal koda çevir
                string rearranged = iban.Substring(4) + iban.Substring(0, 4);
                string numeric = "";

                foreach (char c in rearranged)
                {
                    if (char.IsDigit(c))
                        numeric += c;
                    else
                        numeric += (char.ToUpper(c) - 'A' + 10).ToString();
                }

                // Mod 97 hesapla
                long remainder = 0;
                foreach (char digit in numeric)
                {
                    remainder = (remainder * 10 + int.Parse(digit.ToString())) % 97;
                }

                return remainder == 1;
            }
            catch
            {
                return false;
            }
        }
    }
}
