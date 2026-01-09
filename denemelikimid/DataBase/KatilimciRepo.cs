using System;
using System.Data;
using MySql.Data.MySqlClient;

namespace denemelikimid.DataBase
{
    public class KatilimciRepo
    {
        private readonly DbRepository _db = new DbRepository();

        public DataTable Listele()
        {
            return _db.GetByQuery(
                @"SELECT * FROM program_katilimcilari ");
        }

        public void Ekle(double tc_no, String ad_soyad, String iban_no, String gorev_yeri, DateTime ise_giris_tarihi, DateTime isten_cikis_tarihi)
        {
            _db.Execute(
               @"INSERT INTO program_katilimcilari " +
                "(pk_tc, pk_ad_soyad, pk_iban, pk_gorev_yeri, pk_ise_baslama_tarihi, pk_isten_ayrilma_tarihi) " +
                "VALUES (@tc, @adsoyad, @iban, @gorevyeri, @isebaslama, @istenayrilma)",
                new MySqlParameter("@tc", tc_no),
                new MySqlParameter("@adsoyad", ad_soyad),
                new MySqlParameter("@iban", iban_no),
                new MySqlParameter("@gorevyeri", gorev_yeri),
                new MySqlParameter("@isebaslama", ise_giris_tarihi),
                new MySqlParameter("@istenayrilma", isten_cikis_tarihi));
        }

        public void Guncelle(double tc_no, String ad_soyad, String iban_no, String gorev_yeri, DateTime ise_giris_tarihi, DateTime isten_cikis_tarihi)
        {
            _db.Execute(
               @"UPDATE program_katilimcilari SET " +
                "pk_ad_soyad = @adsoyad, pk_iban = @iban, pk_gorev_yeri = @gorevyeri, " +
                "pk_ise_baslama_tarihi = @isebaslama, pk_isten_ayrilma_tarihi = @istenayrilma " +
                "WHERE pk_tc = @tc",
                new MySqlParameter("@tc", tc_no),
                new MySqlParameter("@adsoyad", ad_soyad),
                new MySqlParameter("@iban", iban_no),
                new MySqlParameter("@gorevyeri", gorev_yeri),
                new MySqlParameter("@isebaslama", ise_giris_tarihi),
                new MySqlParameter("@istenayrilma", isten_cikis_tarihi));
        }

        public void Sil(double tc_no)
        {
            _db.Execute(
                @"DELETE FROM program_katilimcilari WHERE pk_tc = @tc",
                new MySqlParameter("@tc", tc_no));
        }
    }
}

