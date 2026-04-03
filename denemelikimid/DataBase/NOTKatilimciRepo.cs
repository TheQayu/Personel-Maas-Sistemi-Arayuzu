using System;
using System.Data;
using Microsoft.Data.Sqlite;

namespace denemelikimid.DataBase
{
    public class NOTKatilimciRepo
    {
        private readonly NOTDbRepository _db = new NOTDbRepository();

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
                new SqliteParameter("@tc", tc_no),
                new SqliteParameter("@adsoyad", ad_soyad),
                new SqliteParameter("@iban", iban_no),
                new SqliteParameter("@gorevyeri", gorev_yeri),
                new SqliteParameter("@isebaslama", ise_giris_tarihi),
                new SqliteParameter("@istenayrilma", isten_cikis_tarihi));
        }

        public void Guncelle(double tc_no, String ad_soyad, String iban_no, String gorev_yeri, DateTime ise_giris_tarihi, DateTime isten_cikis_tarihi)
        {
            _db.Execute(
               @"UPDATE program_katilimcilari SET " +
                "pk_ad_soyad = @adsoyad, pk_iban = @iban, pk_gorev_yeri = @gorevyeri, " +
                "pk_ise_baslama_tarihi = @isebaslama, pk_isten_ayrilma_tarihi = @istenayrilma " +
                "WHERE pk_tc = @tc",
                new SqliteParameter("@tc", tc_no),
                new SqliteParameter("@adsoyad", ad_soyad),
                new SqliteParameter("@iban", iban_no),
                new SqliteParameter("@gorevyeri", gorev_yeri),
                new SqliteParameter("@isebaslama", ise_giris_tarihi),
                new SqliteParameter("@istenayrilma", isten_cikis_tarihi));
        }

        public void Sil(double tc_no)
        {
            _db.Execute(
                @"DELETE FROM program_katilimcilari WHERE pk_tc = @tc",
                new SqliteParameter("@tc", tc_no));
        }
    }
}







