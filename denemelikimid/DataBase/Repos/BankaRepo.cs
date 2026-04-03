using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using Microsoft.Data.Sqlite;
using denemelikimid.DataBase.Models;
namespace denemelikimid.DataBase.Repos
{
    internal class BankaRepo
    {
        private readonly DbRepository _db = new DbRepository();

        public DataTable Listele()
        {
            return _db.GetByQuery(
                @"SELECT * FROM banka_listesi ");
        }

        public void Ekle(BankaModel model)
        {
            string query = @"INSERT INTO banka_listesi (bl_iban_no, bl_tutar, bl_ad_soyad, bl_tc, bl_aciklama)
                            VALUES (@iban, @tutar, @ad_soyad, @tc, @aciklama)0";
            _db.Execute(query,
             new SqliteParameter("@iban", model.Bnk_iban),
             new SqliteParameter("@tutar", model.Bnk_tutar),
             new SqliteParameter("@ad_soyad", model.Bnk_ad_soyad),
             new SqliteParameter("@tc", model.Bnk_tc),
             new SqliteParameter("@aciklama", model.Bnk_aciklama)


           );
        }
        public void Sil(BankaModel model)
        {
            _db.Execute(
                @"DELETE FROM banka_listesi WHERE bl_id = @id",
                new SqliteParameter("@id", model.Bnk_id));
        }

        public void Guncelle(BankaModel model)
        {
            string query = @"UPDATE banka_listesi SET 
                bl_iban_no = @iban,
                bl_tutar = @tutar,
                bl_ad_soyad = @ad_soyad,
                bl_tc = @tc,
                bl_aciklama = @aciklama
                WHERE bl_id = @id";
            _db.Execute(query,
             new SqliteParameter("@iban", model.Bnk_iban),
             new SqliteParameter("@tutar", model.Bnk_tutar),
             new SqliteParameter("@ad_soyad", model.Bnk_ad_soyad),
             new SqliteParameter("@tc", model.Bnk_tc),
             new SqliteParameter("@aciklama", model.Bnk_aciklama),
             new SqliteParameter("@id", model.Bnk_id)
           );
        }
    }
}
