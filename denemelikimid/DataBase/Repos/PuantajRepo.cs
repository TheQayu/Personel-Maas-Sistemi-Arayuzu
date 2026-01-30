using denemelikimid.DataBase.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Text;

namespace denemelikimid.DataBase.Repos
{
    internal class PuantajRepo
    {
        private readonly DbRepository _db = new DbRepository();
        public DataTable Listele()
        {
            return _db.GetByQuery(
                @"SELECT * FROM puantaj ");
        }
        public class PuantajRepository
        {
            private readonly DbContext _db;

 
            public void Ekle(PuantajModel model)
            {
                string query = @"
            INSERT INTO puantaj (
                p_iban, p_tc, p_ad_soyad, p_calistigi_gun_sayisi, 
                p_devamsizlik, p_yillik_izin, p_ise_baslama_tarihi, p_isten_ayrilma_tarihi
            ) 
            VALUES (
                @iban, @tc, @ad_soyad, @cal_gun, 
                @devam, @y_izin, @ise_bas, @ist_ayril
            )";

                _db.Execute(query,
                    new MySqlParameter("@iban", model.Pnt_iban),
                    new MySqlParameter("@tc", model.Pnt_tc),
                    new MySqlParameter("@ad_soyad", model.Pnt_ad_soyad),
                    new MySqlParameter("@cal_gun", model.Pnt_calisilan_gun_sayisi),
                    new MySqlParameter("@devam", model.Pnt_devamsizlik),
                    new MySqlParameter("@y_izin", model.Pnt_yillik_izin),
                    new MySqlParameter("@ise_bas", model.Pnt_ise_baslama_tarihi),
                    new MySqlParameter("@ist_ayril", model.Pnt_isten_ayrilma_tarihi));
            }

            
            public void Guncelle(PuantajModel model)
            {
                string query = @"
            UPDATE puantaj SET 
                p_iban = @iban, 
                p_tc = @tc, 
                p_ad_soyad = @ad_soyad, 
                p_calistigi_gun_sayisi = @cal_gun, 
                p_devamsizlik = @devam, 
                p_yillik_izin = @y_izin, 
                p_ise_baslama_tarihi = @ise_bas, 
                p_isten_ayrilma_tarihi = @ist_ayril
            WHERE idpuantaj = @id";

                _db.Execute(query,
                    new MySqlParameter("@id", model.Pnt_id),
                    new MySqlParameter("@iban", model.Pnt_iban),
                    new MySqlParameter("@tc", model.Pnt_tc),
                    new MySqlParameter("@ad_soyad", model.Pnt_ad_soyad),
                    new MySqlParameter("@cal_gun", model.Pnt_calisilan_gun_sayisi),
                    new MySqlParameter("@devam", model.Pnt_devamsizlik),
                    new MySqlParameter("@y_izin", model.Pnt_yillik_izin),
                    new MySqlParameter("@ise_bas", model.Pnt_ise_baslama_tarihi),
                    new MySqlParameter("@ist_ayril", model.Pnt_isten_ayrilma_tarihi));
            }

            
            public void Sil(int id)
            {
                string query = "DELETE FROM puantaj WHERE idpuantaj = @id";
                _db.Execute(query, new MySqlParameter("@id", id));
            }
        }
    }
}
