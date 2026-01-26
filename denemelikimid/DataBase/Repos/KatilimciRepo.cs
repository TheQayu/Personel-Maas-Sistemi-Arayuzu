using System;
using System.Collections.Generic;
using System.Text;
using MySql.Data.MySqlClient;
using System.Data;
using IskurTakipSistemi.DataBase.Models;

namespace IskurTakipSistemi.DataBase.Repos
{
    internal class KatilimciRepo
    {
        private readonly DbRepository _db = new DbRepository();

        public DataTable Listele()
        {
            return _db.GetByQuery(
                @"SELECT * FROM program_katilimcilari ");
        }
        public void Ekle(KatilimciModel model)
        {
            string query = @"INSERT INTO program_katilimcilari 
                (pk_tc, pk_ad_soyad, pk_iban,pk_gorev_yeri, pk_ise_baslama_tarihi,pk_isten_ayrilma_tarihi) 
                VALUES (@tc, @adsoyad, @iban,@gorevyeri, @isebaslama,@istenayrilma)";
            _db.Execute(query,
            new MySqlParameter("@tc",model.Ktc_tc),
            new MySqlParameter("@adsoyad", model.Ktc_ad_soyad),
            new MySqlParameter("@iban", model.Ktc_iban),
            new MySqlParameter("@gorevyeri", model.Ktc_gorev_yeri),
            new MySqlParameter("@isebaslama", model.Ktc_ise_baslama_tarihi),
            new MySqlParameter("@istenayrilma", model.Ktc_isten_cikma_tarihi)
            );
           
            
        }
        public void Sil(KatilimciModel model)
        {

            _db.Execute(
                @"DELETE FROM program_katilimcilari WHERE pk_id= @id",
                new MySqlParameter("@id",model.Ktc_id));
            
        }
        public void Guncelle(KatilimciModel model)
        {
            string query = @"UPDATE program_katilimcilari SET 
                pk_tc = @tc,
                pk_ad_soyad = @adsoyad,
                pk_iban_no = @iban,
                pk_gorev_yeri = @gorevyeri,
                pk_is_baslama_tarihi = @isebaslama,
                pk_isten_ayrilma_tarihi = @istenayrilma
                WHERE idprogram_katilimcilari = @id";
            _db.Execute(query,
            new MySqlParameter("@tc", model.Ktc_tc),
            new MySqlParameter("@adsoyad", model.Ktc_ad_soyad),
            new MySqlParameter("@iban", model.Ktc_iban),
            new MySqlParameter("@gorevyeri", model.Ktc_gorev_yeri),
            new MySqlParameter("@isebaslama", model.Ktc_ise_baslama_tarihi),
            new MySqlParameter("@istenayrilma", model.Ktc_isten_cikma_tarihi),
            new MySqlParameter("@id", model.Ktc_id)
            );
            
            
        }   
    }
}
