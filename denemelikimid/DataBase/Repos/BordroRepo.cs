using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using MySql.Data.MySqlClient;
using denemelikimid.DataBase.Models;

namespace denemelikimid.DataBase.Repos
{
    internal class BordroRepo
    {
        private readonly DbRepository _db = new DbRepository();

        public DataTable Listele()
        {
            return _db.GetByQuery(
                @"SELECT * FROM bordro ");
        }

        public void Ekle(BordroModel model) {
            string query = @"
         INSERT INTO bordro (
        b_tc, b_ad_soyad, b_gorev_yeri, b_aylik_calisilan_gun, 
        b_cep_harcligi_tutari, b_sosyal_guvenlik_primi, b_tahakkuk_toplami, 
        b_gelir_vergisi_matrahi, b_hesaplanan_gelir_vergisi, b_gv_istisna_tutari, 
        b_gelir_vergisi_kesintisi, b_hesaplanan_damga_vergisi, b_dv_istisna_tutari, 
        b_damga_vergisi_kesintisi, b_gssp_kvsk, b_icra_kesintisi, 
        b_toplam_kesinti, b_odenmesi_gereken_net_tutar )
         VALUES (
        @tc, @adsoyad, @gorevyeri, @aylik_calisilan_gun, 
        @cep_harcligi, @sosyal_guvenlik_primi, @tahakkuk_tplm, 
        @gelir_vergisi_matrahi, @hesaplanan_gelir_vergisi, @gc_istisna_tutari, 
        @gelir_vergisi_kesintisi, @hesaplanan_damga_vergisi, @dv_istisna_tutari, 
        @damga_verigisi_kesintisi, @gssp_kvsk, @icra_kesintisi, 
        @toplam_kesinti, @odenmesi_gereken_net_tutar )";
            _db.Execute(query, 
            
                new MySqlParameter("@tc", model.Bdr_tc),
                new MySqlParameter("@adsoyad", model.Bdr_ad_soyad),
                new MySqlParameter("@gorevyeri", model.Bdr_gorev_yeri),
                new MySqlParameter("@aylik_calisilan_gun", model.Bdr_aylik_calisilan_gun),
                new MySqlParameter("@cep_harcligi", model.Bdr_cep_harcligi),
                new MySqlParameter("@sosyal_guvenlik_primi", model.Bdr_sosyal_guvenlik_primi),
                new MySqlParameter("@tahakkuk_tplm", model.Bdr_tahakkuk_toplami),
                new MySqlParameter("@gelir_vergisi_matrahi", model.Bdr_gelir_vergisi_matrahi),
                new MySqlParameter("@hesaplanan_gelir_vergisi", model.Bdr_hesaplanan_gelir_vergisi),
                new MySqlParameter("@gc_istisna_tutari", model.Bdr_gv_istisna_tutari),
                new MySqlParameter("@gelir_vergisi_kesintisi", model.Bdr_hesaplanan_damga_vergisi),
                new MySqlParameter("@hesaplanan_damga_vergisi", model.Bdr_hesaplanan_damga_vergisi),
                new MySqlParameter("@dv_istisna_tutari", model.Bdr_dv_istisna_tutari),
                new MySqlParameter("@damga_verigisi_kesintisi", model.Bdr_damga_vergisi_kesintisi),
                new MySqlParameter("@gssp_kvsk", model.Bdr_gssp_kvsk),
                new MySqlParameter("@icra_kesintisi", model.Bdr_icra_kesintisi),
                new MySqlParameter("@toplam_kesinti", model.Bdr_toplam_kesinti),
                new MySqlParameter("@odenmesi_gereken_net_tutar", model.Bdr_net_odenek)
            );
            
        }
        public void Sil(BordroModel model)
        {
            _db.Execute(
                @"DELETE FROM bordro WHERE b_id = @id",
                new MySqlParameter("@id", model.Bdr_id)
            );
        }
        public void Guncelle(BordroModel model)
        {
            string query = @"UPDATE bordro SET 
                b_tc = @tc,
                b_ad_soyad = @adsoyad,
                b_gorev_yeri = @gorevyeri,
                b_aylik_calisilan_gun = @aylik_calisilan_gun,
                b_cep_harcligi_tutari = @cep_harcligi,
                b_sosyal_guvenlik_primi = @sosyal_guvenlik_primi,
                b_tahakkuk_toplami = @tahakkuk_tplm,
                b_gelir_vergisi_matrahi = @gelir_vergisi_matrahi,
                b_hesaplanan_gelir_vergisi = @hesaplanan_gelir_vergisi,
                b_gv_istisna_tutari = @gc_istisna_tutari,
                b_gelir_vergisi_kesintisi = @gelir_vergisi_kesintisi,
                b_hesaplanan_damga_vergisi = @hesaplanan_damga_vergisi,
                b_dv_istisna_tutari = @dv_istisna_tutari,
                b_damga_vergisi_kesintisi = @damga_verigisi_kesintisi,
                b_gssp_kvsk = @gssp_kvsk,
                b_icra_kesintisi = @icra_kesintisi,
                b_toplam_kesinti = @toplam_kesinti,
                b_odenmesi_gereken_net_tutar = @odenmesi_gereken_net_tutar
                WHERE idbordo = @id";
            _db.Execute(query,
                new MySqlParameter("@tc", model.Bdr_tc),
                new MySqlParameter("@adsoyad", model.Bdr_ad_soyad),
                new MySqlParameter("@gorevyeri", model.Bdr_gorev_yeri),
                new MySqlParameter("@aylik_calisilan_gun", model.Bdr_aylik_calisilan_gun),
                new MySqlParameter("@cep_harcligi", model.Bdr_cep_harcligi),
                new MySqlParameter("@sosyal_guvenlik_primi", model.Bdr_sosyal_guvenlik_primi),
                new MySqlParameter("@tahakkuk_tplm", model.Bdr_tahakkuk_toplami),
                new MySqlParameter("@gelir_vergisi_matrahi", model.Bdr_gelir_vergisi_matrahi),
                new MySqlParameter("@hesaplanan_gelir_vergisi", model.Bdr_hesaplanan_gelir_vergisi),
                new MySqlParameter("@gc_istisna_tutari", model.Bdr_gv_istisna_tutari),
                new MySqlParameter("@gelir_vergisi_kesintisi", model.Bdr_hesaplanan_gelir_vergisi),
                new MySqlParameter("@hesaplanan_damga_vergisi", model.Bdr_hesaplanan_damga_vergisi),
                new MySqlParameter("@dv_istisna_tutari", model.Bdr_dv_istisna_tutari),
                new MySqlParameter("@damga_verigisi_kesintisi", model.Bdr_damga_vergisi_kesintisi),
                new MySqlParameter("@gssp_kvsk", model.Bdr_gssp_kvsk),
                new MySqlParameter("@icra_kesintisi", model.Bdr_icra_kesintisi),
                new MySqlParameter("@toplam_kesinti", model.Bdr_toplam_kesinti),
                new MySqlParameter("@odenmesi_gereken_net_tutar", model.Bdr_net_odenek)
            );
         }

                  
    }
}
