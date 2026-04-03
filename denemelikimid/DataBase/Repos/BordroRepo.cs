using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using Microsoft.Data.Sqlite;
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
            
                new SqliteParameter("@tc", model.Bdr_tc),
                new SqliteParameter("@adsoyad", model.Bdr_ad_soyad),
                new SqliteParameter("@gorevyeri", model.Bdr_gorev_yeri),
                new SqliteParameter("@aylik_calisilan_gun", model.Bdr_aylik_calisilan_gun),
                new SqliteParameter("@cep_harcligi", model.Bdr_cep_harcligi),
                new SqliteParameter("@sosyal_guvenlik_primi", model.Bdr_sosyal_guvenlik_primi),
                new SqliteParameter("@tahakkuk_tplm", model.Bdr_tahakkuk_toplami),
                new SqliteParameter("@gelir_vergisi_matrahi", model.Bdr_gelir_vergisi_matrahi),
                new SqliteParameter("@hesaplanan_gelir_vergisi", model.Bdr_hesaplanan_gelir_vergisi),
                new SqliteParameter("@gc_istisna_tutari", model.Bdr_gv_istisna_tutari),
                new SqliteParameter("@gelir_vergisi_kesintisi", model.Bdr_hesaplanan_damga_vergisi),
                new SqliteParameter("@hesaplanan_damga_vergisi", model.Bdr_hesaplanan_damga_vergisi),
                new SqliteParameter("@dv_istisna_tutari", model.Bdr_dv_istisna_tutari),
                new SqliteParameter("@damga_verigisi_kesintisi", model.Bdr_damga_vergisi_kesintisi),
                new SqliteParameter("@gssp_kvsk", model.Bdr_gssp_kvsk),
                new SqliteParameter("@icra_kesintisi", model.Bdr_icra_kesintisi),
                new SqliteParameter("@toplam_kesinti", model.Bdr_toplam_kesinti),
                new SqliteParameter("@odenmesi_gereken_net_tutar", model.Bdr_net_odenek)
            );
            
        }
        public void Sil(BordroModel model)
        {
            _db.Execute(
                @"DELETE FROM bordro WHERE b_id = @id",
                new SqliteParameter("@id", model.Bdr_id)
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
                new SqliteParameter("@tc", model.Bdr_tc),
                new SqliteParameter("@adsoyad", model.Bdr_ad_soyad),
                new SqliteParameter("@gorevyeri", model.Bdr_gorev_yeri),
                new SqliteParameter("@aylik_calisilan_gun", model.Bdr_aylik_calisilan_gun),
                new SqliteParameter("@cep_harcligi", model.Bdr_cep_harcligi),
                new SqliteParameter("@sosyal_guvenlik_primi", model.Bdr_sosyal_guvenlik_primi),
                new SqliteParameter("@tahakkuk_tplm", model.Bdr_tahakkuk_toplami),
                new SqliteParameter("@gelir_vergisi_matrahi", model.Bdr_gelir_vergisi_matrahi),
                new SqliteParameter("@hesaplanan_gelir_vergisi", model.Bdr_hesaplanan_gelir_vergisi),
                new SqliteParameter("@gc_istisna_tutari", model.Bdr_gv_istisna_tutari),
                new SqliteParameter("@gelir_vergisi_kesintisi", model.Bdr_hesaplanan_gelir_vergisi),
                new SqliteParameter("@hesaplanan_damga_vergisi", model.Bdr_hesaplanan_damga_vergisi),
                new SqliteParameter("@dv_istisna_tutari", model.Bdr_dv_istisna_tutari),
                new SqliteParameter("@damga_verigisi_kesintisi", model.Bdr_damga_vergisi_kesintisi),
                new SqliteParameter("@gssp_kvsk", model.Bdr_gssp_kvsk),
                new SqliteParameter("@icra_kesintisi", model.Bdr_icra_kesintisi),
                new SqliteParameter("@toplam_kesinti", model.Bdr_toplam_kesinti),
                new SqliteParameter("@odenmesi_gereken_net_tutar", model.Bdr_net_odenek)
            );
         }

                  
    }
}
