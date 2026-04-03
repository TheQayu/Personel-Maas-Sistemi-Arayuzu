using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Data.Sqlite;
using System.Data;
using denemelikimid.DataBase;
using denemelikimid.DataBase.Models;


namespace denemelikimid.DataBase.Repos
{
    internal class MuhtasarRepo
    {
        private readonly DbRepository _db = new DbRepository();

        public DataTable Listele()
        {
            return _db.GetByQuery(
                @"SELECT * FROM muhtasar_raporu ");
        }
        public void Ekle(MuhtasarModel model)
        {
            // 1. Düzeltme: Kolon isimlerinin başında @ olmaz, sadece VALUES tarafında olur.
            string query = @"INSERT INTO muhtasar_raporu (
                        mh_belgenin_mahiyeti, mh_belge_turu, mh_kanun_no, mh_yeni_unite_kodu, 
                        mh_isyeri_sira_no, mh_il_kodu, mh_isveren_no, mh_ssk_sicil, 
                        mh_tc, mh_ad_soyad, mh_prim_odeme_gunu, mh_uzaktan_calisma_gunu, 
                        mh_hak_edilen_ucret, mh_prim_ikramiye_vb_istihkak, mh_ise_giris_tarihi, 
                        mh_isten_cikis_tarihi, mh_isten_cikma_nedeni, mh_eksik_gun_sayisi, 
                        mh_eksik_gun_nedeni, mh_meslek_kodu, mh_istirahat_suresinde_calismamistir, 
                        mh_tahakkuk_nedeni, mh_hizmet_donem_ay, mh_gelir_vergisinden_muaflik, 
                        mh_doneme_ait_gelir_vergisi_matrahi, mh_hizmet_donem_yil, 
                        mh_gv_engellilik_orani, mh_hesaplanan_gelir_vergisi, 
                        mh_asgari_ucret_istisna_gelir_vergisi_tutari, mh_gelir_vergisi_kesintisi, 
                        mh_asgari_ucret_istisna_damga_vergisi_tutari, mh_damga_vergisi_kesintisi)
                    VALUES (
                        @belgenin_mahiyeti, @belge_turu, @kanun_no, @yeni_unite_kodu, 
                        @isyeri_sira_no, @il_kodu, @isveren_no, @ssk_sicil, 
                        @tc, @ad_soyad, @prim_odeme_gunu, @uzaktan_calisma_gunu, 
                        @hak_edilen_ucret, @prim_ikramiye_vb_istihkak, @ise_giris_tarihi, 
                        @isten_cikis_tarihi, @isten_cikma_nedeni, @eksik_gun_sayisi, 
                        @eksik_gun_nedeni, @meslek_kodu, @istirahat_suresinde_calismamistir, 
                        @tahakkuk_nedeni, @hizmet_donem_ay, @gelir_vergisinden_muaflik, 
                        @doneme_ait_gelir_vergisi_matrahi, @hizmet_donem_yil, 
                        @gv_engellilik_orani, @hesaplanan_gelir_vergisi, 
                        @asgari_ucret_istisna_gelir_vergisi_tutari, @gelir_vergisi_kesintisi, 
                        @asgari_ucret_istisna_damga_vergisi_tutari, @damga_vergisi_kesintisi)";

            _db.Execute(
                query,
                new SqliteParameter("@belgenin_mahiyeti", model.Mht_belgenin_mahiyeti),
                new SqliteParameter("@belge_turu", model.Mht_belge_turu),
                new SqliteParameter("@kanun_no", model.Mht_kanun_no),
                new SqliteParameter("@yeni_unite_kodu", model.Mht_yeni_unite_kodu),
                new SqliteParameter("@isyeri_sira_no", model.Mht_isyeri_sira_no),
                new SqliteParameter("@il_kodu", model.Mht_il_kodu),
                new SqliteParameter("@isveren_no", model.Mht_isveren_no),
                new SqliteParameter("@ssk_sicil", model.Mht_ssk_sicil),
                new SqliteParameter("@tc", model.Mht_tc),
                new SqliteParameter("@ad_soyad", model.Mht_ad_soyad),
                new SqliteParameter("@prim_odeme_gunu", model.Mht_prim_odeme_gunu),
                new SqliteParameter("@uzaktan_calisma_gunu", model.Mht_uzaktan_calisma_gunu),
                new SqliteParameter("@hak_edilen_ucret", model.Mht_hakedilen_ucret),
                new SqliteParameter("@prim_ikramiye_vb_istihkak", model.Mht_prim_ikramiye_vb_istihkak),
                new SqliteParameter("@ise_giris_tarihi", model.Mht_ise_giris_tarihi),
                new SqliteParameter("@isten_cikis_tarihi", model.Mht_isten_cikis_tarihi),
                new SqliteParameter("@isten_cikma_nedeni", model.Mht_isten_cikma_nedeni),
                new SqliteParameter("@eksik_gun_sayisi", model.Mht_eksik_gun_sayisi),
                new SqliteParameter("@eksik_gun_nedeni", model.Mht_eksik_gun_nedeni),
                new SqliteParameter("@meslek_kodu", model.Mht_meslek_kodu),
                new SqliteParameter("@istirahat_suresinde_calismamistir", model.Mht_istirahat_suresinde_calismamistir),
                new SqliteParameter("@tahakkuk_nedeni", model.Mht_tahhakkuk_nedeni),
                new SqliteParameter("@hizmet_donem_ay", model.Mht_hizmet_donem_ay),
                new SqliteParameter("@gelir_vergisinden_muaflik", model.Mht_gelir_vergisinden_muaflik),
                new SqliteParameter("@doneme_ait_gelir_vergisi_matrahi", model.Mht_doneme_ait_gelir_vergisi_matrahi),
                new SqliteParameter("@hizmet_donem_yil", model.Mht_hizmet_donem_yil),
                new SqliteParameter("@gv_engellilik_orani", model.Mht_gv_engellilik_orani),
                new SqliteParameter("@hesaplanan_gelir_vergisi", model.Mht_hesaplanan_gelir_vergisi),
                new SqliteParameter("@asgari_ucret_istisna_gelir_vergisi_tutari", model.Mht_asgari_ucret_istisna_gelir_vergisi_tutari),
                new SqliteParameter("@gelir_vergisi_kesintisi", model.Mht_gelir_vergisi_kesintisi),
                new SqliteParameter("@asgari_ucret_istisna_damga_vergisi_tutari", model.Mht_asgari_ucret_istisna_damga_vergisi_tutari),
                new SqliteParameter("@damga_vergisi_kesintisi", model.Mht_damga_vergisi_kesintisi)
            );
        }
        public void Sil(MuhtasarModel model)
        {
            _db.Execute(
                @"DELETE FROM muhtasar_raporu WHERE mht_id= @id",
                new SqliteParameter("@id", model.Mht_id));
        }

        public void Guncelle(MuhtasarModel model)
        {
            string query = @"
        UPDATE muhtasar_raporu 
        SET 
            mh_belgenin_mahiyeti = @belgenin_mahiyeti, 
            mh_belge_turu = @belge_turu, 
            mh_kanun_no = @kanun_no, 
            mh_yeni_unite_kodu = @yeni_unite_kodu, 
            mh_isyeri_sira_no = @isyeri_sira_no, 
            mh_il_kodu = @il_kodu, 
            mh_isveren_no = @isveren_no, 
            mh_ssk_sicil = @ssk_sicil, 
            mh_ad_soyad = @ad_soyad, 
            mh_prim_odeme_gunu = @prim_odeme_gunu, 
            mh_uzaktan_calisma_gunu = @uzaktan_calisma_gunu, 
            mh_hak_edilen_ucret = @hak_edilen_ucret, 
            mh_prim_ikramiye_vb_istihkak = @prim_ikramiye_vb_istihkak, 
            mh_ise_giris_tarihi = @ise_giris_tarihi, 
            mh_isten_cikis_tarihi = @isten_cikis_tarihi, 
            mh_isten_cikma_nedeni = @isten_cikma_nedeni, 
            mh_eksik_gun_sayisi = @eksik_gun_sayisi, 
            mh_eksik_gun_nedeni = @eksik_gun_nedeni, 
            mh_meslek_kodu = @meslek_kodu, 
            mh_istirahat_suresinde_calismamistir = @istirahat_suresinde_calismamistir, 
            mh_tahakkuk_nedeni = @tahakkuk_nedeni, 
            mh_hizmet_donem_ay = @hizmet_donem_ay, 
            mh_gelir_vergisinden_muaflik = @gelir_vergisinden_muaflik, 
            mh_doneme_ait_gelir_vergisi_matrahi = @doneme_ait_gelir_vergisi_matrahi, 
            mh_hizmet_donem_yil = @hizmet_donem_yil, 
            mh_gv_engellilik_orani = @gv_engellilik_orani, 
            mh_hesaplanan_gelir_vergisi = @hesaplanan_gelir_vergisi, 
            mh_asgari_ucret_istisna_gelir_vergisi_tutari = @asgari_ucret_istisna_gelir_vergisi_tutari, 
            mh_gelir_vergisi_kesintisi = @gelir_vergisi_kesintisi, 
            mh_asgari_ucret_istisna_damga_vergisi_tutari = @asgari_ucret_istisna_damga_vergisi_tutari, 
            mh_damga_vergisi_kesintisi = @damga_vergisi_kesintisi
        WHERE idmuhtasar_raporu = @mht_id";

            _db.Execute(
                query,
                new SqliteParameter("@belgenin_mahiyeti", model.Mht_belgenin_mahiyeti),
                new SqliteParameter("@belge_turu", model.Mht_belge_turu),
                new SqliteParameter("@kanun_no", model.Mht_kanun_no),
                new SqliteParameter("@yeni_unite_kodu", model.Mht_yeni_unite_kodu),
                new SqliteParameter("@isyeri_sira_no", model.Mht_isyeri_sira_no),
                new SqliteParameter("@il_kodu", model.Mht_il_kodu),
                new SqliteParameter("@isveren_no", model.Mht_isveren_no),
                new SqliteParameter("@ssk_sicil", model.Mht_ssk_sicil),
                new SqliteParameter("@tc", model.Mht_tc), // Bu WHERE koşulu için lazım
                new SqliteParameter("@ad_soyad", model.Mht_ad_soyad),
                new SqliteParameter("@prim_odeme_gunu", model.Mht_prim_odeme_gunu),
                new SqliteParameter("@uzaktan_calisma_gunu", model.Mht_uzaktan_calisma_gunu),
                new SqliteParameter("@hak_edilen_ucret", model.Mht_hakedilen_ucret),
                new SqliteParameter("@prim_ikramiye_vb_istihkak", model.Mht_prim_ikramiye_vb_istihkak),
                new SqliteParameter("@ise_giris_tarihi", model.Mht_ise_giris_tarihi),
                new SqliteParameter("@isten_cikis_tarihi", model.Mht_isten_cikis_tarihi),
                new SqliteParameter("@isten_cikma_nedeni", model.Mht_isten_cikma_nedeni),
                new SqliteParameter("@eksik_gun_sayisi", model.Mht_eksik_gun_sayisi),
                new SqliteParameter("@eksik_gun_nedeni", model.Mht_eksik_gun_nedeni),
                new SqliteParameter("@meslek_kodu", model.Mht_meslek_kodu),
                new SqliteParameter("@istirahat_suresinde_calismamistir", model.Mht_istirahat_suresinde_calismamistir),
                new SqliteParameter("@tahakkuk_nedeni", model.Mht_tahhakkuk_nedeni),
                new SqliteParameter("@hizmet_donem_ay", model.Mht_hizmet_donem_ay),
                new SqliteParameter("@gelir_vergisinden_muaflik", model.Mht_gelir_vergisinden_muaflik),
                new SqliteParameter("@doneme_ait_gelir_vergisi_matrahi", model.Mht_doneme_ait_gelir_vergisi_matrahi),
                new SqliteParameter("@hizmet_donem_yil", model.Mht_hizmet_donem_yil),
                new SqliteParameter("@gv_engellilik_orani", model.Mht_gv_engellilik_orani),
                new SqliteParameter("@hesaplanan_gelir_vergisi", model.Mht_hesaplanan_gelir_vergisi),
                new SqliteParameter("@asgari_ucret_istisna_gelir_vergisi_tutari", model.Mht_asgari_ucret_istisna_gelir_vergisi_tutari),
                new SqliteParameter("@gelir_vergisi_kesintisi", model.Mht_gelir_vergisi_kesintisi),
                new SqliteParameter("@asgari_ucret_istisna_damga_vergisi_tutari", model.Mht_asgari_ucret_istisna_damga_vergisi_tutari),
                new SqliteParameter("@damga_vergisi_kesintisi", model.Mht_damga_vergisi_kesintisi),
                new SqliteParameter("@mht_id", model.Mht_id)
            );
        }
    }
}
