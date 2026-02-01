using System;
using System.Collections.Generic;
using System.Text;
using MySql.Data.MySqlClient;
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
                new MySqlParameter("@belgenin_mahiyeti", model.Mht_belgenin_mahiyeti),
                new MySqlParameter("@belge_turu", model.Mht_belge_turu),
                new MySqlParameter("@kanun_no", model.Mht_kanun_no),
                new MySqlParameter("@yeni_unite_kodu", model.Mht_yeni_unite_kodu),
                new MySqlParameter("@isyeri_sira_no", model.Mht_isyeri_sira_no),
                new MySqlParameter("@il_kodu", model.Mht_il_kodu),
                new MySqlParameter("@isveren_no", model.Mht_isveren_no),
                new MySqlParameter("@ssk_sicil", model.Mht_ssk_sicil),
                new MySqlParameter("@tc", model.Mht_tc),
                new MySqlParameter("@ad_soyad", model.Mht_ad_soyad),
                new MySqlParameter("@prim_odeme_gunu", model.Mht_prim_odeme_gunu),
                new MySqlParameter("@uzaktan_calisma_gunu", model.Mht_uzaktan_calisma_gunu),
                new MySqlParameter("@hak_edilen_ucret", model.Mht_hakedilen_ucret),
                new MySqlParameter("@prim_ikramiye_vb_istihkak", model.Mht_prim_ikramiye_vb_istihkak),
                new MySqlParameter("@ise_giris_tarihi", model.Mht_ise_giris_tarihi),
                new MySqlParameter("@isten_cikis_tarihi", model.Mht_isten_cikis_tarihi),
                new MySqlParameter("@isten_cikma_nedeni", model.Mht_isten_cikma_nedeni),
                new MySqlParameter("@eksik_gun_sayisi", model.Mht_eksik_gun_sayisi),
                new MySqlParameter("@eksik_gun_nedeni", model.Mht_eksik_gun_nedeni),
                new MySqlParameter("@meslek_kodu", model.Mht_meslek_kodu),
                new MySqlParameter("@istirahat_suresinde_calismamistir", model.Mht_istirahat_suresinde_calismamistir),
                new MySqlParameter("@tahakkuk_nedeni", model.Mht_tahhakkuk_nedeni),
                new MySqlParameter("@hizmet_donem_ay", model.Mht_hizmet_donem_ay),
                new MySqlParameter("@gelir_vergisinden_muaflik", model.Mht_gelir_vergisinden_muaflik),
                new MySqlParameter("@doneme_ait_gelir_vergisi_matrahi", model.Mht_doneme_ait_gelir_vergisi_matrahi),
                new MySqlParameter("@hizmet_donem_yil", model.Mht_hizmet_donem_yil),
                new MySqlParameter("@gv_engellilik_orani", model.Mht_gv_engellilik_orani),
                new MySqlParameter("@hesaplanan_gelir_vergisi", model.Mht_hesaplanan_gelir_vergisi),
                new MySqlParameter("@asgari_ucret_istisna_gelir_vergisi_tutari", model.Mht_asgari_ucret_istisna_gelir_vergisi_tutari),
                new MySqlParameter("@gelir_vergisi_kesintisi", model.Mht_gelir_vergisi_kesintisi),
                new MySqlParameter("@asgari_ucret_istisna_damga_vergisi_tutari", model.Mht_asgari_ucret_istisna_damga_vergisi_tutari),
                new MySqlParameter("@damga_vergisi_kesintisi", model.Mht_damga_vergisi_kesintisi)
            );
        }
        public void Sil(MuhtasarModel model)
        {
            _db.Execute(
                @"DELETE FROM muhtasar_raporu WHERE mht_id= @id",
                new MySqlParameter("@id", model.Mht_id));
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
                new MySqlParameter("@belgenin_mahiyeti", model.Mht_belgenin_mahiyeti),
                new MySqlParameter("@belge_turu", model.Mht_belge_turu),
                new MySqlParameter("@kanun_no", model.Mht_kanun_no),
                new MySqlParameter("@yeni_unite_kodu", model.Mht_yeni_unite_kodu),
                new MySqlParameter("@isyeri_sira_no", model.Mht_isyeri_sira_no),
                new MySqlParameter("@il_kodu", model.Mht_il_kodu),
                new MySqlParameter("@isveren_no", model.Mht_isveren_no),
                new MySqlParameter("@ssk_sicil", model.Mht_ssk_sicil),
                new MySqlParameter("@tc", model.Mht_tc), // Bu WHERE koşulu için lazım
                new MySqlParameter("@ad_soyad", model.Mht_ad_soyad),
                new MySqlParameter("@prim_odeme_gunu", model.Mht_prim_odeme_gunu),
                new MySqlParameter("@uzaktan_calisma_gunu", model.Mht_uzaktan_calisma_gunu),
                new MySqlParameter("@hak_edilen_ucret", model.Mht_hakedilen_ucret),
                new MySqlParameter("@prim_ikramiye_vb_istihkak", model.Mht_prim_ikramiye_vb_istihkak),
                new MySqlParameter("@ise_giris_tarihi", model.Mht_ise_giris_tarihi),
                new MySqlParameter("@isten_cikis_tarihi", model.Mht_isten_cikis_tarihi),
                new MySqlParameter("@isten_cikma_nedeni", model.Mht_isten_cikma_nedeni),
                new MySqlParameter("@eksik_gun_sayisi", model.Mht_eksik_gun_sayisi),
                new MySqlParameter("@eksik_gun_nedeni", model.Mht_eksik_gun_nedeni),
                new MySqlParameter("@meslek_kodu", model.Mht_meslek_kodu),
                new MySqlParameter("@istirahat_suresinde_calismamistir", model.Mht_istirahat_suresinde_calismamistir),
                new MySqlParameter("@tahakkuk_nedeni", model.Mht_tahhakkuk_nedeni),
                new MySqlParameter("@hizmet_donem_ay", model.Mht_hizmet_donem_ay),
                new MySqlParameter("@gelir_vergisinden_muaflik", model.Mht_gelir_vergisinden_muaflik),
                new MySqlParameter("@doneme_ait_gelir_vergisi_matrahi", model.Mht_doneme_ait_gelir_vergisi_matrahi),
                new MySqlParameter("@hizmet_donem_yil", model.Mht_hizmet_donem_yil),
                new MySqlParameter("@gv_engellilik_orani", model.Mht_gv_engellilik_orani),
                new MySqlParameter("@hesaplanan_gelir_vergisi", model.Mht_hesaplanan_gelir_vergisi),
                new MySqlParameter("@asgari_ucret_istisna_gelir_vergisi_tutari", model.Mht_asgari_ucret_istisna_gelir_vergisi_tutari),
                new MySqlParameter("@gelir_vergisi_kesintisi", model.Mht_gelir_vergisi_kesintisi),
                new MySqlParameter("@asgari_ucret_istisna_damga_vergisi_tutari", model.Mht_asgari_ucret_istisna_damga_vergisi_tutari),
                new MySqlParameter("@damga_vergisi_kesintisi", model.Mht_damga_vergisi_kesintisi),
                new MySqlParameter("@mht_id", model.Mht_id)
            );
        }
    }
}
