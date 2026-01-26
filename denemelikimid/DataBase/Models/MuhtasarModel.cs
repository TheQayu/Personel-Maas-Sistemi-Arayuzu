using System;

namespace IskurTakipSistemi.DataBase.Models
{
    internal class MuhtasarModel
    {
        public int Mht_id { get; set; }
        public int Mht_uzaktan_calisma_gunu { get; set; }
        public int Mht_eksik_gun_sayisi { get; set; }
        public string Mht_belgenin_mahiyeti { get; set; }
        public string Mht_belge_turu { get; set; }
        public string Mht_kanun_no { get; set; }
        public string Mht_yeni_unite_kodu { get; set; }
        public string Mht_isyeri_sira_no { get; set; }
        public string Mht_ssk_sicil { get; set; }
        public string Mht_tc { get; set; }
        public string Mht_ad_soyad { get; set; }
        public string Mht_prim_odeme_gunu { get; set; }
        public string Mht_isten_cikma_nedeni { get; set; }
        public string Mht_meslek_kodu { get; set; }
        public string Mht_tahhakkuk_nedeni { get; set; }
        public string Mht_hizmet_donem_ay { get; set; }
        public string Mht_hizmet_donem_yil { get; set; }
        public string Mht_gv_engellilik_orani { get; set; }
        public string Mht_il_kodu { get; set; }
        public string Mht_isveren_no { get; set; }
        public string Mht_eksik_gun_nedeni { get; set; }
        public float Mht_hakedilen_ucret { get; set; }
        public float Mht_prim_ikramiye_vb_istihkak { get; set; }
        public float Mht_doneme_ait_gelir_vergisi_matrahi { get; set; }
        public float Mht_hesaplanan_gelir_vergisi { get; set; }
        public float Mht_asgari_ucret_gelir_vergisi_tutari { get; set; }
        public float Mht_gelir_vergisi_kesintisi { get; set; }
        public float Mht_asgari_ucret_istisna_damga_vergisi_tutari { get; set; }
        public float Mh_damga_vergisi_kesintisi { get; set; }
        public bool Mht_istirahat_suresinde_calismamistir { get; set; }
        public bool Mht_gelir_vergisinden_muaflik { get; set; }
        public DateTime Mht_ise_giris_tarihi { get; set; }
        public DateTime Mht_isten_cikis_tarihi { get; set; }
    }
}
