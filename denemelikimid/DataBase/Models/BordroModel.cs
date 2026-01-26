using System;
using System.Collections.Generic;
using System.Text;

namespace IskurTakipSistemi.DataBase.Models
{
    internal class BordroModel
    {
        public float Bdr_cep_harcligi { get; set; }
        public float Bdr_sosyal_guvenlik_primi { get; set; }
        public float Bdr_tahakkuk_toplami { get; set; }
        public float Bdr_gelir_vergisi_matrahi { get; set; }
        public float Bdr_hesaplanan_gelir_vergisi { get; set; }
        public float Bdr_gv_istisna_tutari { get; set; }
        public float Bdr_hesaplanan_damga_vergisi { get; set; }
        public float Bdr_dv_istisna_tutari { get; set; }
        public float Bdr_damga_vergisi_kesintisi { get; set; }
        public float Bdr_gssp_kvsk { get; set; }
        public float Bdr_icra_kesintisi { get; set; }
        public float Bdr_toplam_kesinti { get; set; }
        public float Bdr_net_odenek { get; set; }
        public int Bdr_id { get; set; }
        public int Bdr_aylik_calisilan_gun { get; set; }
        public string Bdr_tc { get; set; }
        public string Bdr_ad_soyad { get; set; }
        public string Bdr_gorev_yeri { get; set; }
    }
}
