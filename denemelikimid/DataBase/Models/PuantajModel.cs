using System;
using System.Collections.Generic;
using System.Text;

namespace IskurTakipSistemi.DataBase.Models
{
    internal class PuantajModel
    {
        public int Pnt_id { get; set; }
        public int Pnt_calisilan_gun_sayisi { get; set; }
        public int Pnt_yillik_izin { get; set; }
        public int Pnt_devamsizlik { get; set; }
        public string Pnt_iban { get; set; }
        public string Pnt_tc { get; set; }
        public string Pnt_ad_soyad { get; set; }
        public DateTime Pnt_ise_baslama_tarihi { get; set; }
        public DateTime Pnt_isten_ayrilma_tarihi { get; set; }
    }
}
