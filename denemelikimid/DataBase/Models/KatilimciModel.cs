using System;
using System.Collections.Generic;
using System.Text;

namespace IskurTakipSistemi.DataBase.Models
{
    internal class KatilimciModel
    {
        public int Ktc_id { get; set; }
        public string Ktc_tc { get; set; }
        public string Ktc_ad_soyad { get; set; }
        public string Ktc_gorev_yeri { get; set; }
        public string Ktc_iban { get; set; }
        public DateTime Ktc_ise_baslama_tarihi { get; set; }
        public DateTime Ktc_isten_cikma_tarihi { get; set; }
      
    }
}
