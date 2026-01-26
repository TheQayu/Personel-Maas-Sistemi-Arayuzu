using System;
using System.Collections.Generic;
using System.Text;

namespace IskurTakipSistemi.DataBase.Models
{
    internal class BankaModel
    {
        public int Bnk_id { get; set; }
        public string Bnk_iban { get; set; }
        public string Bnk_ad_soyad { get; set; }
        public string Bnk_tc { get; set; }
        public string Bnk_aciklama { get; set; }
        public float Bnk_tutar { get; set; }
    }
}
