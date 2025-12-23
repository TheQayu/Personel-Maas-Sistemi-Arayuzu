using System;
using System.Collections.Generic; // Dictionary için gerekli
using System.Linq;              // Count fonksiyonu için gerekli

namespace denemelikimid  // <-- Buranın proje adınızla aynı olduğundan emin olun
{
    public class PersonelVerisi
    {
        // Temel Bilgiler (Puantaj.csv'den)
        public string TCNo { get; set; }
        public string AdSoyad { get; set; }
        public string IBAN { get; set; } // Banka Listesi için

        // Puantaj Bilgisi (1-31 Gün)
        // Her gün için durumu tutar: "X", "R" (Rapor), "İ" (İzin) veya boş
        public Dictionary<int, string> GunlukPuantaj { get; set; } = new Dictionary<int, string>();

        // Hesaplama Parametreleri (BORDRO.csv için)
        public decimal GunlukBrutUcret { get; set; } // Örn: 666.75 TL (2024 Asgari)

        // Hesaplanan Veriler (Otomatik dolacak)
        // "X" veya "Tam" olan günleri sayar
        public int ToplamCalismaGunu => GunlukPuantaj.Count(x => x.Value == "X" || x.Value == "Tam");

        public decimal ToplamBrut => ToplamCalismaGunu * GunlukBrutUcret;

        public decimal NetEleGecen => MaasHesapla(ToplamBrut);

        // Basit bir Net Maaş Hesaplama Fonksiyonu
        private decimal MaasHesapla(decimal brut)
        {
            // ÖRNEK: Sadece Damga Vergisi kesintisi varsa (0.00759)
            // İŞKUR kurallarına göre burayı sonra güncelleriz
            decimal damgaVergisi = brut * 0.00759m;
            return brut - damgaVergisi;
        }
    }
}