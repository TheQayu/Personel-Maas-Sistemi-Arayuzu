# 🏛️ Üniversite Personel Takip Sistemi (Demo / Taslak)

Bu proje, C# ve Windows Forms (WinForms) kullanılarak geliştirilmiş, üniversite personel yönetimi, puantaj takibi ve maaş bordrosu işlemleri için tasarlanmış bir **arayüz (UI) ve mantık iskeletidir.**

⚠️ **Dikkat:** Bu proje şu an için bir **Demo/Taslak** niteliğindedir. Arayüz ve temel fonksiyonlar hazırlanmış olup, veritabanı bağlantıları ve karmaşık hesaplama mantıkları henüz entegre edilmemiştir.

##  Namespace Neden `denemelikimid`?

Kodları incelediğinizde namespace (isim uzayı) olarak **`denemelikimid`** ismini göreceksiniz.
* Bu proje, C# programlama dilini ve Windows Forms yapısını öğrenme sürecinde, **"Deneme amaçlı bir proje"** olarak başlatılmıştır.
* Projenin temel yapısı bu isim üzerine kurulduğu için, geliştirme sürecinde orijinalliği bozulmadan bu şekilde bırakılmıştır.
* Özetle: Evet, bu bir denemedir! :)

##  Projenin Amacı

Üniversite idari süreçlerinde kullanılan;
* **Puantaj Cetvelleri** (Günlük katılım durumu)
* **Banka Listeleri**
* **SGK Raporları**

gibi belgelerin Excel formatlarına uygun olarak, masaüstü uygulamasından otomatik üretilmesini simüle etmek ve bu süreçleri dijitalleştirmektir.

##  Kullanılan Teknolojiler ve Kütüphaneler

* **Dil:** C# (.NET Framework / .NET Core)
* **Arayüz:** Windows Forms (Code-First yaklaşımı ile tasarlanmıştır, Designer kullanılmamıştır).
* **Excel İşlemleri:** `ClosedXML` kütüphanesi kullanılmıştır.

##  Özellikler (Mevcut Durum)

* [x] **Modern Arayüz:** Sol menü (Sidebar), Üst Başlık (Header) ve İçerik Alanı (Content) ile bölünmüş responsive yapı.
* [x] **Dinamik Tablo:** 1'den 31'e kadar günleri otomatik oluşturan DataGridView yapısı.
* [x] **Excel Motoru:** Verilen listeyi `ClosedXML` kullanarak, formüllü ve biçimlendirilmiş gerçek bir Excel dosyasına dönüştürme yeteneği.
* [x] **Örnek Veri:** Test amaçlı "Ahmet", "Ayşe" gibi dummy (sahte) verilerle çalışır.

##  Yapılacaklar (To-Do)

Proje geliştirilmeye açıktır ve şu adımların tamamlanması hedeflenmektedir:
* [ ] SQL Veritabanı bağlantısının yapılması.
* [ ] Kullanıcıların (Personel) veritabanından çekilmesi.
* [ ] "X", "R", "İ" gibi puantaj kodlarının arayüzden girilebilir hale gelmesi.
* [ ] Girilen verilere göre net maaş hesaplama modülünün (Vergi dilimleri vb.) yazılması.

##  Kurulum ve Çalıştırma

1.  Projeyi bilgisayarınıza indirin (Clone veya Download ZIP).
2.  Visual Studio ile `.sln` dosyasını açın.
3.  **NuGet Paketlerini Yükleyin:**
    * Solution Explorer'da projeye sağ tıklayın -> `Manage NuGet Packages`.
    * **ClosedXML** paketinin yüklü olduğundan emin olun (Yoksa "Restore" yapın).
4.  Projeyi Derleyin (Build) ve Çalıştırın (Run).

---
*Geliştirici Notu: Kodlar öğrenme amaçlı yazıldığı için profesyonel mimari standartlarından (SOLID vb.) ziyade, çalışır bir prototip üretmeye odaklanılmıştır.*
