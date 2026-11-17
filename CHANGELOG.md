# Sürüm notları

## v0.9.1 Genel Beta — 18 Kasım 2026

- **İş Masam** için seçilen görevin e-posta metnini, konusunu, gönderenini ve alıcısını aynı pencerede gösteren VBA UserForm kodu eklendi.
- Outlook'un bulabildiği konuşma yanıtları tarih sırasıyla gösterilebilir (en çok 45 öğe incelenir, en son 16 mesaj gösterilir; her biri en fazla 1.700 karakter).
- Tek pencerede **Orijinal E-postayı Aç**, **Hatırlatma Taslağı**, **Tamamlandı** ve **Yenile** düğmeleri sağlandı.
- Bozuk `.frm` form içe aktarma yolu **kaldırıldı**. Artık bir defalık **boş UserForm + kod yapıştırma** kurulumu açıklanıyor.
- `Radar_11_IsMasam` form sınıfını çalışma anında yükleyerek form kurulmamışsa diğer makroların derlenmesini engellemeyecek şekilde düzenlendi.
- Mevcut `TakipKolay` görev kayıtları ve `ThisOutlookSession` değişmedi.
- **Beta uyarısı:** Kaynak ve paket kontrolleri yapıldı; güncel form gerçek Outlook'ta henüz uçtan uca doğrulanmadı.

## v0.8.3 Genel Beta — 16 Kasım 2026

- Hızlı tarih menüsü `0`–`5` ile eklendi. Kullanıcı seçmeden tarih atanmaz.
- Outlook'un gruplu görev görünümü ve yapacağım işler/yanıt bekleyenler ayrımı devam eder.
- Yeni takiplerde e-postanın metnini görev içine otomatik kopyalama kaldırıldı.

## v0.7.1

- Yapılacak işlerin raporda doğru kategoride görünmesi ve aynı yazışmada iki ayrı takip desteği için düzeltmeler.
