# Takip Kolay Android - kabul ve gizlilik testleri

Cihazda gerçek kurumsal veri kullanmadan, sentetik içerikle test edin.

- [ ] İlk kurulumda açık/kayıtlı görev yok.
- [ ] Yeni görevde tarih ve saat **boş**; hiç otomatik tarih verilmez.
- [ ] 20.11.2026 gibi gelecek bir tarih ve saat elle seçilip kaydedilir.
- [ ] "Tarihsiz" ile tarih kaldırılabilir ve tarihsiz görev kaydedilebilir.
- [ ] ACTION_SEND ile gönderilen başlık ve gövde alanlara aktarılır; yalnızca link gelirse tüm e-posta okunduğu iddia edilmez.
- [ ] Android PROCESS_TEXT destekleyen uygulamada seçim ile açılabilir.
- [ ] Metin 200.000 karakteri aşıyorsa sessiz kırpma veya kayıt yok.
- [ ] Görev kaydı uygulama kapatılıp açıldıktan sonra durur.
- [ ] Görev aç/kaydet/tamamlandı/yeniden aç/sil işlemleri çalışır.
- [ ] Kullanıcıdan izin alındığında seçilen gelecek zamana yakın yerel bildirim gelir.
- [ ] Android 13+ bildirim izni reddedilince görev yine kaydedilir; bildirim görünmeyeceği açıkça belirtilir.
- [ ] Tamamlanan veya silinen göreve daha sonra bildirim gelmez.
- [ ] Telefon kapatılıp açılınca gelecekteki hatırlatmalar kurulur.
- [ ] Uçak modunda uygulama, görevler ve mevcut yerel bildirimler çalışır.
- [ ] Uygulama manifestinde INTERNET izni yok; ağ trafiği veya telemetri yok.
- [ ] Yerel tasks.enc düz metin e-posta içeriği barındırmaz (AES-GCM ile şifreli).
- [ ] Ekran görüntüsü / son uygulamalar önizlemesi hassas metni göstermez.
- [ ] Şifreli dosya bozuksa kaydı sıfırlamadan hata gösterir.
- [ ] İş profili tarafından paylaşım yasaksa kurallar aşılmaz.
- [ ] Klasik Outlook masaüstü VBA sürümü değişmeden kalır.
