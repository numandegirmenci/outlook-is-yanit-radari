# Takip Kolay Android (0.1.0-alpha)

Bu klasör, Outlook İş ve Yanıt Radarı'nın klasik Outlook VBA sürümüne dokunmadan geliştirilmiş, **bağımsız Android yardımcısıdır**.

## Özellikler

- Android paylaş menüsünden **text/plain** metni alır (ACTION_SEND / EXTRA_TEXT, varsa EXTRA_SUBJECT).
- Destekleyen uygulamalarda metin seçimi üzerinden PROCESS_TEXT ile açılabilir.
- Outlook metni paylaştıramıyorsa, kullanıcı e-posta içeriğini kopyalayıp giriş alanına yapıştırabilir.
- Başlık, işlem notu ve e-posta konuşma metni görevde saklanır.
- Hatırlatma tarihi **kendiliğinden atanmaz**. Kullanıcı tarih ve saat seçerse tek seferlik yerel bildirim kurulur. Tarihsiz kayıt da mümkündür.
- Görev düzenleme, tamamlama/yeniden açma ve onaylı silme.
- Veriler Android Keystore ile AES-256-GCM şifrelenerek uygulamanın özel dosya alanında saklanır.
- Uygulama INTERNET izni istemez; analitik, reklam, Microsoft Graph, API, hesap girişi veya dış sunucu yoktur.
- Bildirim metni kilit ekranında şirket e-postası veya not içermez; ekran görüntüsü engellenmiştir.
- Telefon yeniden başladıktan sonra ilerideki uyarılar yeniden planlanır.

## Gerçek sınırlar

1. **Outlook'un e-posta metnini uygulamaya göndermesi garanti değildir.** Android paylaş menüsü yalnızca kaynak uygulamanın verdiği metni iletir. Yalnızca link/konu gelirse konuşmanın tamamı okunmaz. Gerekirse metni elle yapıştırın.
2. Bu sürüm **Outlook eklentisi değildir**; klasik masaüstü VBA sürümü ile görev senkronizasyonu yoktur.
3. Android'in pil ve bildirim kısıtları nedeniyle hatırlatmalar **yaklaşık saatte** gösterilir; kesin dakikası garanti edilmez. Android 13+ bildirim izni gerektirir.
4. Şirketin MDM/Intune veri paylaşım kuralları izin vermiyorsa uygulama iş profilinden metin alamaz. Kurumsal korumayı aşmaya çalışmaz.
5. Veriler telefona bağlıdır; başka telefona taşınmaz, yeniden kurulumda kaybolabilir. Android yedekleme kapalıdır. Ekran kilidi ve cihaz güvenliği yine önemlidir.

## Derleme

- Android Studio, JDK 17, Android SDK platform 35.
- Android Gradle Plugin 8.7.3 ve Gradle 8.9.
- Projeyi Android Studio ile mobile/android klasöründen açın; yerel Gradle 8.9 seçin veya kurulu Gradle ile komut satırında aşağıdakini çalıştırın:

    gradle :app:assembleDebug

- Debug APK: app/build/outputs/apk/debug/app-debug.apk
- Ayrıca GitHub Actions > Android debug APK iş akışında başarılı bir derlemenin artifacts bölümünden APK alınabilir. CI henüz geçmeden APK'nın çalıştığı iddia edilmez.

## Telefonda kullanım

1. Uygulamayı yükleyin, Görevi kaydet ekranını açın.
2. Outlook'taki e-posta metnini seçip sistem **Paylaş** veya **Metni işle** seçeneğinde Takip Kolay görünüyorsa seçin. Outlook paylaşmayı desteklemiyorsa içeriği kopyalayıp Takip Kolay'a yapıştırın.
3. Konuyu ve notu düzenleyin. Örnek not: Ruhsat evraklarının tamamlanmasını takip et.
4. Hatırlatma istiyorsanız **Hatırlatma: tarih seçilmedi** düğmesinden gün ve saati bilinçli seçin. Örnek: 20 Kasım 2026. Tarih istemiyorsanız değiştirmeyin.
5. **Görevi kaydet** düğmesine dokunun; bildirim izni istenirse karar verin.
6. Kayda dokunarak tam yazışma içeriğini açın; tamamlayın veya düzenleyin.

Bu kaynak kodda veya herkese açık sorun bildirimlerinde **gerçek şirket yazışması, e-posta adresi ya da belge paylaşmayın**.

## Geliştirici notu

Android Studio / SDK bulunmayan ortamda kaynak oluşturulmuştur. Uçtan uca gerçek cihaz, Outlook paylaşımı, bildirim ve derleme testi yapılmadan kararlı sürüm olarak yayımlanmamalıdır.
