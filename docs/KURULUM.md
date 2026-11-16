# Kurulum ve ilk test

## Önkoşullar

- Windows üzerinde **klasik Outlook** kullanın. Yeni Outlook ile çalışmaz.
- Kurumsal hesaplarda makro güvenlik politikaları nedeniyle kurulum veya çalıştırma engellenebilir. Kurumunuzun politikasını **aşmayın**.
- Kurulum öncesinde mevcut VBA modüllerini `File → Export File` ile yedekleyin.

## İlk kez kurulum

1. ZIP arşivini bilgisayarda çıkarın.
2. Klasik Outlook'ta `Alt + F11`.
3. Menüden `File → Import File` → `src/IsYanitRadari.bas` dosyasını seçin.
4. Sol taraftan `ThisOutlookSession` kod sayfasını açın, `src/ThisOutlookSession.txt` içindeki olay kodunu ekleyin.
5. **Önemli:** `Application_Startup` / `Application_NewMailEx` zaten varsa ikinci prosedür açmak yerine, var olanın içine ilgili çağrıları birleştirin; aksi halde `Ambiguous name detected` hatası alabilirsiniz.
6. `Debug → Compile VBAProject` ile derlemeyi kontrol edin. Hata varsa yayına/takibe geçmeden düzeltin.
7. `Ctrl + S` ile kaydedip Outlook'u **tamamen kapatın**, yeniden açın.

## Mevcut deneme sürümünden güncelleme

- Eski `TakipKolay` modülünü önce dışa aktararak yedekleyin. Sonra yalnızca bu modülü kaldırıp yenisini yükleyin.
- Eski özel görev klasörünü silmeyin. Bu sürüm onunla uyumludur.
- `ThisOutlookSession` içinde çalışan ve aynı olay adını kullanan başka kod varsa üzerine yazmayın veya kopyalamayın; önce birleştirin.

## Kullanım: üç adım

1. Outlook'ta bir e-posta seçin.
2. `Alt + F8` → `Takip_1_YanitBekliyorum` (**karşı taraftan yanıt bekliyorum**) ya da `Takip_2_YapacagimIs` (**benim işim var**).
3. Tarih/saat sorulduğunda `24.11.2026 14:30` biçiminde tarih girin; tarih istemiyorsanız `YOK` yazın. Boş/İptal takip oluşturmaz. Ardından `Takip_3_Islerim` ile raporu açın.

## Hata durumunda

- **'İşlerim' sıfır gösteriyor:** `Takip_7_SistemKontrolu` komutunu çalıştırın. Takipler Outlook Görevler altındaki **Takip Kolay** klasörüne yazılır.
- **'Type mismatch / Hata 13':** Kullandığınız sürümü, tam hata metnini, hangi makroda olduğunu ve `Debug` ile durulan satırı (varsa) paylaşın. V0.7'de kayıt bazlı hata kontrolü vardır, ancak farklı Outlook yapılandırmalarında yine hata oluşabilir.
- **Yanıt bulunmuyor:** Yanıtın varsayılan Gelen Kutusu veya alt klasörlerinde olduğunu ve aynı konuşma zincirinde geldiğini kontrol edin. 45 gün ve performans sınırları nedeniyle bazı iletiler atlanabilir.
- **Makrolar görünmüyor:** Modül yanlışlıkla `ThisOutlookSession`'a yapıştırılmış olabilir; `.bas` dosyasını **modül olarak** içeri aktarın.

**Güvenlik:** Makro ayarlarında *tüm makroları koşulsuz etkinleştirmeyin*. Kurumunuzun politikalarına uyun.
