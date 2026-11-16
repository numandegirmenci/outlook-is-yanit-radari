# Kurulum ve ilk kullanım — Outlook İş ve Yanıt Radarı v0.8.3

Bu rehber, teknik terimlerle uğraşmadan ilk kurulumu yapabilmen için hazırlandı.

## Önce kontrol et

- Bilgisayarında **klasik Outlook** olmalı. Yeni Outlook veya tarayıcıda Outlook çalışmaz.
- Şirketin makro çalıştırmayı engelliyorsa **güvenlik politikasını aşma**; bilgi işlem biriminden izin al.
- Mevcut bir VBA projen varsa önce yedekle.

## İlk kurulum

1. GitHub'da **Code → Download ZIP** düğmesiyle kodları indir, ZIP'i bir klasöre aç.
2. Outlook açıkken **`Alt + F11`** tuşlarına bas.
3. VBA penceresinin üst menüsünden **File → Import File** seç ve `src/IsYanitRadari.bas` dosyasını yükle. `TakipKolay` adlı modül görünür.
4. Aynı şekilde `src/RadarIsMasam.bas` dosyasını da yükle.
5. Yanıt algılama olaylarını kullanmak istiyorsan `src/ThisOutlookSession.txt` dosyasını incele. `ThisOutlookSession` altında zaten **Application_Startup** ya da **Application_NewMailEx** varsa aynı adlı ikinci prosedür oluşturma; kodların uygun şekilde birleştirilmesi gerekir. Bu konuda emin değilsen bu adımı atla, yanıtları `Takip_4_YanitlariKontrolEt` ile elle kontrol et.
6. **Debug → Compile VBAProject** seç. Hata varsa kodu çalıştırma; hata metnini not al.
7. `Ctrl + S` ile kaydet, Outlook'u tamamen kapatıp yeniden aç.
8. Deneme mesajı seç. `Alt + F8` → `Takip_2_YapacagimIs` → Çalıştır. Açılan menüde `2` yazarsan ertesi gün 09.00'a hatırlatma konur.
9. `Alt + F8` → `Radar_11_IsMasam`. Açılan görev listesinden yeni deneme kaydını kontrol et.

## v0.7.x veya v0.8.x'ten güncelleme

1. `Alt + F11` → **Modules** altındaki `TakipKolay` modülüne sağ tık → **Export File** ile yedekle.
2. Eski `TakipKolay` modülünü **Remove** ile kaldır. Bu işlem Outlook görevlerini silmez.
3. Yeni `src/IsYanitRadari.bas` dosyasını **Import File** ile yükle.
4. Bilgisayarında `RadarIsMasam` modülü yoksa `src/RadarIsMasam.bas` dosyasını da yükle. V0.8.2 kullanıyorsan mevcut modül aynı sürümdür; yeniden yüklemen gerekmez.
5. `ThisOutlookSession` bölümünü sırf güncelleme yapıyorsun diye değiştirme.
6. Derle, kaydet ve Outlook'u yeniden aç.

**Görevler → Takip Kolay** klasörünü silme. Eski takipler aynı klasörde kalır.

## Hızlı tarih seçimi

`1`: 1 saat sonra. `2`: yarın 09.00. `3`: 3 gün sonra 09.00. `4`: bir hafta sonra 09.00. `5`: tarih ve saatini yaz (ör. `25.11 15:30`). `0`: hatırlatma olmadan kaydet. **Boş giriş / İptal:** vazgeç.

Kayıt veya tarih değişikliği için seçimi kullanıcı yapar; makro tarihi kendiliğinden belirlemez.

## Bir şey çalışmazsa

- VBA derleme hatasının **tam metnini** not al.
- İş listesi görünmüyorsa `Takip_7_SistemKontrolu` veya `Takip_10_DurumleriKontrolEt` çalıştır.
- Outlook görünümü hata verirse görev klasörünü silme; görünüm ile görev kaydı farklı şeylerdir.
- Yanıt algılaması her klasörde ve her hesabın mesajlarında kusursuz çalışmayı garanti etmez.
- GitHub'da hata bildirirken **gerçek iş e-postası, adres, sözleşme, evrak veya kişisel bilgi paylaşma**.
