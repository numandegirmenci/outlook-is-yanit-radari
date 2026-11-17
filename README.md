# Outlook İş ve Yanıt Radarı

**Outlook'ta çok sayıda yazışma ve işi takip edenler için ücretsiz, açık kaynaklı bir iş masası.**

> **v0.9.1 Genel Beta — 18 Kasım 2026.** Kod yayımlanmıştır, ancak yeni form arayüzünün farklı Outlook kurulumlarında canlı çalışması henüz doğrulanmamıştır. Önce test e-postalarıyla deneyin.

## Ne işe yarıyor?

Her gün e-postayla birinden evrak, bir başkasından cevap veya onay bekleyen biri için Outlook'taki işlerin kolayca kaybolması mümkündür. Bu araç, bunları **başka bir iş takip servisine taşımadan** mevcut Outlook'ta izlemeyi hedefler.

- **Yapacağım İş:** Gelen veya gönderilen e-postadan kendi işini kaydet.
- **Yanıt Bekliyorum:** Birinden cevap beklediğin yazışmayı izlemeye al.
- **İş Masam:** İşlerini tek pencerede gör; **Yapılacak / Yanıt Bekleyen / İncelenecek / Tamamlanan** filtrelerinden birini seç.
- **E-posta içeriğini oku (v0.9.1):** İşin üzerine tıkladığında konuyu, göndereni, alıcıyı, mesaj metnini ve Outlook'un bulabildiği ilgili yanıtları **aynı pencerede** gör.
- **E-postayı Aç:** Orijinal mesajı aç. Mesaj taşındıysa veya bulunamıyorsa konusuna göre Outlook araması için önce onay ister.
- **Hatırlatma:** Görev açarken `1` (1 saat), `2` (yarın), `3` (3 gün), `4` (1 hafta), `5` (özel tarih), `0` (hatırlatmasız) seç. Program tarihi kendiliğinden belirlemez.
- **Hatırlatma Taslağı / Tamamlandı / Yenile:** İş Masam'daki düğmelerden kullan. **Hiçbir e-posta otomatik gönderilmez.**

### Kısa örnek

Bir iş arkadaşından evrak bekliyorsun. İlgili e-postayı **Yanıt Bekliyorum** olarak kaydet. Sonra **İş Masam** ekranından göreve tıkla: mesajın içeriğini ve Outlook'un aynı konuşmada bulabildiği yanıtları gör. Gerekirse orijinal e-postayı aç veya göndermeden önce kontrol edeceğin bir hatırlatma taslağı hazırla.

## Kimler için?

Şirket bilgisayarında Notion, Trello veya başka bir harici iş takip servisine erişemeyen; yoğun Outlook trafiğini ayrıca Excel'e aktarmak istemeyen çalışanlar için. **Klasik Windows Outlook + kurum tarafından izin verilmiş VBA** gerekir. Yeni Outlook, Outlook Web ve mobil Outlook desteklenmez.

## Kurulum — v0.9.1'de önemli değişiklik

Önceki test paketindeki `frmRadarIsMasam.frm` dosyası bazı bilgisayarlarda **&H80004005** hatası verdiği için **artık dağıtılmıyor**. Form **bir defa Outlook'un VBA editöründe boş olarak oluşturuluyor**; arayüz düğmelerini kod kendisi ekliyor.

**İlk kurulum:**

1. [Code → Download ZIP](https://github.com/numandegirmenci/outlook-is-yanit-radari/archive/refs/heads/main.zip) ile projeyi indirip aç.
2. Klasik Outlook'ta `Alt + F11` → **File → Import File**: önce `src/IsYanitRadari.bas`, sonra `src/RadarIsMasam.bas` dosyasını ekle.
3. **Insert → UserForm** ile boş form oluştur; `F4` (Properties) → **(Name)** alanını `frmRadarIsMasam` yap.
4. Forma `F7` ile **kod görünümünde** geç; `src/frmRadarIsMasam_KODU.txt` dosyasının **tüm içeriğini** yapıştır. Forma elle buton ekleme.
5. **Debug → Compile VBAProject** ile kontrol et; `Ctrl + S` ile kaydet ve Outlook'u yeniden aç.
6. Önce bir **test e-postasında** `Takip_2_YapacagimIs`, sonra `Radar_11_IsMasam` çalıştır. Görev satırını seçtiğinde e-posta metnini görebilmelisin.

**Mevcut v0.8.3 kullanıcıları:** `TakipKolay` modülüne, `ThisOutlookSession` koduna veya Görevler içindeki **Takip Kolay** klasörüne dokunmayın. `RadarIsMasam` modülünü önce yedekleyin, ardından yeni dosyayla değiştirin; yalnızca formu yukarıdaki şekilde ekleyin. Form varsa tekrar oluşturmayın.

**[Adım adım kurulum ve hata çözümü](docs/KURULUM.md)** · **[Form kurulum rehberi](docs/FORM_KURULUMU.md)** · **[Test planı](docs/TEST_PLANI.md)**

## Gizlilik — bilgiler nereye gidiyor?

Proje **bağımsız bir üçüncü taraf iş takip sunucusu, telemetri veya uzaktan yapay zeka servisi kullanmaz**. Yeni görevlerde e-postanın tam gövdesi yeniden kopyalanmaz; form, seçilen e-postanın metnini Outlook'tan gerektiğinde okur.

**Ancak:** Outlook verileri kurumunuzun Exchange/Microsoft 365 sunucularıyla eşitlenebilir. Bu nedenle "veriler bilgisayardan hiç çıkmaz" iddiası doğru değildir. Önceki sürümlerde oluşturulmuş görevlerin içinde yazışma kopyaları olabilir; otomatik silinmez. Taslaklar Outlook'un eşitleme kurallarına tabidir. **Şirketin VBA/makro politikasını aşmayın.**

[Detaylı gizlilik ve güvenlik açıklaması](docs/GIZLILIK_VE_GUVENLIK.md).

## Bilinen sınırlar

- **Kurulumda form elle oluşturulur**; bu henüz tek tıklamalı bir eklenti değildir.
- Konuşma detaylarında Outlook'un bulabildiği en fazla **45 ileti** incelenir; en son **16 ileti** gösterilir; her iletiden en fazla **1.700 karakter** alınır. Yanıtların tamamı bulunamayabilir.
- Taşınan, silinen veya başka depoya giden kaynak e-postanın kimliği çalışmayabilir; konu araması kesin eşleşme garantilemez.
- Listede en fazla **600 görev** taranır. Görev panelinde yapılan **Tamamlandı** işlemi gerçek görev durumunu değiştirir.
- Kullanıcı hatırlatma taslağını kendi kontrol edip göndermelidir; alıcı, metin ve önceki yazışmalar özellikle incelenmelidir.
- **Beta:** Formun gerçek Outlook üzerinde derlendiği ve her senaryoda sorunsuz çalıştığı henüz doğrulanmadı.

## Hata bildirme ve katkı

Sorunu [GitHub Issues](https://github.com/numandegirmenci/outlook-is-yanit-radari/issues) üzerinden bildirebilirsin. **Gerçek e-posta içeriği, alıcı, müşteri veya kurum belgelerini paylaşma.** Ekran görüntüsündeki hassas bilgileri gizle.

**Lisans:** [MIT](LICENSE). **Resmi Microsoft ürünü değildir.**
