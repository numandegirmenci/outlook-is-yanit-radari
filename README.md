# Outlook İş ve Yanıt Radarı

**Outlook'taki e-postalarını yapılacak işlere dönüştür. Yanıt beklediğin konuları unutma. İşlerini başka bir uygulamaya taşımadan takip et.**

> **Güncel sürüm: v0.8.3 Genel Beta (16 Kasım 2026).** Ücretsiz ve açık kaynaklıdır. Klasik Outlook için geliştirilmiştir. Tüm bilgisayarlarda test edilmiş kararlı bir ürün değildir.

## Bu program ne işe yarıyor?

Gün boyunca birçok kişiye e-posta gönderdiğini düşün. Birinden evrak, diğerinden onay, bir başkasından ödeme bilgisi bekliyorsun. Araya yeni e-postalar girdikçe neyi, kimden beklediğini hatırlamak zorlaşıyor.

**İş ve Yanıt Radarı**, bu yazışmaları Outlook'tan çıkmadan takip etmen için hazırlanmış bir araçtır:

- **Yanıt Bekliyorum:** Birine e-posta gönderdin, cevap bekliyorsun.
- **Yapacağım İş:** Bir e-posta geldi, yapman gereken bir işlem var.
- **İş Masam:** Açık işleri tek yerde görürsün. İşler `TODO` (Yapılacak), `WAIT` (Yanıt bekleniyor), `REVIEW` (İncelenecek) ve `DONE` (Tamamlandı) olarak gruplanabilir.
- **Gelen yanıtı kontrol et:** Takip edilen yazışmaya cevap gelirse görev incelemeye alınabilir. Cevap geldi diye iş otomatik tamamlanmaz.
- **Hatırlatma:** Zamanı sen belirlersin; hatırlatmasız görev de oluşturabilirsin.
- **Kaynak yazışmayı aç:** Takip kaydından ilgili e-postaya geri dönebilirsin.
- **İşlerim raporu:** Bekleyen, yapılacak ve geciken işleri özetler. E-posta olarak *gönderilmez*.

**Örnek:** Bir kurumdan ruhsatla ilgili belge istedin. E-postayı seçip `Yanıt Bekliyorum` dedin, yarına hatırlatma koydun. Ertesi gün İş Masam'da bu yazışmayı bulabilirsin. Yanıt gelirse ayrıca kontrol edebilirsin.

## v0.8.3 yeniliği: Tarih yazmak yerine rakam seç

Yeni iş açarken veya var olan işin tarihini değiştirirken hızlı seçim menüsü çıkar:

| Seçim | Ne olacak? |
|---|---|
| `1` | 1 saat sonra |
| `2` | Yarın 09.00 |
| `3` | 3 gün sonra 09.00 |
| `4` | 1 hafta sonra 09.00 |
| `5` | Kendi gün ve saatini yaz: `25.11 15:30` |
| `0` | Tarih ve hatırlatma olmadan kaydet |
| Boş / İptal | İşlemden vazgeç; var olan tarih değişmez |

**Program senin yerine tarih seçmez.** Saat 09.00 ancak ilgili seçeneği bilerek tercih edersen kullanılır. Özel tarih girişinde saat yazmazsan 09.00 kabul edilir.

## Kimler kullanabilir?

Özellikle e-postayla çok iş yürüten; kurumundaki kısıtlamalar yüzünden Notion, Trello gibi harici platformlara veri gönderemeyen veya Excel'de ayrıca iş listesi tutmak istemeyen kişiler için.

- Windows'ta **klasik Microsoft Outlook** gerekir.
- **Yeni Outlook, Outlook Web ve mobil Outlook** desteklenmez.
- Outlook VBA makrolarının kurumun güvenlik politikaları tarafından **izin verilmiş** olması gerekir.
- Ayrı bir iş yönetimi hesabı veya kendi sunucumuzu kullanman gerekmez.

## Kurulum (ilk kez kullananlar)

1. Bu sayfadaki **Code → Download ZIP** üzerinden kodları indir, ZIP dosyasını aç.
2. Klasik Outlook'ta **`Alt + F11`** ile VBA editörünü aç.
3. **File → Import File** üzerinden önce `src/IsYanitRadari.bas`, ardından `src/RadarIsMasam.bas` dosyalarını içeri aktar.
4. Otomatik gelen yanıt kontrolünü kullanmak istiyorsan `src/ThisOutlookSession.txt` içindeki kodu mevcut `ThisOutlookSession` bölümüne **çakışma oluşturmadan** ekle. Bu aşamayı anlamıyorsan ilgili kodu körü körüne yapıştırma.
5. **Debug → Compile VBAProject** ile derle, `Ctrl + S` ile kaydet.
6. Outlook'u tamamen kapatıp yeniden aç.
7. Deneme e-postası seç, `Alt + F8` ile `Takip_2_YapacagimIs` komutunu çalıştır ve tarih için `0` ya da `2` seç. Sonra `Radar_11_IsMasam` komutunu çalıştır.

**Önceki sürümden güncelliyorsan:** Eski VBA modüllerini **Export File** ile yedekle. Mevcut `TakipKolay` modülünü kaldırıp yeni `IsYanitRadari.bas` dosyasını yükle. V0.8.2'nin `RadarIsMasam` modülünü kullanıyorsan o modül aynı kaldı. **Outlook Görevler → Takip Kolay klasörünü silme.** Kurulum ayrıntıları: [Türkçe kurulum rehberi](docs/KURULUM.md).

## En sık kullanılan komutlar

| Makro | Ne işe yarar? |
|---|---|
| `Takip_1_YanitBekliyorum` | E-postadan yanıt bekleyen iş açar |
| `Takip_2_YapacagimIs` | E-postadan yapılacak iş açar |
| `Radar_11_IsMasam` | Outlook'ta gruplu iş listesini açar |
| `Takip_3_Islerim` | İşlerim özet raporunu gösterir |
| `Takip_6_TarihiDegistir` | İşe yeni tarih/saat belirler |
| `Radar_12_KaynakYazismayiAc` | Seçili görevin kaynak e-postasını açar |
| `Radar_13_KategoriDegistir` | Görevin kategorisini değiştirir |
| `Takip_8_HatirlatmaTaslagi` | Gönderilmemiş hatırlatma taslağı hazırlar |
| `Takip_10_DurumleriKontrolEt` | Durumlara göre iş sayılarını gösterir |

Tüm komutların listesi ve test adımları: [Test planı](docs/TEST_PLANI.md).

## Verilerim başka yere gider mi?

**Amaç, e-posta içeriğini bağımsız bir üçüncü taraf iş takip servisine göndermeden çalışmaktır.**

- Bu VBA kaynak kodunda **harici bir iş takip sunucusuna gönderim, telemetri ve reklam servisi yoktur**.
- **v0.8'den itibaren yeni oluşturulan görevlerde** kaynak e-posta metni ayrıca görev gövdesine kopyalanmaz; gerekli bağlantı ve iş bilgileri tutulur. İlgili yazışma gerektiğinde Outlook'tan okunur.
- **Önemli ayrım:** Outlook e-posta ve görev verileri, hesabın yapılandırmasına göre **şirketinin Exchange/Microsoft 365 altyapısına senkronize olabilir**. Bu, tamamen çevrimdışı veya yalnızca bilgisayarda saklama anlamına gelmez.
- Eski sürümlerde oluşturulmuş görevlerin gövdesinde kopya yazışma metinleri bulunabilir; bu sürüm onları otomatik silmez.
- Hatırlatma taslakları **kendiliğinden gönderilmez**. Outlook'ta açılan taslaklar Outlook'un kendi kayıt ve senkronizasyon kurallarına tabi olabilir.
- **Gerçek e-postaları, müşteri bilgilerini veya şirket evraklarını GitHub Issues'a yükleme.**

Detaylar: [Gizlilik ve güvenlik](docs/GIZLILIK_VE_GUVENLIK.md).

## Bilinen sınırlamalar

Bu bir **beta** sürümüdür; bazı özellikler farklı Outlook kurulumlarında hatalı çalışabilir. İş Masam tek gruplu Outlook görev görünümünü kullanır; eski filtreli görünümlerin Outlook tarafından reddedildiği durumlar yaşandı. Kaynak e-posta taşındıysa kimliği değişebilir ve bağlantı açılamayabilir. Yanıt kontrolü 45 günlük ve performans sınırları nedeniyle her yanıtı bulmayı garanti etmez. Kurumun makro politikalarını aşmaya çalışmayın.

**Proje Microsoft'a ait resmi bir ürün değildir.** Hata ve geliştirme önerilerini kişisel/kurumsal bilgileri çıkartarak [GitHub Issues](https://github.com/numandegirmenci/outlook-is-yanit-radari/issues) üzerinden paylaşabilirsin.

## Projenin durumu

v0.8.3 kaynakları yayımlandı ve yapısal dosya kontrolleri yapıldı; **gerçek Outlook üzerinde bu sürümün bütün işlemleri uçtan uca doğrulanmadı**. Şimdilik örnek e-postalarla test edilmelidir.

**Lisans:** [MIT](LICENSE) — ücretsiz, açık kaynak kodlu.

**Hedefimiz:** Outlook içinde, Notion benzeri iş takip kolaylığı; üçüncü taraf platformlara iş verisi göndermeden ve yeni bir hesap oluşturmadan.
