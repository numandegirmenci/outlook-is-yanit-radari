# Outlook İş ve Yanıt Radarı

**Klasik Outlook için ücretsiz, açık kaynaklı e-posta ve iş takip makroları.** Gönderdiğiniz bir e-postanın yanıtını beklerken veya gelen bir e-postaya ilişkin yapacağınız işi takip ederken Excel doldurmanız gerekmez.

> **Durum: v0.7 Genel Beta / Pilot.** Kaynak kodu paylaşım için hazırlanmıştır; farklı Outlook sürümlerinde uçtan uca test edilmiş, kararlı bir ürün değildir. Özellikle otomatik yanıt eşleştirmesi ve HTML raporu bu sürümde kullanıcı testine ihtiyaç duyar.

## Ürün amacı — Outlook içinde, veriyi dışarı taşımadan iş yönetimi

Bu proje, iş yerinde Notion, Trello veya benzeri üçüncü taraf servisleri kullanamayan; fakat klasik Outlook'ta çok sayıda talep, yazışma ve bekleyen işi yöneten kişiler için tasarlanır. Amaç **ayrı bir iş takip hesabı ya da sunucu gerektirmeden Outlook'un kendi içinde iş yönetimi** sağlamaktır.

- **Mevcut iş akışı:** Gelen veya gönderilen e-postadan bir tıklamayla takip aç; yanıt bekle, yapılacak iş veya inceleme durumuna geçir; iş bitince tamamla.
- **Bağlam:** Takibi ilgili yazışmaya bağla. Kullanıcı neyin neden beklediğini görev üstünden görebilsin.
- **Kontrol kullanıcıda:** Hatırlatma tarihi kullanıcı seçer; hiçbir mesaj otomatik gönderilmez veya silinmez.
- **Veri ilkesi:** Harici API, üçüncü taraf iş takip servisi, reklam, analiz, telemetri veya uzaktan yapay zeka işleme **ürünün hedef mimarisinde yoktur**. GitHub yalnızca açık kaynak kodu dağıtmak içindir, kullanıcı e-postalarını toplamak için değil.
- **Kurumsal gerçeklik:** Outlook görevleri Exchange/Microsoft 365 hesabında tutulursa şirketin mevcut Microsoft altyapısıyla eşitlenebilir. Bu, "yalnızca cihazda saklanır" demek değildir.

**V0.7 gizlilik sınırı:** Mevcut beta sürümü seçili e-postadan en fazla 10.000 karakteri takip görevinin `Body` alanına kopyalar; yanıt metinlerinden de parçalar kaydedebilir. Böylece içerik zaten kurumsal Outlook/Exchange ortamı içindeyken **ikinci bir görev kaydı içinde çoğaltılır**. Bunu daha veri-minimum bir sürümde, asıl e-postayı gerektiğinde okuyup görevde yalnızca bağlantı ve durum bilgisi tutarak azaltmayı hedefliyoruz. Henüz uygulanmamış bir özelliği mevcut gibi göstermiyoruz.

Ayrıntılı yaklaşım: [Ürün vizyonu ve veri sınırları](docs/URUN_VIZYONU.md).

## Neler yapar?

- **Yanıt bekliyorum:** Seçilen e-posta için takip oluşturur.
- **Yapacağım iş:** Seçilen e-postadan bir görev oluşturur.
- **İşlerim:** Açık, geciken, yanıt bekleyen ve tamamlanan işleri Outlook'ta **gönderilmeyen bir e-posta önizlemesi** olarak listeler. E-postaların metninden kısa parçalar gösterir.
- **Yanıtları kontrol et:** Outlook'un varsayılan Gelen Kutusu ve alt klasörlerindeki uygun iletileri inceleyerek takip durumunu **İncelenecek** olarak değiştirebilir. **İşleri otomatik tamamlamaz.**
- **Hatırlatma taslağı:** Takip edilen yazışmayı temel alan **gönderilmeyen** bir taslak hazırlar.
- **Tarih kontrolü:** Hatırlatma tarihi kullanıcının seçimine bırakılmıştır; `YOK` ile tarihsiz takip oluşturulabilir.

## Android mobil — deneysel geliştirme

Outlook Android'de görev takibi için ayrı bir yerel yardımcı uygulama üzerinde çalışılıyor: [Takip Kolay Android (0.1.0-alpha)](mobile/android/README.md).

Bu çözüm mevcut VBA makrosunun Android'de çalışması anlamına gelmez. Android uygulaması yalnızca sistemin paylaştığı veya kullanıcının yapıştırdığı metinle görev oluşturur; Outlook'un tüm yazışmalarına doğrudan erişmez. Görevler cihazda şifrelenir; tarih kullanıcı seçer. Mobil sürüm henüz beta öncesi test aşamasındadır.

## Uyumlu sistem

- Windows'ta **klasik Microsoft Outlook** (VBA desteği ve şirketin güvenlik politikalarının izin vermesi gerekir).
- **Yeni Outlook**, Outlook web ve mobil sürümler desteklenmez.
- Excel veya ayrı bir hesap gerekmez.
- ZIP içindeki VBA kodunu Outlook'a **elle** eklemek gerekir; bir .exe veya tek tıklamalı yükleyici yoktur.

## Kurulum — kısa yol

1. Mevcut Outlook VBA kodunun yedeğini alın.
2. Outlook'ta `Alt + F11` ile VBA düzenleyicisini açın.
3. `File → Import File` ile `src/IsYanitRadari.bas` dosyasını ekleyin. Modül içeride **TakipKolay** adıyla görünür.
4. `src/ThisOutlookSession.txt` içindeki olay prosedürlerini **ThisOutlookSession** bölümüne uygun şekilde ekleyin. `Application_Startup` veya `Application_NewMailEx` zaten varsa **ikinci kopya oluşturmayın**.
5. `Debug → Compile VBAProject`, ardından kaydedip Outlook'u tamamen kapatın ve yeniden açın.
6. Bir **örnek e-posta** üzerinde `Alt + F8` menüsünden `Takip_1_YanitBekliyorum` veya `Takip_2_YapacagimIs` komutunu deneyin, sonra `Takip_3_Islerim` raporunu açın.

Daha ayrıntılı kurulum ve hata çözümü: [Kurulum rehberi](docs/KURULUM.md).

## Başlıca komutlar

| Outlook makrosu | İşlev |
|---|---|
| `Takip_1_YanitBekliyorum` | Yanıt için takip kaydı oluştur |
| `Takip_2_YapacagimIs` | Yapılacak iş oluştur |
| `Takip_3_Islerim` | HTML takip raporunu aç |
| `Takip_4_YanitlariKontrolEt` | Gelen yanıtları elle kontrol et |
| `Takip_5_Tamamlandi` | Seçilen işi tamamla |
| `Takip_6_TarihiDegistir` | Tarihi değiştir |
| `Takip_7_SistemKontrolu` | Takip klasörünü ve kayıt sayısını kontrol et |
| `Takip_8_HatirlatmaTaslagi` | Gönderilmeyen hatırlatma taslağı aç |

## Rapor örneği

[Örnek HTML rapor](examples/RAPOR_TEMSILI.html) **tamamen temsili veriler** içerir. Gerçek görev ve e-posta içeriği içermez.

## Gizlilik

Makro doğrudan üçüncü taraf bir sunucuya belge göndermek için tasarlanmamıştır ve otomatik e-posta **göndermez**. Ancak takip görevleri Outlook'un **Görevler / Takip Kolay** klasörüne yazılır; e-posta içeriğinden kopyalanan metinler bu görevlere eklenir ve Microsoft 365/Exchange ile **senkronize olabilir**. Kurumsal verilerle kullanmadan önce [gizlilik ve güvenlik notlarını](docs/GIZLILIK_VE_GUVENLIK.md) okuyun.

## Bilinen sınırlamalar

Yanıt taraması son **45 gün** ile sınırlıdır; performans için klasör, ileti ve süre sınırları da bulunur. Bu yüzden her yanıtın bulunacağı garanti değildir. Yalnızca varsayılan Gelen Kutusu ve alt klasörleri taranır. E-posta eşleştirmesi `ConversationID`'ye dayanır.

Geniş bir kullanıcı kitlesine dağıtım için Outlook VBA ideal yöntem değildir; uzun vadede imzalı eklentiye geçiş düşünülebilir. [Microsoft'un dağıtım açıklaması](https://learn.microsoft.com/en-us/troubleshoot/outlook/deployment/manage-distribute-outlook-vba-project).

## Katkı ve geri bildirim

Hata bildirmek için [Issue şablonunu](.github/ISSUE_TEMPLATE/hata-bildirimi.md) kullanın. **Gerçek e-posta, müşteri adı, alıcı adresi, şirket belgesi veya ekran görüntüsündeki hassas bilgileri paylaşmayın.**

## Lisans

MIT — [LICENSE](LICENSE).
