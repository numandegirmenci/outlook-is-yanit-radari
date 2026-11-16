# Outlook İş ve Yanıt Radarı — v0.7.0-beta

**Açık kaynaklı genel beta / pilot.** Klasik Outlook'ta seçilen e-postalardan yapılacak iş veya yanıt bekleme takibi oluşturur; hatırlatma zamanını kullanıcı belirler, Outlook içinde HTML iş özeti ve gönderilmeyen hatırlatma taslağı hazırlar.

**Kurulum:** ZIP'i çıkarıp `src/IsYanitRadari.bas` dosyasını Outlook VBA'ya modül olarak içeri aktarın. Olay kodunu `ThisOutlookSession` içine ekleyin. `docs/KURULUM.md` dosyasını okuyun.

**Gereksinim:** Windows + klasik Outlook + izinli VBA. **Yeni Outlook desteklenmiyor.**

**Beta uyarısı:** V0.7 gerçek Outlook ortamında geniş kapsamlı test edilmedi. Kritik işlerde yalnızca bu yazılıma güvenmeyin; otomatik yanıt tespiti 45 gün / performans sınırları nedeniyle eksik olabilir.

**Gizlilik:** Takip görevleri Outlook/Exchange ile senkronize olabilir; e-posta metinleri görev gövdelerine kopyalanabilir. Kurum izinleri gerekli olabilir. Makro kendi başına e-posta göndermez.


## v0.7.1 (genel beta hata düzeltmesi, 16 Kasım 2026)

- **İşlerim** raporunda yapılacak işler, tarihleri gecikmiş veya bugün olsa da kendi kategorilerinde gösterilir.
- Aynı e-posta üzerinden bir **Yanıt Bekliyorum** ve bir **Yapacağım İş** takibi ayrı tutulabilir.
- Yeni `Takip_10_DurumleriKontrolEt` komutu durum sayılarını gösterir.
- Görevler silinmez veya otomatik taşınmaz; Outlook üzerinde uçtan uca test henüz tamamlanmadı.
- Makro dosyası Windows Türkçe (CP1254) kodlamalıdır; GitHub'dan **Download raw file** veya ZIP ile indirilmelidir.
