# Sürüm notları

## v0.7.0-beta — Genel pilot kaynak paketi

- Kullanıcının belirlediği hatırlatma tarihi/saatine geçildi; `YOK` seçeneği eklendi.
- Gelen yanıtla mevcut hatırlatma tarihinin otomatik değiştirilmesi kaldırıldı.
- `İşlerim` raporunda görev bazında hata toleransı ve yazışma özeti yer alıyor.
- Hatırlatma taslağına yazışma içeriği eklendi (9000 karakter sınırı).
- Manuel/açılış yanıt taraması son 45 güne çıkarıldı; performans sınırları eklendi.
- Marka adı **Outlook İş ve Yanıt Radarı** olarak değişti. Kod uyumluluğu nedeniyle eski VBA modül adı `TakipKolay` ve Outlook görev klasörü `Takip Kolay` olarak korunuyor.
- Durum: **Beta**. Gerçek Outlook üzerinde v0.7 tam uçtan uca test edilmedi.


## v0.7.1 (genel beta hata düzeltmesi, 16 Kasım 2026)

- **İşlerim** raporunda yapılacak işler, tarihleri gecikmiş veya bugün olsa da kendi kategorilerinde gösterilir.
- Aynı e-posta üzerinden bir **Yanıt Bekliyorum** ve bir **Yapacağım İş** takibi ayrı tutulabilir.
- Yeni `Takip_10_DurumleriKontrolEt` komutu durum sayılarını gösterir.
- Görevler silinmez veya otomatik taşınmaz; Outlook üzerinde uçtan uca test henüz tamamlanmadı.
- Makro dosyası Windows Türkçe (CP1254) kodlamalıdır; GitHub'dan **Download raw file** veya ZIP ile indirilmelidir.
