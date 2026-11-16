# Ürün vizyonu — Outlook İş ve Yanıt Radarı

## Kimin için?

Kurumsal bilgisayarında yeni bir SaaS hesabı açamayan, Notion/Trello ile entegrasyon kuramayan veya iş e-postalarını üçüncü taraf sistemlere taşımak istemeyen; ancak Outlook üzerinde çok sayıda paralel işi takip eden çalışanlar.

**Ürün ilkesi:** Outlook'taki yazışmaları ayrı bir iş takip platformuna kopyalamadan; Outlook'un izin verilen yerleşik nesneleri üzerinden takip etmek.

## Temel kullanım akışı

1. Kullanıcı seçili e-posta üzerinden "Yanıt bekliyorum" ya da "Benim işim" der.
2. İsterse hatırlatma tarihi ve kısa not seçer; tarihi sistem zorunlu kılmaz.
3. Outlook görevi durum, kişi, tarih ve asıl iletiye referansla açılır.
4. Aynı yazışma yeniden geldiğinde görev "İncelenecek" olarak işaretlenir; yanıt geldi diye kendiliğinden tamamlanmaz.
5. "İşlerim" görünümünde gecikenler, yanıtlardan inceleme bekleyenler, yapılacaklar ve tamamlananlar yer alır.
6. Hatırlatma taslağı kullanıcıya gösterilir, kendi onayı olmadan gönderilmez.

## Gizlilik mimarisi — hedef

- **Üçüncü taraf aktarımı yok:** Notion, Trello, Zapier, Make, herhangi bir geliştirici sunucusu, çevrimiçi OCR/AI, harici API veya telemetri kullanılmaz.
- **Zorunlu yeni hesap yok:** Outlook'un bağlı olduğu mevcut hesap yeterlidir.
- **İçerik çoğaltmayı azalt:** Takip kaydında minimum üstveri ve kaynak e-posta referansları tut; konuşmanın tam içeriğini depolamadan, yalnızca kullanıcının görüntülediği sırada Outlook'tan oku.
- **Saklama ayrımı:** Outlook'un kendi e-posta ve Görevler verileri şirketin mevcut Exchange/Microsoft 365 sunucularına eşitlenebilir. "Üçüncü tarafa gönderilmez" ile "cihazdan hiç çıkmaz" aynı şey değildir.
- **Kullanıcı kontrolü:** Otomatik gönderim, silme, taşıma, veri dışa aktarma yok; kullanıcı açıkça istemeden paylaşım yapılmaz.
- **Kurumsal uyumluluk:** Makro/eklenti kurulumunu şirket güvenlik politikalarını aşmadan yap; kısıt varsa kullanma. İmzalı sürümler, kaynak kod denetimi ve izinler gelecekte değerlendirilir.
- **Hata raporları:** Gerçek e-posta içeriklerini veya müşteri kimliklerini GitHub Issues'a koyma; yalnızca kişisel verilerden arındırılmış hatalar kullan.
- **Rapor riski:** Outlook'ta gönderilmeyen e-posta taslağı olarak açılan rapor, otomatik taslak kaydı nedeniyle kurumsal posta kutusunda tutulabilir; daha az kopyalama için yerel rapor görünümü seçeneği değerlendirilecek.

## V0.7'de gerçekten durum ne?

Mevcut kod Outlook VBA üzerinde çalışır ve doğrudan bir üçüncü taraf HTTP bağlantısı için tasarlanmamıştır; bir kurumsal güvenlik sertifikası veya eksiksiz denetim sunmuyoruz. Kaynak e-posta metninin 10.000 karakterine kadar olan bölümünü görev gövdesine, gelen yanıtın 4.500 karakterine kadar olan kısmını da aynı kayda kopyalayabilir. İşlerim raporu Outlook compose penceresinde açılır.

Bu nedenle **"Hiçbir veri bilgisayardan çıkmaz"** iddiası V0.7 için kullanılmamalıdır.

## Öncelikli geliştirmeler

1. V0.7'nin takip oluşturma / rapor / yanıt eşleştirme akışını farklı klasik Outlook kurulumlarında test et ve hata kayıtlarını düzelt.
2. Görev içindeki kopyalanmış e-posta içeriğini kaldırıp, kaynağa güvenli referansla gerektiğinde eriş.
3. Düşük maliyetli ve Outlook'u kilitlemeyen 45 günlük artımlı yanıt kontrolü.
4. Tek panelde sade iş durumları ve anlaşılır kullanıcı dili.
5. Kurumsal dağıtım için imza, izin ve kurulum belgelerini tamamla.

Bu belge hedef ürün kararlarıdır; henüz uygulanmamış özellikler çalışan V0.7 özelliği olarak sunulmaz.
