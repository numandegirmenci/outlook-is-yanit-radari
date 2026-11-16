# Gizlilik ve güvenlik — v0.8.3

Bu yazılım, klasik Outlook'un kendi VBA ortamında çalışmak üzere hazırlanmış açık kaynaklı bir araçtır. Bağımsız bir iş takip hesabı ya da geliştiricinin sunucusunu gerektirmez.

## Üçüncü taraflara aktarım

Kaynak kodda kullanıcı e-postalarını **Notion, Trello, reklam, analiz, telemetri veya geliştirici sunucularına göndermek** için bir bağlantı bulunmaz. Bu, bağımsız dış hizmete otomatik aktarım yapılmaması hedefidir; kurum genelinde onaylanmış güvenlik incelemesinin yerine geçmez.

## Outlook/Exchange içinde saklananlar

- Yeni takip kaydında durum, kaynak e-posta kimliği, kişi, konu, zaman bilgisi ve benzeri takip üstverileri Outlook **Görevler / Takip Kolay** klasörüne yazılır.
- **v0.8 ve sonrasında yeni görev oluşturulurken** e-posta gövdesi ayrıca görev içine kopyalanmaz. Gelen yanıtların gövdesi yerine, yakalanan yanıt kimliği saklanır ve metni gerektiğinde Outlook'tan okunur.
- **Eski sürümlerden kalan görevler** e-posta metni içeriyor olabilir; güncelleme bunları kendiliğinden temizlemez.
- Görevler, e-postalar ve Outlook tarafından kaydedilen taslaklar **Exchange/Microsoft 365 üzerinden şirketin var olan Microsoft altyapısıyla eşitlenebilir**. Buna "tamamen yerel, cihazdan hiç çıkmıyor" denemez.
- HTML rapor bir Outlook e-posta düzenleme penceresi olarak görüntülenir. Makro otomatik göndermez; kullanıcı yanlışlıkla gönderebilir veya Outlook taslağı senkronize edebilir.
- Kaynak e-posta taşınmış ya da silinmişse bağlantı açılamayabilir; bu durumda eski bir görev kaydı korunur.

## Kullanıcı kontrolü

Kullanıcı onayı olmadan mail gönderme, silme veya taşıma özelliği eklenmedi. Hatırlatma taslakları kullanıcıya gösterilir; alıcı ve metin kontrol edilmelidir. Kurumsal makro kısıtları aşılmamalıdır.

## Hata bildirimi

GitHub Issues üzerinde **gerçek e-posta, belge, şirket adı, müşteri bilgisi, kimlik, ekran görüntüsündeki gizli alanlar veya adres** paylaşma. Hata bildirirken örnek/anonimleştirilmiş veriler kullan.

**Durum:** Beta, kurumun güvenlik/onay sürecinden geçmiş bir ürün olduğu iddia edilmez.
