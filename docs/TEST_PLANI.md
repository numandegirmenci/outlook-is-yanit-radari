# v0.9.1 test planı

**Dikkat:** Testlerde kişisel/kurumsal gerçek e-posta kullanmayın. Önceki sürümün `.frm` dosyasını import etmeyin.

1. `Alt + F11` → `Insert → UserForm` ile `frmRadarIsMasam` oluştur; kodunu yapıştır.
2. `Debug → Compile VBAProject`: hata çıkmamalı. Hata çıkarsa dur; tam metni kaydet.
3. Test e-postası seç → `Takip_2_YapacagimIs` → `0` (hatırlatmasız).
4. `Radar_11_IsMasam`: tek pencere açılmalı; test görevi listede görünmeli.
5. Görevi seç: ilk e-postanın **konu + tarih + gönderen + alıcı + metni** görünmeli.
6. Ayrı test e-postasında `Takip_1_YanitBekliyorum` oluştur; **Yanıt Bekleyen** filtresinde görünmeli.
7. Aynı konuşmada test yanıtı varsa, Yenile/yeniden seçim sonrası yeni mesaj görünmeli (Outlook Conversation desteği ve erişime bağlı).
8. **Orijinal E-postayı Aç** kaynak mesajı göstermeli. Bulunamazsa Outlook araması için onay sorulmalı.
9. **Hatırlatma Taslağı** açıldığında otomatik gönderim olmamalı; alıcıyı ve metni elle kontrol et.
10. **Tamamlandı** düğmesini yalnızca test görevinde kullan; durum gerçekten DONE olmalı.
11. `Takip_3_Islerim` raporunda Yapacağım İşler/ Yanıt Bekleyenler birbirine karışmamalı.
12. Birden çok kez açma/kapama, 600'den fazla kayıt, Outlook klasörü taşıma, farklı Exchange hesapları gibi uç durumlar **henüz doğrulanmadı**.

Hata bildirirken yalnızca hata kodunu, işlem adını ve kişisel verilerden arındırılmış ekran görüntüsünü paylaşın.
