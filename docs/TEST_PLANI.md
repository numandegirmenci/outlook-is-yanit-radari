# v0.8.3 — Basit test planı

Önemli: İlk denemeleri gerçek iş yerine **test e-postalarıyla** yap.

1. `Alt + F11` → **Debug → Compile VBAProject**: derleme hatası olmamalı.
2. Test e-postası seç → `Takip_2_YapacagimIs` → `2`: yarın 09.00 için hatırlatma oluşturulmalı.
3. Farklı e-posta seç → `Takip_1_YanitBekliyorum` → `0`: hatırlatmasız görev açılmalı.
4. Bir test görevi seç → `Takip_6_TarihiDegistir` → `5` → gelecekteki `gg.aa 15:30`: tarih güncellenmeli.
5. Aynı görev → `Takip_6_TarihiDegistir` → boş/İptal: önceki tarih değişmemeli.
6. Geçersiz tarih (`31.02`) veya geçmiş tarih gir: işlem reddedilmeli.
7. `Radar_11_IsMasam`: Yapılacaklar `TODO`, Yanıt bekleyenler `WAIT` gibi gruplarda görünmeli.
8. Test görevi seç → `Radar_12_KaynakYazismayiAc`: kaynak mesaj bulunabiliyorsa açılmalı.
9. Deneme görevi seç → `Radar_13_KategoriDegistir`: durum değişikliği kaydedilmeli.
10. `Takip_3_Islerim` raporunu aç: yapılacak işler kaybolmamalı, hiçbir mail otomatik gönderilmemeli.
11. `Takip_10_DurumleriKontrolEt`: sayılar görevlerle uyumlu olmalı.
12. Yanıt takibi testi için deneme yazışmasına yanıt ver: `Takip_4_YanitlariKontrolEt` ile kontrol et; görev otomatik tamamlanmamalı.

**Beta notu:** Bu testlerin tamamı gerçek Outlook ortamında doğrulanmadan kararlı sürüm iddiası kullanma. İş Masam görünümü Outlook sürümüne bağlı farklılık gösterebilir.
