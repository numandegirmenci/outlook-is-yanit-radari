# Genel beta test planı

> **Doğrulama durumu:** V0.7'nin tüm işlevleri gerçek klasik Outlook üzerinde henüz uçtan uca doğrulanmadı. Kaynak paketinin bütünlüğü kontrol edilmiştir; bu, VBA derleme ve çalışma testinin yerine geçmez.

**Gerçek müşteri/şirket yazışmaları yerine örnek e-postalar kullanın.**

| Test | Beklenen sonuç | Durum |
|---|---|---|
| Klasik Outlook VBA derlemesi | `Debug → Compile` hatasız | Kullanıcı testi gerekli |
| E-postayı 'Yanıt bekliyorum' olarak işaretle | Görevler > Takip Kolay altında yeni kayıt | Kullanıcı testi gerekli (önceki sürümde başarılı örnek var) |
| `YOK` ile tarihsiz takip | Tarih ve hatırlatma oluşturulmamalı | Kullanıcı testi gerekli |
| `İşlerim` raporu | Gerçek açık görev sayısı ve içerik özeti | Kullanıcı testi gerekli |
| Farklı hesapla yazışmaya yanıt | Takip `İncelenecek` olmalı, otomatik kapanmamalı | Kullanıcı testi gerekli |
| Outlook kapalıyken gelen yanıt | Açılış taramasında bulunabilir, garanti değil | Kullanıcı testi gerekli |
| Hatırlatma taslağı | Gerçek konuşma içeriği görülmeli, otomatik gönderim olmamalı | Kullanıcı testi gerekli |
| 'Tamamlandı' | Onay istenmeli, görev tamamlanmalı | Kullanıcı testi gerekli |

## Hata raporu

- Outlook sürümü (klasik/yeni), Windows sürümü:
- Makro adı:
- Tam hata numarası ve metni:
- Hatadan hemen önce yapılan işlem:
- `Debug` düğmesiyle vurgulanan kod satırı (varsa):
- İlgili test e-postasının gizlilikten arındırılmış örnek konusu:

## Tarama limitleri

Elle kontrol: son 45 gün, en fazla 80 klasör, klasör başına 200 e-posta, yaklaşık 10 saniyelik bütçe.

Açılış kontrolü: son 45 gün, 30 klasör, klasör başına 80 e-posta, yaklaşık 3 saniyelik bütçe. **Bu kapsam eksiksiz tarama garantisi değildir.**


## v0.7.1 kontrolü

1. Tarihi geçmiş bir e-postayı **Yapacağım İş** olarak kaydedin; raporda **Yapılacak İşler** altında görünmeli.
2. Aynı e-postayı **Yanıt Bekliyorum** olarak kaydedin; iki ayrı görev oluşmalı.
3. `Takip_10_DurumleriKontrolEt` sayıları raporla karşılaştırın.
4. Yeni makroyu import etmeden eski modülü yedekleyin. `ThisOutlookSession` dosyası değişmedi.
5. Farklı Outlook kurulumlarında VBA compile ve canlı akış testleri tamamlanmadan v1.0 kararlı demeyin.
