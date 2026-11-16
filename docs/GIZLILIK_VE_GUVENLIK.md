# Gizlilik ve güvenlik

Bu proje Classic Outlook VBA içinde çalışır. **Harici bir analiz servisine posta içeriği gönderen bağlantı kodu içermez.** Buna rağmen 'tamamen yerel/veri dışarı çıkmaz' demek yanıltıcı olur:

- Yeni takip oluşturulurken kaynak e-posta metninden **en fazla 10.000 karakter**, Outlook görev kaydının `Body` alanına kopyalanır.
- Yanıt bulunduğunda görev içine yanıt metninden de bölüm kaydedilebilir.
- Görev kayıtları varsayılan Outlook Görevler hesabındadır. Exchange/Microsoft 365 senkronizasyonuna açıksa bu veriler ilgili **kurumsal Microsoft sunucularına** senkronize olur.
- Rapor e-posta içeriğinden özetler gösterir ve Outlook'un **gönderilmeyen** compose penceresinde açılır. Yanlışlıkla kullanıcı tarafından gönderilmesine karşı raporda alıcı belirtmeyin.
- Hatırlatma e-postası bir **taslak** olarak açılır; makro kendisi `Send` çağrısı yapmaz. Taslağın alıcısını, geçmiş yazışmayı ve eki olup olmadığını kullanıcının kontrol etmesi gerekir.
- VBA projesini bilinmeyen kaynaklardan çalıştırmayın. Önce kodu inceleyin ve kaynağı doğrulayın; kuruluş politikasını aşmayın.

## Sorun bildirirken

Gerçek e-postaları, ekleri, müşteri/satıcı isimlerini, e-posta adreslerini, sözleşmeleri, fatura veya kurumsal belgeleri GitHub Issues üzerine **yüklemeyin**. Hata bilgisini kişisel verilerden arındırın.

Bu yazılım **pilot** niteliğindedir; hassas/kurumsal verilerde kullanımı kurum iznine tabidir.
