# İş Masam formunun kurulumu (v0.9.1)

**Önemli:** `frmRadarIsMasam.frm` dosyasını içeri aktarmayın. Önceki testte bazı kullanıcılarda `System Error &H80004005` oluştu. Bu pakette bozuk/eksik form dosyası yoktur.

## İlk kurulum

1. **Klasik** Outlook'ta `Alt + F11` tuşlarına bas. `Project - VBAProject.OTM` altında **Modules** bölümünü gör.
2. **File → Import File** ile `src/IsYanitRadari.bas` ve `src/RadarIsMasam.bas` dosyalarını içeri al. Daha önce aynı adlı modüller varsa önce **Export File** ile yedekle; aynı isimden iki kopya bırakma.
3. VBE penceresinde `[break]` yazıyorsa **Run → Reset** işlemini yap.
4. Solda `VbaProject.OTM` seçiliyken **Insert → UserForm** seç. `UserForm1` adlı boş form oluşur.
5. Form seçiliyken `F4` ile **Properties** aç, `(Name)` satırını **`frmRadarIsMasam`** yap. `Caption` dilediğin gibi olabilir.
6. Form açıkken `F7` ile tasarım ekranından **kod penceresine** geç. `src/frmRadarIsMasam_KODU.txt` içindeki bütün metni kod penceresine yapıştır. `Option Explicit` satırı en başta olmalı.
7. Form tasarımına manuel olarak düğme, kutu veya etiket ekleme; form ilk açıldığında bunları VBA koduyla oluşturur.
8. **Debug → Compile VBAProject** çalıştır; hata yoksa `Ctrl + S` ile kaydet. Outlook'u kapatıp yeniden aç.
9. Bir test e-postası seç → `Alt + F8` → `Takip_2_YapacagimIs`. Sonra `Radar_11_IsMasam` ile paneli aç.

## Güncelleme yapıyorsan

- `TakipKolay`, `ThisOutlookSession` ve Outlook içindeki **Takip Kolay** görev klasörünü değiştirme.
- Eski `RadarIsMasam` modülünü önce **Export File** ile yedekleyip kaldır, ardından yeni `src/RadarIsMasam.bas` dosyasını import et.
- `Forms` altında **`frmRadarIsMasam`** zaten varsa **ikinci form oluşturma**; `F7` ile kodunu açıp eski kodun tamamını yeni sürümdeki metinle değiştir.
- Aynı adlı form görünmüyorsa yukarıdaki 4-7. adımları uygula.

## Hata olursa

- `System Error &H80004005` → Hatalı `.frm` dosyasını import etmeyi bırak; boş UserForm oluşturma yolunu kullan.
- `Variable not defined` veya `frmRadarIsMasam not defined` → Yeni `RadarIsMasam.bas` sürümü formu adıyla, çalışma anında yükler. Hem modülün güncel olduğundan hem de form adının birebir doğru yazıldığından emin ol.
- **Compile error** → Hata veren satırı ve uyarının tam metnini kontrol et. Yeni kodu rastgele başka modüllere yapıştırma.
- Panel açılıyor ama e-posta içeriği yok → Kaynak mesaj taşınmış, silinmiş veya başka posta deposunda olabilir. **Orijinal E-postayı Aç** seçeneğinde Outlook araması sorulabilir.

**Veri güvenliği:** Kurulum görevlere dokunmaz. Ancak formdaki **Tamamlandı** düğmesi kullanıldığında seçili görevin durumu gerçekten değiştirilir. Testleri deneme görevleriyle yap.
