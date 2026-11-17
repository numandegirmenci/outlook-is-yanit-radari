# Kurulum ve güncelleme — v0.9.1 Genel Beta

Bu proje klasik Windows Outlook ve izin verilmiş VBA makrolarıyla çalışır. **Yeni Outlook** desteklenmez.

## İlk kurulum

1. GitHub üzerinden **Code → Download ZIP** indirip klasöre çıkartın.
2. Outlook'ta `Alt + F11` → **File → Import File**: `src/IsYanitRadari.bas` ve `src/RadarIsMasam.bas` yükleyin.
3. **Insert → UserForm** ile boş form oluşturun, **(Name)** özelliğini `frmRadarIsMasam` yapın.
4. Formun **kod penceresine** `src/frmRadarIsMasam_KODU.txt` içeriğinin tamamını yapıştırın. Form tasarımında manuel düğme eklemeyin.
5. **Debug → Compile VBAProject**, `Ctrl + S`, Outlook'u yeniden açın.
6. Test e-postası seçin → `Alt + F8` → `Takip_2_YapacagimIs` → tarih seçin → `Radar_11_IsMasam` → görevi tıklayın.

**Bozuk eski formu içeri aktarmayın:** `frmRadarIsMasam.frm`/`.frx` dosyaları dağıtılmıyor; form yalnızca `Insert → UserForm` ile eklenir.

## Mevcut v0.8.3 sürümünden güncelleme

1. `Alt + F11` → eski `RadarIsMasam` modülünü **Export File** ile yedekleyin, sonra **Remove** ile kaldırın.
2. `src/RadarIsMasam.bas` dosyasını import edin.
3. Mevcut `frmRadarIsMasam` formu varsa yeniden oluşturmayın; kodunu yenileyin. Yoksa yeni form oluşturun. [Ayrıntılı rehber](FORM_KURULUMU.md).
4. **`TakipKolay` ve `ThisOutlookSession` değiştirilmez.** `Görevler → Takip Kolay` klasörünü silmeyin.
5. Derleyin ve Outlook'u yeniden açın.

## Otomatik yanıt algılama (isteğe bağlı)

`src/ThisOutlookSession.txt` içindeki olay prosedürleri yalnızca uygun kurulumlarda ve çakışan `Application_Startup` / `Application_NewMailEx` prosedürleri birleştirilerek kullanılmalıdır. Emin değilseniz otomatik kurulum yapmayın; `Takip_4_YanitlariKontrolEt` ile manuel yanıt kontrolünü kullanın.

## Hızlı tarih seçimi

`1`: 1 saat sonra; `2`: yarın 09.00; `3`: 3 gün sonra; `4`: 1 hafta sonra; `5`: özel tarih/saat (`25.11 15:30`); `0`: tarih olmadan kaydet; boş/İptal: işlemi iptal et. Tarih otomatik atanmaz.

## Dikkat

- Macro güvenliğini kapatmayın; kurum onayı gerekir.
- İlk testlerde kurum evrakı veya kişisel bilgiler içermeyen deneme e-postaları kullanın.
- Görevler Exchange/Microsoft 365 ile eşitlenebilir. Harici üçüncü taraf sunucu kullanmamak, bütün bilgilerin yalnızca cihazda olduğu anlamına gelmez.
- Kod beta aşamasındadır; hata raporlarını şirket verilerinden arındırın.
