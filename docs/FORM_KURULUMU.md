# İş Masam formu — V1.0 RC3 güncelleme

## Mevcut v0.9.1, RC1 veya RC2'den geçiş

1. Klasik Outlook'ta `Alt + F11` açın.
2. **Forms → frmRadarIsMasam** formunu seçin. **File → Export File** ile yedekleyin; varsa `.frx` dosyasını da beraber tutun.
3. `F7` ile formun **kod penceresine** geçin.
4. Eski kodun tamamını `src/frmRadarIsMasam_KODU.txt` dosyasındaki kodla değiştirin.
5. **Debug → Compile VBAProject** yapın; hata varsa devam etmeyin ve yedeğe dönün.
6. `Ctrl + S` ile kaydedin; Outlook'u tamamen kapatıp açın. `Alt + F8` → `Radar_11_IsMasam` komutunu çalıştırın.

**Önemli:** `TakipKolay`, `RadarIsMasam`, `ThisOutlookSession` ve kişisel ek makrolarınızın bulunduğu `Module1` veya benzer modülleri silmeyin. Görevler → Takip Kolay klasörüne dokunmayın.

## İlk kurulum

1. `src/IsYanitRadari.bas` ve `src/RadarIsMasam.bas` dosyalarını **File → Import File** ile yükleyin.
2. **Insert → UserForm** ile boş form ekleyin; `F4` ile adını `frmRadarIsMasam` yapın.
3. `F7` ile kod alanını açın; `src/frmRadarIsMasam_KODU.txt` içeriğini yapıştırın. Düğmeler otomatik oluşturulur.
4. Derleyin, kaydedin, Outlook'u yeniden başlatıp test e-postasında deneyin.

**Eski bozuk .frm dosyasını içe aktarmayın.** Bazı sistemlerde `&H80004005` hatası oluşuyordu.

## RC3 testleri

- **Tüm açık işler** listesi tamamlanmışları göstermemeli.
- **Tamamlanan** kartı ayrı pencere açmalı.
- Soldaki sütun başlıkları hizalı olmalı.
- **Yanıtla / Tümüne Yanıtla** kaynak e-postaya gerçek yanıt taslağı açmalı, otomatik göndermemeli.
- Göndermeden önce **Kime / Bilgi (CC)** alanlarını mutlaka kontrol edin.
