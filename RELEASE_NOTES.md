# v0.9.1 Genel Beta — İş Masam'da e-posta içerikleri

**Ne yeni?** Görevini İş Masam'da seçtiğinde **yalnızca başlığı değil, e-postada yazılanları** da aynı pencerede görebiliyorsun. Outlook'un bulabildiği yazışma yanıtları da listelenebilir. Kaynak e-postaya dönmek, hatırlatma taslağı açmak ve işi tamamlamak için ayrı makrolar çalıştırmak gerekmiyor.

**Kurulum değişti:** Önceki testte `&H80004005` hatası veren bozuk `.frm` dosyası kaldırıldı. İlk kurulumda VBA editöründe boş **UserForm** oluşturup `src/frmRadarIsMasam_KODU.txt` içeriğini yapıştırmak gerekir. [Adım adım kurulum](docs/FORM_KURULUMU.md).

**Gizlilik:** Harici iş takip servisi yoktur. Yazışma metni panelde Outlook'tan okunur; ancak kurumsal Outlook/Exchange senkronizasyonu ayrı devam edebilir. Taslaklar kendiliğinden gönderilmez.

**Sürüm durumu:** Genel Beta — farklı Outlook ortamlarında derleme ve canlı test henüz tamamlanmamıştır. Önce test e-postasıyla deneyin.
