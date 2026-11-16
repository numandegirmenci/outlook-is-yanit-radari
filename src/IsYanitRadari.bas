Attribute VB_Name = "TakipKolay"
Option Explicit

' OUTLOOK ÝÞ VE YANIT RADARI - Klasik Outlook, PILOT v0.7
' Yalnizca secilen mail uzerinden kullanici istegiyle gorev olusturur.
' Mail gondermez, silmez, tasimaz. Eski takipleri aktarmaya CALISMAZ.
' Kayitlar Outlook > Gorevler > Takip Kolay alt klasorunde saklanir.

Private Const KLASOR_ADI As String = "Takip Kolay"
Private Const ALAN_DURUM As String = "TKState"
Private Const ALAN_KAYNAK As String = "TKMailID"
Private Const ALAN_POSTA As String = "TKMailStore"
Private Const ALAN_SOHBET As String = "TKConvID"
Private Const ALAN_TARIH As String = "TKCreated"
Private Const ALAN_KISI As String = "TKPerson"
Private Const ALAN_KONU As String = "TKTopic"
Private Const BEKLIYOR As String = "WAIT"
Private Const YAPILACAK As String = "TODO"
Private Const KONTROL As String = "REVIEW"
Private Const BITTI As String = "DONE"
Private Const KAYIT_SINIRI As Long = 4000
Private Const TARAMA_KLASOR_SINIRI As Long = 80
Private Const KLASOR_BASINA_MAIL As Long = 200
Private Const TARAMA_GUN As Long = 45
Private Const ALAN_ALICI As String = "TKRecipient"
Private Const ICERIK_AYRAC As String = "---- KAYNAK E-POSTA METNI ----"
Private Const YANIT_AYRAC As String = "---- GELEN YANIT METNI ----"
Private Const BASLANGIC_GUN As Long = 45
Private Const BASLANGIC_KLASOR As Long = 30
Private Const BASLANGIC_MAIL As Long = 80

' 1) Gonderdigim ya da secili yazismadan yanit bekliyorum.
Public Sub Takip_1_YanitBekliyorum()
    YeniTakip BEKLIYOR
End Sub

' 2) Secili mail ile ilgili benim yapacagim bir is var.
Public Sub Takip_2_YapacagimIs()
    YeniTakip YAPILACAK
End Sub

' 3) Islerim: Kisa sayisal ozet + Outlook'un Gorevler listesini ac.
Public Sub Takip_3_Islerim()
    On Error GoTo RaporHatasi
    Dim hedef As Outlook.MAPIFolder, nesne As Object, t As Outlook.TaskItem
    Dim rapor As Outlook.MailItem, kod As String, konu As String, kisi As String
    Dim tarihEtiketi As String, ozet As String, satir As String, html As String
    Dim toplam As Long, acik As Long, tamam As Long, hatali As Long
    Dim yanit As Long, yapilacak As Long, incelenecek As Long
    Dim geciken As Long, bugun As Long, tarihsiz As Long, listelenen As Long
    Dim sGec As String, sKontrol As String, sBugun As String
    Dim sYanit As String, sYap As String, sDiger As String
    Dim asama As String, vadeVar As Boolean, isGec As Boolean, isBugun As Boolean

    ' Rapor, gorevleri sadece OKUR. Arka planda posta taramasi yapmaz.
    ' Bu sayede rapor acmak Outlook'u yavaslatmaz veya kayitlari degistirmez.
    asama = "Takip klasorune erisim"
    Set hedef = TakipKlasoru()
    asama = "Gorevleri okuma"

    For Each nesne In hedef.Items
        If toplam >= KAYIT_SINIRI Then Exit For
        If TypeOf nesne Is Outlook.TaskItem Then
            Set t = nesne
            toplam = toplam + 1
            On Error GoTo KayitHatasi
            kod = Durum(t)
            If kod = BITTI Then
                tamam = tamam + 1
            Else
                acik = acik + 1
                konu = GorunenKonu(t)
                kisi = AlanOku(t, ALAN_KISI)
                If Len(Trim$(kisi)) = 0 Then kisi = "Belirtilmedi"
                ozet = MetinOzeti(TakipMetni(t, 480), 380)

                vadeVar = TakipTarihiVar(t)
                isGec = False
                isBugun = False
                If vadeVar Then
                    tarihEtiketi = Format$(t.DueDate, "dd.mm.yyyy")
                    isGec = (DateValue(t.DueDate) < Date)
                    isBugun = (DateValue(t.DueDate) = Date)
                    If isGec Then geciken = geciken + 1
                    If isBugun Then bugun = bugun + 1
                Else
                    tarihEtiketi = "Tarih belirlenmedi"
                    tarihsiz = tarihsiz + 1
                End If
                Select Case kod
                    Case BEKLIYOR: yanit = yanit + 1
                    Case YAPILACAK: yapilacak = yapilacak + 1
                    Case KONTROL: incelenecek = incelenecek + 1
                End Select

                If listelenen < 300 Then
                    satir = RaporSatiri(konu, kisi, tarihEtiketi, kod, ozet)
                    If isGec Then
                        sGec = sGec & satir
                    ElseIf kod = KONTROL Then
                        sKontrol = sKontrol & satir
                    ElseIf isBugun Then
                        sBugun = sBugun & satir
                    ElseIf kod = BEKLIYOR Then
                        sYanit = sYanit & satir
                    ElseIf kod = YAPILACAK Then
                        sYap = sYap & satir
                    Else
                        sDiger = sDiger & satir
                    End If
                    listelenen = listelenen + 1
                End If
            End If
            GoTo KayitTamam
KayitHatasi:
            ' Tek bir uyumsuz gorev, tum raporu engellemesin.
            hatali = hatali + 1
            Err.Clear
            Resume KayitTamam
KayitTamam:
            On Error GoTo RaporHatasi
        End If
    Next nesne

    asama = "Rapor tasarimi"
    html = "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>" & _
           "<body style='margin:0;padding:0;background:#eef2f7;font-family:Calibri,Arial,sans-serif;color:#172a3a'>" & _
           "<table width='100%' cellpadding='0' cellspacing='0'><tr><td align='center' style='padding:24px'>" & _
           "<table width='820' cellpadding='0' cellspacing='0' style='max-width:820px;width:100%;background:#ffffff;border:1px solid #dfe6ef'>" & _
           "<tr><td style='padding:26px 30px;background:#12395b;color:white'>" & _
           "<div style='font-size:12px;letter-spacing:1px;color:#bcd9f0'>OUTLOOK ÝÞ VE YANIT RADARI</div>" & _
           "<div style='font-size:25px;font-weight:bold;padding-top:6px'>Islerim / Gunluk durum</div>" & _
           "<div style='font-size:12px;color:#dceaf4;margin-top:8px'>" & Format$(Now, "dd.mm.yyyy HH:nn") & "</div>" & _
           "</td></tr><tr><td style='padding:24px 28px'>"
    html = html & "<table width='100%' cellpadding='8' cellspacing='0' style='text-align:center;border:1px solid #dde5ef'><tr>" & _
           TakipHtmlKutu("Açýk iþler", acik, "#12395b") & _
           TakipHtmlKutu("Yanýt bekliyor", yanit, "#155cb3") & _
           TakipHtmlKutu("Yapýlacak", yapilacak, "#047857") & _
           TakipHtmlKutu("Ýncelenecek", incelenecek, "#b25a00") & "</tr></table>"
    html = html & "<p style='font-size:13px;color:#435366'>" & _
           "<b>Geciken:</b> " & geciken & " &nbsp; <b>Bugun:</b> " & bugun & _
           " &nbsp; <b>Tarihsiz:</b> " & tarihsiz & " &nbsp; <b>Tamamlanan:</b> " & tamam & "</p>"

    If acik = 0 Then
        html = html & "<p style='padding:16px;background:#ecfdf3;color:#24643f'>Acik takip bulunamadi.</p>"
    Else
        html = html & RaporBolumu("Gecikenler", sGec, "#b42318")
        html = html & RaporBolumu("Yanýt geldi - kontrol edilecek", sKontrol, "#ae6d00")
        html = html & RaporBolumu("Bugun kontrol edilecek", sBugun, "#285ba6")
        html = html & RaporBolumu("Yanýt bekleyenler", sYanit, "#285ba6")
        html = html & RaporBolumu("Yapýlacak iþler", sYap, "#117153")
        html = html & RaporBolumu("Diðer açýk iþler", sDiger, "#44556a")
    End If
    If hatali > 0 Then html = html & "<p style='color:#a33'>" & hatali & " kayit okunamadi; digerleri listelendi.</p>"
    If acik > listelenen Then html = html & "<p style='color:#a76a00'>Ilk 300 acik is gosteriliyor.</p>"
    html = html & "<div style='padding:15px 0 0;border-top:1px solid #e5e9ef;font-size:12px;color:#64748b'>" & _
           "Bu bir önizlemedir. Otomatik gönderilmez. Yazýþma içerikleri yalnýzca yerel Outlook'tan okunur. " & _
           "Yanýtlarý kontrol etmek için Takip_4_YanitlariKontrolEt komutunu kullanin." & _
           "</div></td></tr></table></td></tr></table></body></html>"

    asama = "E-posta onizlemesi"
    Set rapor = Application.CreateItem(olMailItem)
    rapor.Subject = "Outlook Ýþ ve Yanýt Radarý | " & Format$(Date, "dd.mm.yyyy")
    rapor.BodyFormat = olFormatHTML
    rapor.HTMLBody = html
    rapor.Display
    Exit Sub
RaporHatasi:
    MsgBox "Islerim raporu acilamadi." & vbCrLf & _
           "Adim: " & asama & vbCrLf & "Hata " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Outlook Ýþ ve Yanýt Radarý"
End Sub

Private Function RaporSatiri(ByVal konu As String, ByVal kisi As String, _
                            ByVal tarihEtiketi As String, ByVal kod As String, _
                            ByVal ozet As String) As String
    Dim ad As String
    Select Case kod
        Case BEKLIYOR: ad = "Yanýt bekleniyor"
        Case YAPILACAK: ad = "Yapýlacak"
        Case KONTROL: ad = "Yanýt geldi"
        Case Else: ad = "Açýk"
    End Select
    RaporSatiri = "<tr><td style='padding:11px;border-bottom:1px solid #edf0f3'>" & _
                 "<b>" & TakipHtmlKacir(Left$(konu, 200)) & "</b>" & _
                 "<div style='font-size:12px;color:#4c5d70;margin-top:6px;white-space:pre-wrap'>" & _
                 HtmlSatirli(ozet) & "</div></td>" & _
                 "<td style='padding:10px;vertical-align:top;border-bottom:1px solid #edf0f3'>" & TakipHtmlKacir(Left$(kisi, 85)) & "</td>" & _
                 "<td style='padding:10px;vertical-align:top;border-bottom:1px solid #edf0f3'>" & TakipHtmlKacir(tarihEtiketi) & "</td>" & _
                 "<td style='padding:10px;vertical-align:top;border-bottom:1px solid #edf0f3'>" & ad & "</td></tr>"
End Function

Private Function RaporBolumu(ByVal baslik As String, ByVal satirlar As String, _
                            ByVal vurgu As String) As String
    If Len(satirlar) = 0 Then Exit Function
    RaporBolumu = "<h2 style='font-size:17px;color:" & vurgu & ";margin:25px 0 10px'>" & baslik & "</h2>" & _
         "<table width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;font-size:12px'>" & _
         "<tr style='background:#f4f6fa;color:#53637a'>" & _
         "<th align='left' width='52%' style='padding:9px'>Konu ve yazisma ozeti</th>" & _
         "<th align='left' width='22%' style='padding:9px'>Kisi</th>" & _
         "<th align='left' width='14%' style='padding:9px'>Tarih</th>" & _
         "<th align='left' width='12%' style='padding:9px'>Durum</th></tr>" & _
         satirlar & "</table>"
End Function

Private Function HtmlSatirli(ByVal metin As String) As String
    metin = TakipHtmlKacir(metin)
    metin = Replace(metin, vbCrLf, "<br>")
    metin = Replace(metin, vbCr, "<br>")
    HtmlSatirli = Replace(metin, vbLf, "<br>")
End Function

Private Function MetinOzeti(ByVal metin As String, ByVal uzunluk As Long) As String
    metin = Replace(metin, vbTab, " ")
    metin = Replace(metin, Chr$(0), " ")
    metin = Trim$(metin)
    If Len(metin) > uzunluk Then metin = Left$(metin, uzunluk) & " ..."
    If Len(metin) = 0 Then metin = "Yazisma icerigi bulunamadi."
    MetinOzeti = metin
End Function


Private Function TakipHtmlKutu(ByVal baslik As String, ByVal sayi As Long, _
                               ByVal renk As String) As String
    TakipHtmlKutu = "<td width='25%' style='padding:12px;border-right:1px solid #e4e9ef;background:#f9fbfd'>" & _
       "<div style='font-size:26px;color:" & renk & ";font-weight:bold'>" & CStr(sayi) & "</div>" & _
       "<div style='font-size:11px;color:#53657b'>" & TakipHtmlKacir(baslik) & "</div></td>"
End Function

Private Function TakipHtmlKacir(ByVal metin As String) As String
    metin = Replace(metin, "&", "&amp;")
    metin = Replace(metin, "<", "&lt;")
    metin = Replace(metin, ">", "&gt;")
    metin = Replace(metin, """", "&quot;")
    TakipHtmlKacir = Replace(metin, "'", "&#39;")
End Function

' 4) Manuel kontrol: 45 gun, hiz ve klasor limitleri korunuyor.
Public Sub Takip_4_YanitlariKontrolEt()
    On Error GoTo Hata
    Dim taranan As Long, bulunan As Long, klasorSayisi As Long
    Dim beklenen As Object
    Set beklenen = BeklenenlerSozlugu()
    If beklenen.Count = 0 Then
        MsgBox "Yanýt bekleyen bir iþ bulunamadý.", vbInformation, "Outlook Ýþ ve Yanýt Radarý"
        Exit Sub
    End If
    YanitiTara Application.Session.GetDefaultFolder(olFolderInbox), _
          0, klasorSayisi, taranan, bulunan, beklenen, _
          TARAMA_GUN, TARAMA_KLASOR_SINIRI, KLASOR_BASINA_MAIL, 10
    MsgBox "Yeni yanýt bulunan iþ: " & bulunan & vbCrLf & _
           "Kontrol edilen e-posta: " & taranan & vbCrLf & _
           "Taranan klasör: " & klasorSayisi & vbCrLf & vbCrLf & _
           "Yanýt bulunduysa iþ 'Ýncelenecek' olur; otomatik tamamlanmaz." & vbCrLf & _
           "Kontrol varsayýlan Gelen Kutusu ve alt klasörleri ile sýnýrlýdýr.", _
           vbInformation, "Outlook Ýþ ve Yanýt Radarý"
    Exit Sub
Hata:
    HataGoster "Yanýt kontrolü"
End Sub

' Outlook acilisinda ve Islerim ekranina girerken kisa tarama.
' Bu bir tam arsiv taramasi degildir, Outlook'u gereksiz yavaslatmaz.
Public Sub Takip_BaslangicTaramasi()
    On Error GoTo Hata
    Dim beklenen As Object
    Dim klasorSayisi As Long, taranan As Long, bulunan As Long
    Set beklenen = BeklenenlerSozlugu()
    If beklenen.Count = 0 Then Exit Sub
    YanitiTara Application.Session.GetDefaultFolder(olFolderInbox), _
          0, klasorSayisi, taranan, bulunan, beklenen, _
          BASLANGIC_GUN, BASLANGIC_KLASOR, BASLANGIC_MAIL, 3
    Exit Sub
Hata:
    Debug.Print "Takip Kolay baslangic taramasi: " & Err.Number & " - " & Err.Description
End Sub

' ThisOutlookSession olaylari bu fonksiyonu cagirir. Bir yaniti iki kez isleme koymaz;
' cunku bir kez bulunan gorev WAIT -> REVIEW olur ve sonraki turda sozlukten cikar.
Public Function Takip_OtomatikYanit(ByVal GelenNesne As Object) As Boolean
    On Error GoTo Hata
    Dim beklenen As Object
    If GelenNesne Is Nothing Then Exit Function
    If Not TypeOf GelenNesne Is Outlook.MailItem Then Exit Function
    Set beklenen = BeklenenlerSozlugu()
    If beklenen.Count = 0 Then Exit Function
    Takip_OtomatikYanit = YanitEslesir(GelenNesne, beklenen)
    If Takip_OtomatikYanit Then
        MsgBox "Takip ettiðiniz bir yazýþmaya yanýt geldi." & vbCrLf & vbCrLf & _
               Left$(GelenNesne.Subject, 150) & vbCrLf & vbCrLf & _
               "Ýþ 'Ýncelenecek' durumuna alýndý; tamamlanmýþ sayýlmadý.", _
               vbInformation, "Outlook Ýþ ve Yanýt Radarý - Yanýt geldi"
    End If
    Exit Function
Hata:
    Debug.Print "Takip Kolay gelen yanit: " & Err.Number & " - " & Err.Description
End Function

' 5) Secili e-postanin veya Takip Kolay gorevinin takibini bitirir.
Public Sub Takip_5_Tamamlandi()
    On Error GoTo Hata
    Dim nesne As Object, gorev As Outlook.TaskItem
    Set nesne = SecilenNesne()
    If nesne Is Nothing Then
        MsgBox "Bitirmek istediðin e-postayý veya takip görevini seç.", vbInformation, "Outlook Ýþ ve Yanýt Radarý"
        Exit Sub
    End If
    If TypeOf nesne Is Outlook.TaskItem Then
        Set gorev = nesne
        If Not BizimGorevimiz(gorev) Then Set gorev = Nothing
    ElseIf TypeOf nesne Is Outlook.MailItem Then
        Set gorev = MailTakibi(nesne)
    End If
    If gorev Is Nothing Then
        MsgBox "Bu e-posta veya görev için açýk takip bulunamadý.", vbInformation, "Outlook Ýþ ve Yanýt Radarý"
        Exit Sub
    End If
    If Durum(gorev) = BITTI Then
        MsgBox "Bu iþ zaten tamamlanmýþ.", vbInformation, "Outlook Ýþ ve Yanýt Radarý"
        Exit Sub
    End If
    If MsgBox("Bu iþ gerçekten tamamlandý mý?" & vbCrLf & vbCrLf & _
              GorunenKonu(gorev), vbYesNo + vbQuestion + vbDefaultButton2, "Outlook Ýþ ve Yanýt Radarý") <> vbYes Then Exit Sub
    AlanYaz gorev, ALAN_DURUM, BITTI
    gorev.Status = olTaskComplete
    gorev.ReminderSet = False
    gorev.Save
    MsgBox "Ýþ tamamlandý olarak kaydedildi.", vbInformation, "Outlook Ýþ ve Yanýt Radarý"
    Exit Sub
Hata:
    HataGoster "Tamamlama"
End Sub

' 6) Secilen takibin tarihini degistirme.
Public Sub Takip_6_TarihiDegistir()
    On Error GoTo Hata
    Dim nesne As Object, gorev As Outlook.TaskItem
    Dim varMi As Boolean, bitis As Date, uyari As Date
    Set nesne = SecilenNesne()
    If nesne Is Nothing Then
        MsgBox "Degistirmek istediginiz e-postayi veya takip gorevini secin.", vbInformation
        Exit Sub
    End If
    If TypeOf nesne Is Outlook.TaskItem Then
        Set gorev = nesne
        If Not BizimGorevimiz(gorev) Then Set gorev = Nothing
    ElseIf TypeOf nesne Is Outlook.MailItem Then
        Set gorev = MailTakibi(nesne)
    End If
    If gorev Is Nothing Then
        MsgBox "Bu secime ait bir takip bulunamadi.", vbInformation
        Exit Sub
    End If
    If Durum(gorev) = BITTI Then Exit Sub
    If Not TarihSor(varMi, bitis, uyari) Then Exit Sub
    Uyarla gorev, varMi, bitis, uyari
    gorev.Save
    MsgBox "Tercihiniz kaydedildi: " & TakipTarihiMetni(gorev), vbInformation, "Outlook Ýþ ve Yanýt Radarý"
    Exit Sub
Hata:
    HataGoster "Tarih degistirme"
End Sub

' Kullanici karar vermeden hatirlatma veya vade olusturulmaz.
' "YOK" secenegi, hatirlatmasiz ve tarihsiz takip olusturur.
Private Function TarihSor(ByRef varMi As Boolean, ByRef bitis As Date, _
                          ByRef uyari As Date) As Boolean
    Dim veri As String, bolum() As String, tarih() As String, saat() As String
    Dim gg As Long, aa As Long, yyyy As Long, ss As Long, dd As Long
    Dim kontrol As Date
    veri = Trim$(InputBox("Hatirlatma tarihi ve saatini siz belirleyin." & vbCrLf & _
         "Ornek: 24.11.2026 14:30" & vbCrLf & _
         "Tarih olmadan takip icin YOK yazin." & vbCrLf & _
         "Vazgecmek icin Iptal veya bos birakin.", _
         "Takip / Hatirlatma zamani", ""))
    If Len(veri) = 0 Then Exit Function
    If UCase$(veri) = "YOK" Then
        varMi = False
        TarihSor = True
        Exit Function
    End If

    On Error GoTo Hata
    veri = Replace(veri, "/", ".")
    bolum = Split(veri, " ")
    If UBound(bolum) <> 1 Then GoTo Hata
    tarih = Split(bolum(0), ".")
    saat = Split(bolum(1), ":")
    If UBound(tarih) <> 2 Or UBound(saat) <> 1 Then GoTo Hata
    If Not SayiMi(tarih(0)) Or Not SayiMi(tarih(1)) Or Not SayiMi(tarih(2)) Then GoTo Hata
    If Not SayiMi(saat(0)) Or Not SayiMi(saat(1)) Then GoTo Hata
    gg = CLng(tarih(0)): aa = CLng(tarih(1)): yyyy = CLng(tarih(2))
    ss = CLng(saat(0)): dd = CLng(saat(1))
    If yyyy < 2020 Or yyyy > 2100 Or aa < 1 Or aa > 12 Then GoTo Hata
    If gg < 1 Or gg > 31 Or ss < 0 Or ss > 23 Or dd < 0 Or dd > 59 Then GoTo Hata
    kontrol = DateSerial(yyyy, aa, gg)
    If Year(kontrol) <> yyyy Or Month(kontrol) <> aa Or Day(kontrol) <> gg Then GoTo Hata
    uyari = kontrol + TimeSerial(ss, dd, 0)
    If uyari <= Now Then
        MsgBox "Gecmis tarih veya saat secilemez. Islem iptal edildi.", vbExclamation
        Exit Function
    End If
    bitis = kontrol
    varMi = True
    TarihSor = True
    Exit Function
Hata:
    MsgBox "Tarih anlasilamadi. Ornek: 24.11.2026 14:30" & vbCrLf & _
           "Kayit yapilmadi. Lutfen komutu yeniden calistirin.", vbExclamation, "Outlook Ýþ ve Yanýt Radarý"
End Function

Private Function SayiMi(ByVal x As String) As Boolean
    Dim i As Long, ch As Integer
    If Len(x) = 0 Then Exit Function
    For i = 1 To Len(x)
        ch = Asc(Mid$(x, i, 1))
        If ch < 48 Or ch > 57 Then Exit Function
    Next i
    SayiMi = True
End Function

Private Sub Uyarla(ByVal t As Outlook.TaskItem, ByVal varMi As Boolean, _
                   ByVal bitis As Date, ByVal uyari As Date)
    t.ReminderSet = False
    If varMi Then
        t.StartDate = Date
        t.DueDate = bitis
        t.ReminderTime = uyari
        t.ReminderSet = True
    Else
        ' Outlook'un tarihsiz gorev degeri.
        t.DueDate = DateSerial(4501, 1, 1)
        t.StartDate = DateSerial(4501, 1, 1)
    End If
End Sub

Private Function TakipTarihiVar(ByVal t As Outlook.TaskItem) As Boolean
    On Error GoTo Bitti
    If Year(t.DueDate) >= 2020 And Year(t.DueDate) <= 2100 Then
        TakipTarihiVar = True
    End If
Bitti:
End Function

Private Function TakipTarihiMetni(ByVal t As Outlook.TaskItem) As String
    If TakipTarihiVar(t) Then
        TakipTarihiMetni = Format$(t.DueDate, "dd.mm.yyyy")
        If t.ReminderSet Then TakipTarihiMetni = TakipTarihiMetni & " " & Format$(t.ReminderTime, "HH:nn")
    Else
        TakipTarihiMetni = "Tarih belirtilmedi / hatirlatma yok"
    End If
End Function


' Sorun giderme: takip sayilari neden sifir gorunuyor?
Public Sub Takip_7_SistemKontrolu()
    On Error GoTo Hata
    Dim klasor As Outlook.MAPIFolder
    Dim nesne As Object, gorev As Outlook.TaskItem
    Dim toplam As Long, acik As Long, tamam As Long
    Set klasor = TakipKlasoru()
    For Each nesne In klasor.Items
        If TypeOf nesne Is Outlook.TaskItem Then
            Set gorev = nesne
            toplam = toplam + 1
            If Durum(gorev) = BITTI Then
                tamam = tamam + 1
            Else
                acik = acik + 1
            End If
        End If
        If toplam >= KAYIT_SINIRI Then Exit For
    Next nesne
    MsgBox "Takip Kolay klasörü: " & vbCrLf & klasor.FolderPath & vbCrLf & vbCrLf & _
           "Toplam: " & toplam & vbCrLf & _
           "Açýk: " & acik & vbCrLf & _
           "Tamamlanan: " & tamam & vbCrLf & vbCrLf & _
           "Bu klasördeki kayýtlar 'Ýþlerim' ekranýnda sayýlýr.", _
           vbInformation, "Outlook Ýþ ve Yanýt Radarý - Sistem Kontrolü"
    Exit Sub
Hata:
    HataGoster "Sistem kontrolü"
End Sub

Private Function BizimGorevimiz(ByVal t As Outlook.TaskItem) As Boolean
    On Error GoTo Bitti
    If t Is Nothing Then Exit Function
    Dim kendi As Outlook.MAPIFolder, ana As Outlook.MAPIFolder
    Set kendi = t.Parent
    Set ana = TakipKlasoru()
    BizimGorevimiz = (StrComp(kendi.EntryID, ana.EntryID, vbBinaryCompare) = 0 And _
                      StrComp(kendi.StoreID, ana.StoreID, vbBinaryCompare) = 0)
Bitti:
End Function

' --- Islemlerin ici ---
Private Sub YeniTakip(ByVal isKodu As String)
    On Error GoTo Hata
    Dim m As Outlook.MailItem, gorev As Outlook.TaskItem
    Dim varMi As Boolean, bitis As Date, uyari As Date
    Dim kisi As String, kayitKlasoru As Outlook.MAPIFolder
    Dim metin As String, mesaj As String
    Set m = SecilenMail()
    If m Is Nothing Then
        MsgBox "Once bir e-posta secin, sonra takip komutunu calistirin.", vbInformation, "Outlook Ýþ ve Yanýt Radarý"
        Exit Sub
    End If
    If Len(m.EntryID) = 0 Then
        MsgBox "Bu e-posta henuz kaydedilmemis. Once kaydedin.", vbExclamation
        Exit Sub
    End If
    Set gorev = MailTakibi(m)
    If Not gorev Is Nothing Then
        MsgBox "Bu yazisma zaten takipte.", vbInformation, "Outlook Ýþ ve Yanýt Radarý"
        gorev.Display
        Exit Sub
    End If
    ' En onemli degisiklik: hatirlatma tarihini kullanici belirler.
    If Not TarihSor(varMi, bitis, uyari) Then Exit Sub

    kisi = TakipKisi(m, isKodu)
    Set kayitKlasoru = TakipKlasoru()
    Set gorev = kayitKlasoru.Items.Add(olTaskItem)
    gorev.Subject = "[Takip] " & Left$(m.Subject, 220)
    metin = Left$(m.Body, 10000)
    gorev.Body = "Kaynak e-posta konusu: " & m.Subject & vbCrLf & _
                 "Ilgili kisi: " & kisi & vbCrLf & _
                 ICERIK_AYRAC & vbCrLf & metin
    gorev.Status = olTaskNotStarted
    Uyarla gorev, varMi, bitis, uyari
    AlanYaz gorev, ALAN_DURUM, isKodu
    AlanYaz gorev, ALAN_KAYNAK, m.EntryID
    AlanYaz gorev, ALAN_POSTA, m.Parent.StoreID
    AlanYaz gorev, ALAN_SOHBET, m.ConversationID
    AlanYaz gorev, ALAN_TARIH, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    AlanYaz gorev, ALAN_KISI, kisi
    AlanYaz gorev, ALAN_KONU, m.Subject
    AlanYaz gorev, ALAN_ALICI, HatirlatmaAlicisi(m)
    gorev.Save
    If isKodu = BEKLIYOR Then
        mesaj = "Yanit beklenen yazisma takibe alindi."
    Else
        mesaj = "Yapilacak is kaydedildi."
    End If
    MsgBox mesaj & vbCrLf & "Hatirlatma: " & TakipTarihiMetni(gorev), _
           vbInformation, "Outlook Ýþ ve Yanýt Radarý"
    Exit Sub
Hata:
    HataGoster "Takibe alma"
End Sub


Private Function TakipKlasoru() As Outlook.MAPIFolder
    Dim kok As Outlook.MAPIFolder, alt As Outlook.MAPIFolder
    Set kok = Application.Session.GetDefaultFolder(olFolderTasks)
    For Each alt In kok.Folders
        If StrComp(alt.Name, KLASOR_ADI, vbTextCompare) = 0 Then
            Set TakipKlasoru = alt
            Exit Function
        End If
    Next alt
    Set TakipKlasoru = kok.Folders.Add(KLASOR_ADI, olFolderTasks)
End Function

Private Function SecilenNesne() As Object
    On Error GoTo Bitti
    Dim w As Object
    Set w = Application.ActiveWindow
    If TypeOf w Is Outlook.Explorer Then
        If w.Selection.Count > 0 Then Set SecilenNesne = w.Selection.Item(1)
    ElseIf TypeOf w Is Outlook.Inspector Then
        Set SecilenNesne = w.CurrentItem
    End If
Bitti:
End Function

Private Function SecilenMail() As Outlook.MailItem
    Dim nesne As Object
    Set nesne = SecilenNesne()
    If Not nesne Is Nothing Then
        If TypeOf nesne Is Outlook.MailItem Then Set SecilenMail = nesne
    End If
End Function

Private Function MailTakibi(ByVal m As Outlook.MailItem) As Outlook.TaskItem
    On Error GoTo Bitti
    Dim nesne As Object, t As Outlook.TaskItem, n As Long
    Dim id As String, sohbet As String
    id = m.EntryID
    sohbet = m.ConversationID
    For Each nesne In TakipKlasoru().Items
        n = n + 1
        If n > KAYIT_SINIRI Then Exit For
        If TypeOf nesne Is Outlook.TaskItem Then
            Set t = nesne
            If Durum(t) <> BITTI Then
                If Len(id) > 0 And AlanOku(t, ALAN_KAYNAK) = id Then
                    Set MailTakibi = t
                    Exit Function
                End If
                If Len(sohbet) > 0 And AlanOku(t, ALAN_SOHBET) = sohbet Then
                    Set MailTakibi = t
                    Exit Function
                End If
            End If
        End If
    Next nesne
Bitti:
End Function

Private Sub AlanYaz(ByVal t As Outlook.TaskItem, ByVal ad As String, ByVal icerik As String)
    Dim p As Outlook.UserProperty
    Set p = t.UserProperties.Find(ad)
    If p Is Nothing Then Set p = t.UserProperties.Add(ad, olText, False)
    p.Value = icerik
End Sub

Private Function AlanOku(ByVal t As Outlook.TaskItem, ByVal ad As String) As String
    On Error GoTo Bitti
    Dim p As Outlook.UserProperty
    Set p = t.UserProperties.Find(ad)
    If Not p Is Nothing Then AlanOku = CStr(p.Value)
Bitti:
End Function

Private Function Durum(ByVal t As Outlook.TaskItem) As String
    If t.Status = olTaskComplete Then
        Durum = BITTI
    Else
        Durum = AlanOku(t, ALAN_DURUM)
        If Len(Durum) = 0 Then Durum = YAPILACAK
    End If
End Function

Private Function KayitZamani(ByVal t As Outlook.TaskItem) As Date
    On Error GoTo Bitti
    Dim s As String
    s = AlanOku(t, ALAN_TARIH)
    If Len(s) >= 19 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        KayitZamani = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2))) + _
                      TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), CInt(Mid$(s, 18, 2)))
    ElseIf IsDate(s) Then
        KayitZamani = CDate(s)
    End If
Bitti:
End Function

Private Function TakipMetni(ByVal t As Outlook.TaskItem, ByVal limit As Long) As String
    On Error GoTo Bitti
    Dim m As Outlook.MailItem, icerik As String, yer As Long
    ' Rapor icin yanit gelmisse son mesajin metni one ciksin.
    If limit < 8000 Then
        yer = InStr(1, t.Body, YANIT_AYRAC, vbTextCompare)
        If yer > 0 Then
            icerik = "Gelen yanit: " & Mid$(t.Body, yer + Len(YANIT_AYRAC))
        End If
    End If
    ' Yeni gorevlerde kayitli orijinal metin Outlook'u tekrar tekrar acmaz.
    If Len(Trim$(icerik)) = 0 Then
        yer = InStr(1, t.Body, ICERIK_AYRAC, vbTextCompare)
        If yer > 0 Then icerik = Mid$(t.Body, yer + Len(ICERIK_AYRAC))
    End If
    ' Eski gorevlerde veya kapsamli taslakta orijinal e-postaya bak.
    If Len(Trim$(icerik)) = 0 Or limit >= 8000 Then
        Set m = KaynakMail(t)
        If Not m Is Nothing Then
            If Len(Trim$(m.Body)) > 0 Then icerik = m.Body
        End If
    End If
    If Len(Trim$(icerik)) = 0 Then
        ' v0.6 ile olusturulan eski gorevlerde orijinal icerik saklanmamis olabilir.
        icerik = "Kaynak yazisma metni bulunamadi."
    End If
    If Len(icerik) > limit Then icerik = Left$(icerik, limit) & vbCrLf & "[Devami kisaltildi]"
    TakipMetni = icerik
Bitti:
End Function


Private Function GorunenKonu(ByVal t As Outlook.TaskItem) As String
    GorunenKonu = AlanOku(t, ALAN_KONU)
    If Len(GorunenKonu) = 0 Then GorunenKonu = t.Subject
End Function

Private Function BeklenenlerSozlugu() As Object
    Dim d As Object, nesne As Object, gorev As Outlook.TaskItem
    Dim sohbet As String, say As Long
    Set d = CreateObject("Scripting.Dictionary")
    For Each nesne In TakipKlasoru().Items
        say = say + 1
        If say > KAYIT_SINIRI Then Exit For
        If TypeOf nesne Is Outlook.TaskItem Then
            Set gorev = nesne
            If Durum(gorev) = BEKLIYOR Then
                sohbet = AlanOku(gorev, ALAN_SOHBET)
                If Len(sohbet) > 0 Then
                    If Not d.Exists(sohbet) Then d.Add sohbet, gorev
                End If
            End If
        End If
    Next nesne
    Set BeklenenlerSozlugu = d
End Function

Private Sub YanitiTara(ByVal klasor As Outlook.MAPIFolder, ByVal derinlik As Long, _
                       ByRef klasorSayisi As Long, ByRef taranan As Long, _
                       ByRef bulunan As Long, ByVal beklenen As Object, _
                       Optional ByVal gunLimit As Long = 45, _
                       Optional ByVal klasorLimit As Long = 80, _
                       Optional ByVal mailLimit As Long = 200, _
                       Optional ByVal saniyeLimit As Long = 10, _
                       Optional ByVal baslama As Date = 0)
    On Error GoTo Devam
    Dim liste As Outlook.Items, aday As Outlook.Items
    Dim nesne As Object, alt As Outlook.MAPIFolder, m As Outlook.MailItem
    Dim adet As Long, filtre As String
    If baslama = 0 Then baslama = Now
    If derinlik > 5 Or klasorSayisi >= klasorLimit Or beklenen.Count = 0 Then Exit Sub
    If DateDiff("s", baslama, Now) >= saniyeLimit Then Exit Sub
    klasorSayisi = klasorSayisi + 1
    If klasor.DefaultItemType = olMailItem Then
        Set liste = klasor.Items
        ' 45 gunluk adaylari Outlook filtreler; binlerce eski maili tek tek acmayiz.
        filtre = "[ReceivedTime] >= '" & Format$(DateAdd("d", -gunLimit, Now), "ddddd h:nn AMPM") & "'"
        On Error Resume Next
        Set aday = liste.Restrict(filtre)
        If Err.Number <> 0 Then
            Err.Clear
            Set aday = Nothing
        End If
        On Error GoTo Devam
        If Not aday Is Nothing Then
            On Error Resume Next
            aday.Sort "[ReceivedTime]", True
            Err.Clear
            On Error GoTo Devam
            For Each nesne In aday
                adet = adet + 1
                If adet > mailLimit Or beklenen.Count = 0 Then Exit For
                If DateDiff("s", baslama, Now) >= saniyeLimit Then Exit For
                If TypeOf nesne Is Outlook.MailItem Then
                    Set m = nesne
                    taranan = taranan + 1
                    If YanitEslesir(m, beklenen) Then bulunan = bulunan + 1
                End If
                If (adet Mod 25) = 0 Then DoEvents
            Next nesne
        End If
    End If
    For Each alt In klasor.Folders
        If klasorSayisi >= klasorLimit Or beklenen.Count = 0 Then Exit For
        If DateDiff("s", baslama, Now) >= saniyeLimit Then Exit For
        YanitiTara alt, derinlik + 1, klasorSayisi, taranan, bulunan, beklenen, _
                   gunLimit, klasorLimit, mailLimit, saniyeLimit, baslama
    Next alt
Devam:
    ' Problemli alt klasorler taramayi butunlukle durdurmaz.
End Sub


Private Function YanitEslesir(ByVal mail As Outlook.MailItem, ByVal beklenen As Object) As Boolean
    On Error GoTo Bitti
    Dim sohbet As String, gorev As Outlook.TaskItem
    Dim baslangic As Date
    sohbet = mail.ConversationID
    If Len(sohbet) = 0 Then Exit Function
    If Not beklenen.Exists(sohbet) Then Exit Function
    Set gorev = beklenen.Item(sohbet)
    baslangic = 0
    baslangic = KayitZamani(gorev)
    If mail.ReceivedTime <= baslangic Then Exit Function
    ' Kontrol et: Gonderenin kendi hesabinin e-postasi olmamali.
    If BenimMesajim(mail) Then Exit Function
    AlanYaz gorev, ALAN_DURUM, KONTROL
    gorev.Status = olTaskInProgress
    ' Gelen yanit, kullanicinin secmis oldugu tarihi asla degistirmez.
    gorev.Body = Left$(gorev.Body, 12000) & vbCrLf & vbCrLf & YANIT_AYRAC & vbCrLf & _
                 Left$(mail.Body, 4500)
    gorev.Save
    beklenen.Remove sohbet
    YanitEslesir = True
Bitti:
End Function

Private Function TakipKisi(ByVal mail As Outlook.MailItem, ByVal durum As String) As String
    On Error GoTo Bitti
    If durum = BEKLIYOR And BenimMesajim(mail) Then
        TakipKisi = mail.To
    Else
        TakipKisi = mail.SenderName
    End If
    If Len(Trim$(TakipKisi)) = 0 Then TakipKisi = "Belirtilmemiþ"
    Exit Function
Bitti:
    TakipKisi = "Belirtilmemiþ"
End Function

Private Function BenimMesajim(ByVal mail As Outlook.MailItem) As Boolean
    On Error GoTo Bitti
    Dim hesap As Outlook.Account
    Dim gonderici As String, exKisi As Outlook.ExchangeUser
    gonderici = mail.SenderEmailAddress
    If mail.SenderEmailType = "EX" Then
        Set exKisi = mail.Sender.GetExchangeUser
        If Not exKisi Is Nothing Then gonderici = exKisi.PrimarySmtpAddress
    End If
    For Each hesap In Application.Session.Accounts
        If StrComp(gonderici, hesap.SmtpAddress, vbTextCompare) = 0 Then
            BenimMesajim = True
            Exit Function
        End If
    Next hesap
Bitti:
End Function

Private Sub HataGoster(ByVal bolum As String)
    MsgBox bolum & " sýrasýnda iþlem tamamlanamadý." & vbCrLf & _
           "Hata " & Err.Number & ": " & Err.Description, vbExclamation, "Outlook Ýþ ve Yanýt Radarý"
End Sub


' 8) Tek tusla HATIRLATMA TASLAGI olusturur. HICBIR ZAMAN SEND kullanmaz.
' Eski V0.5 gorevleri de destekler: kaynak e-posta erisilebiliyorsa alici bulur.
Public Sub Takip_8_HatirlatmaTaslagi()
    On Error GoTo Hata
    Dim secim As Object, gorev As Outlook.TaskItem
    Dim kaynak As Outlook.MailItem, taslak As Outlook.MailItem
    Dim alici As String, konu As String, metin As String, govde As String

    Set secim = SecilenNesne()
    If secim Is Nothing Then
        MsgBox "Once takibe alinmis e-postayi veya gorevi secin.", vbInformation
        Exit Sub
    End If
    If TypeOf secim Is Outlook.TaskItem Then
        Set gorev = secim
        If Not BizimGorevimiz(gorev) Then Set gorev = Nothing
    ElseIf TypeOf secim Is Outlook.MailItem Then
        Set gorev = MailTakibi(secim)
    End If
    If gorev Is Nothing Then
        MsgBox "Secili e-postaya ait takip bulunamadi.", vbInformation
        Exit Sub
    End If
    If Durum(gorev) <> BEKLIYOR Then
        MsgBox "Hatirlatma taslagi yalnizca yanit bekleyen isler icindir.", vbInformation
        Exit Sub
    End If

    Set kaynak = KaynakMail(gorev)
    alici = Trim$(AlanOku(gorev, ALAN_ALICI))
    If Len(alici) = 0 And Not kaynak Is Nothing Then alici = HatirlatmaAlicisi(kaynak)
    konu = GorunenKonu(gorev)
    If Len(Trim$(konu)) = 0 Then konu = gorev.Subject
    If LCase$(Left$(Trim$(konu), 3)) <> "re:" Then konu = "RE: " & konu

    metin = TakipMetni(gorev, 9000)
    govde = "<html><body style='font-family:Calibri,Arial,sans-serif;font-size:12pt;color:#243444'>" & _
            "Merhaba,<br><br>Asagidaki yazismamizla ilgili guncel durum hakkinda bilgi " & _
            "paylasmanizi rica ederim.<br><br>Iyi calismalar.<br><br>" & _
            "<table width='100%' cellpadding='12' style='background:#f3f6fa;border:1px solid #dce5ef'>" & _
            "<tr><td><b>Onceki yazisma icerigi</b><br>" & _
            "<div style='font-size:10pt;color:#526172;margin-top:10px'>" & _
            HtmlSatirli(metin) & "</div></td></tr></table>" & _
            "</body></html>"
    Set taslak = Application.CreateItem(olMailItem)
    taslak.To = alici
    taslak.Subject = konu
    taslak.BodyFormat = olFormatHTML
    taslak.HTMLBody = govde
    taslak.Display
    ' Sadece gosterilir; Save / Send yok.
    If Len(alici) = 0 Then
        MsgBox "Alici bulunamadi. Taslagi gondermeden once aliciyi belirtin.", vbExclamation
    Else
        MsgBox "Hatýrlatma taslaðý oluþturuldu. Alýcýyý, yazýþma metnini ve alýntýyý kontrol edin." & vbCrLf & _
               "Hiçbir e-posta otomatik gönderilmedi.", vbInformation
    End If
    Exit Sub
Hata:
    HataGoster "Hatirlatma taslagi"
End Sub


' 9) Kullanici kod aramadan Outlook'taki kendi gorev klasorunu acabilir.
Public Sub Takip_9_GorevleriAc()
    On Error GoTo Hata
    Set Application.ActiveExplorer.CurrentFolder = TakipKlasoru()
    Exit Sub
Hata:
    HataGoster "Görevler klasörünü açma"
End Sub

Private Function KaynakMail(ByVal gorev As Outlook.TaskItem) As Outlook.MailItem
    On Error GoTo Bitti
    Dim nesne As Object, kimlik As String, magaza As String
    kimlik = AlanOku(gorev, ALAN_KAYNAK)
    magaza = AlanOku(gorev, ALAN_POSTA)
    If Len(kimlik) = 0 Then Exit Function
    If Len(magaza) > 0 Then
        Set nesne = Application.Session.GetItemFromID(kimlik, magaza)
    Else
        Set nesne = Application.Session.GetItemFromID(kimlik)
    End If
    If TypeOf nesne Is Outlook.MailItem Then Set KaynakMail = nesne
Bitti:
End Function

Private Function HatirlatmaAlicisi(ByVal kaynak As Outlook.MailItem) As String
    On Error GoTo Bitti
    Dim k As Outlook.ExchangeUser, adres As String
    If BenimMesajim(kaynak) Then
        HatirlatmaAlicisi = kaynak.To
    Else
        adres = kaynak.SenderEmailAddress
        If kaynak.SenderEmailType = "EX" Then
            Set k = kaynak.Sender.GetExchangeUser
            If Not k Is Nothing Then adres = k.PrimarySmtpAddress
        End If
        If InStr(1, adres, "@") > 0 Then HatirlatmaAlicisi = adres
    End If
Bitti:
End Function
