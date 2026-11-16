Attribute VB_Name = "RadarIsMasam"
Option Explicit

' IS VE YANIT RADARI v0.8.2 - OZEL TEST / SQL'siz gruplu gorev panosu
' Yeni hesap, HTTP baglantisi, ek servis, otomatik e-posta gonderimi yok.
' Takip kayitlarini silmez veya baska klasore tasimaz.

Private Const RADAR_KLASOR As String = "Takip Kolay"
Private Const GOR_TUM As String = "Radar - Tum Acik Isler"
Private Const GOR_YAP As String = "Radar - Yapilacaklar"
Private Const GOR_YANIT As String = "Radar - Yanit Bekleyen"
Private Const GOR_INCELE As String = "Radar - Incelenecek"
Private Const GOR_BITEN As String = "Radar - Tamamlananlar"
Private Const ALAN_DURUM As String = "TKState"
Private Const ALAN_KISI As String = "TKPerson"
Private Const ALAN_KONU As String = "TKTopic"
Private Const ALAN_KAYNAK As String = "TKMailID"
Private Const ALAN_POSTA As String = "TKMailStore"
Private Const DUR_TODO As String = "TODO"
Private Const DUR_WAIT As String = "WAIT"
Private Const DUR_REVIEW As String = "REVIEW"
Private Const DUR_DONE As String = "DONE"

Private Const GOR_ISMASAM As String = "Radar - Is Masam (Gruplu)"

' Outlook'un yerlesik Gorevler gorunumunde filtre olmadan TKState ile gruplar.
' Gruplar: TODO / WAIT / REVIEW / DONE. Kayitlari degistirmez.
Public Sub Radar_11_IsMasam()
    On Error GoTo Hata
    Dim klasor As Outlook.MAPIFolder, pencere As Outlook.Explorer
    Dim aciklama As String, gruplu As Boolean, ayrinti As String
    Set klasor = RadarKlasoru()

    ' Klasorun kendisini her kosulda ac. Gorev kayitlari degistirilmez.
    Set pencere = Application.ActiveExplorer
    If pencere Is Nothing Then
        Set pencere = klasor.GetExplorer
        pencere.Display
    Else
        Set pencere.CurrentFolder = klasor
    End If

    ' Onceki surumlerin hatali SQL/DASL filtrelerini kullanmayiz.
    ' TKState alanina gore tek gorunumde gruplama yapariz.
    On Error GoTo GorunumHatasi
    RadarAlanlariKaydet klasor
    gruplu = RadarGrupluGorunumHazirla(klasor, aciklama)
    If gruplu Then
        pencere.CurrentView = GOR_ISMASAM
        Exit Sub
    End If
GorunumHatasi:
    ' Gruplama Outlook surumunde engellenmisse bile kullanici
    ' butun islerini Gorevler klasorunde gorebilir.
    If Err.Number <> 0 Then
        ayrinti = "Gorunum hatasi " & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        ayrinti = aciklama
    End If
    On Error Resume Next
    pencere.CurrentView = GOR_ISMASAM
    On Error GoTo Hata
    MsgBox "Is Masam klasoru acildi. Gruplandirilmis gorunum bu Outlook " & _
           "surumunde uygulanamadi." & vbCrLf & vbCrLf & _
           ayrinti & vbCrLf & vbCrLf & _
           "Gorevlerinize dokunulmadi. Mevcut gorev listesi aciktir.", _
           vbExclamation, "Is ve Yanit Radari"
    Exit Sub
Hata:
    MsgBox "Is Masam acilamadi. Hata " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Is ve Yanit Radari"
End Sub

' Gorev secip bu komutu calistirirsaniz kaynak e-posta acilir.
' Mesaj silinmis/tasinmissa yanlis bir yazismayi tahminen acmaz.
Public Sub Radar_12_KaynakYazismayiAc()
    On Error GoTo Hata
    Dim gorev As Outlook.TaskItem, nesne As Object
    Dim kimlik As String, posta As String
    Set gorev = RadarSecilenGorev()
    If gorev Is Nothing Then Exit Sub
    kimlik = RadarAlanOku(gorev, ALAN_KAYNAK)
    posta = RadarAlanOku(gorev, ALAN_POSTA)
    If Len(kimlik) = 0 Then
        MsgBox "Bu gorevde kaynak e-posta referansi yok.", vbInformation
        Exit Sub
    End If

    On Error GoTo KaynakYok
    If Len(posta) > 0 Then
        Set nesne = Application.Session.GetItemFromID(kimlik, posta)
    Else
        Set nesne = Application.Session.GetItemFromID(kimlik)
    End If
    On Error GoTo Hata
    If nesne Is Nothing Then GoTo KaynakYok
    If Not TypeOf nesne Is Outlook.MailItem Then GoTo KaynakYok
    nesne.Display
    Exit Sub
KaynakYok:
    MsgBox "Kaynak e-posta artik bu kimlikle bulunamiyor." & vbCrLf & _
           "Mesaj tasinmis, arsivlenmis veya silinmis olabilir." & vbCrLf & _
           "Gorev kaydiniz degistirilmedi.", vbInformation, "Is ve Yanit Radari"
    Exit Sub
Hata:
    MsgBox "Kaynak yazisma acilamadi. Hata " & Err.Number & _
           ": " & Err.Description, vbExclamation, "Is ve Yanit Radari"
End Sub

' Secili gorevi TODO / WAIT / REVIEW arasinda kullanici karar vererek tasir.
' Fiziksel klasor tasinmaz. Tamamlamak icin mevcut Takip_5_Tamamlandi kullanilir.
Public Sub Radar_13_KategoriDegistir()
    On Error GoTo Hata
    Dim gorev As Outlook.TaskItem, cevap As String, kod As String
    Dim alan As Outlook.UserProperty
    Set gorev = RadarSecilenGorev()
    If gorev Is Nothing Then Exit Sub
    If gorev.Status = olTaskComplete Then
        MsgBox "Tamamlanan isi once Outlook gorevinde yeniden acin.", vbInformation
        Exit Sub
    End If
    cevap = Trim$(InputBox("Gorevi hangi kategoriye alalim?" & vbCrLf & _
                          "1 - Yapacagim is" & vbCrLf & _
                          "2 - Yanit bekliyorum" & vbCrLf & _
                          "3 - Incelenecek" & vbCrLf & _
                          "Bos/Iptal - Degistirme", "Is Masam - Kategori", ""))
    Select Case cevap
        Case "1": kod = DUR_TODO
        Case "2": kod = DUR_WAIT
        Case "3": kod = DUR_REVIEW
        Case "": Exit Sub
        Case Else
            MsgBox "Gecersiz secenek, kayit degistirilmedi.", vbInformation
            Exit Sub
    End Select
    Set alan = gorev.UserProperties.Find(ALAN_DURUM)
    If alan Is Nothing Then Set alan = gorev.UserProperties.Add(ALAN_DURUM, olText, False)
    alan.Value = kod
    gorev.Save
    MsgBox "Kategori kaydedildi. Gorunumu yenilemek icin F5 kullanabilirsiniz.", vbInformation
    Exit Sub
Hata:
    MsgBox "Kategori degistirilemedi. Hata " & Err.Number & _
           ": " & Err.Description, vbExclamation, "Is ve Yanit Radari"
End Sub

Private Function RadarKlasoru() As Outlook.MAPIFolder
    Dim ana As Outlook.MAPIFolder, alt As Outlook.MAPIFolder
    Set ana = Application.Session.GetDefaultFolder(olFolderTasks)
    For Each alt In ana.Folders
        If StrComp(alt.Name, RADAR_KLASOR, vbTextCompare) = 0 Then
            Set RadarKlasoru = alt
            Exit Function
        End If
    Next alt
    Set RadarKlasoru = ana.Folders.Add(RADAR_KLASOR, olFolderTasks)
End Function

Private Sub RadarAlanlariKaydet(ByVal klasor As Outlook.MAPIFolder)
    If klasor.UserDefinedProperties.Find(ALAN_DURUM) Is Nothing Then
        klasor.UserDefinedProperties.Add ALAN_DURUM, olText
    End If
    If klasor.UserDefinedProperties.Find(ALAN_KISI) Is Nothing Then
        klasor.UserDefinedProperties.Add ALAN_KISI, olText
    End If
    If klasor.UserDefinedProperties.Find(ALAN_KONU) Is Nothing Then
        klasor.UserDefinedProperties.Add ALAN_KONU, olText
    End If
End Sub

Private Function RadarGrupluGorunumHazirla(ByVal klasor As Outlook.MAPIFolder, _
        ByRef neden As String) As Boolean
    On Error GoTo Basarisiz
    Dim gorunum As Outlook.View, tablo As Outlook.TableView
    Set gorunum = Nothing
    On Error Resume Next
    Set gorunum = klasor.Views.Item(GOR_ISMASAM)
    Err.Clear
    On Error GoTo Basarisiz
    If gorunum Is Nothing Then
        Set gorunum = klasor.Views.Add(GOR_ISMASAM, olTableView, olViewSaveOptionThisFolderOnlyMe)
    End If
    If gorunum.ViewType <> olTableView Then
        neden = "Ayni adli gorunum tablo turunde degil."
        Exit Function
    End If
    Set tablo = gorunum
    ' Yalnizca bu yeni gorunume dokunulur, gorevlere ASLA dokunulmaz.
    ' Bos filtre kullanilir. @SQL, DASL, MAPI filtreleri YOK.
    ' Bu gorunume Filter degeri hic atanmaz. Parser kullanilmaz.
    tablo.AllowInCellEditing = False
    tablo.AutomaticColumnSizing = True
    tablo.AutomaticGrouping = False
    tablo.GroupByFields.RemoveAll
    tablo.GroupByFields.Add ALAN_DURUM
    tablo.ShowItemsInGroups = True

    RadarAlanEkle tablo, ALAN_DURUM
    RadarAlanEkle tablo, ALAN_KISI
    RadarAlanEkle tablo, ALAN_KONU
    tablo.Save
    RadarGrupluGorunumHazirla = True
    Exit Function
Basarisiz:
    neden = "Gorunum olusturma / gruplama hatasi " & Err.Number & ": " & Err.Description
    ' Outlook eski filtreyi geri tutmasin: en azindan gorunumu temiz
    ' ve filtresiz bir tablo olarak kaydetmeyi dene.
    On Error Resume Next
    If Not tablo Is Nothing Then
        tablo.GroupByFields.RemoveAll
        tablo.ShowItemsInGroups = False
        tablo.Save
    End If
    Err.Clear
End Function

Private Sub RadarAlanEkle(ByVal tablo As Outlook.TableView, ByVal alan As String)
    ' Opsiyonel gorunum sutunlari engellenirse ana gorunumu bozmayiz.
    Dim sutun As Outlook.ViewField
    On Error Resume Next
    Set sutun = tablo.ViewFields.Add(alan)
    Err.Clear
    On Error GoTo 0
End Sub

' Bir onceki testlerin olusturdugu sorunlu bes gorunumu istege bagli sil.
' Yalnizca bizim tam olarak bilinen gorunum adlari hedeflenir.
' GOREV KAYITLARI VE KLASORLAR SILINMEZ.
Public Sub Radar_14_BozukGorunumleriTemizle()
    On Error GoTo Hata
    Dim klasor As Outlook.MAPIFolder, i As Long, adet As Long
    Dim isimler(1 To 5) As String, gorunum As Outlook.View
    Dim metin As String
    isimler(1) = GOR_TUM
    isimler(2) = GOR_YAP
    isimler(3) = GOR_YANIT
    isimler(4) = GOR_INCELE
    isimler(5) = GOR_BITEN

    If MsgBox("Sadece onceki testlerin olusturdugu bes 'Radar - ...' " & _
              "gorunumunu kaldirmak istiyor musunuz?" & vbCrLf & vbCrLf & _
              "Gorevler ve yazismalar SILINMEZ. ", vbYesNo + vbQuestion, _
              "Is Masam - Eski Gorunumler") <> vbYes Then Exit Sub
    Set klasor = RadarKlasoru()
    ' Eski gorunumlerden biri aktifse once yeni gorunume gec.
    On Error Resume Next
    If Not Application.ActiveExplorer Is Nothing Then
        Set Application.ActiveExplorer.CurrentFolder = klasor
        Application.ActiveExplorer.CurrentView = GOR_ISMASAM
    End If
    Err.Clear
    On Error GoTo Hata
    For i = 1 To 5
        Set gorunum = Nothing
        On Error Resume Next
        Set gorunum = klasor.Views.Item(isimler(i))
        Err.Clear
        On Error GoTo Hata
        If Not gorunum Is Nothing Then
            If Not gorunum.Standard Then
                On Error Resume Next
                gorunum.Delete
                If Err.Number = 0 Then
                    adet = adet + 1
                Else
                    metin = metin & vbCrLf & isimler(i) & " : " & Err.Description
                End If
                Err.Clear
                On Error GoTo Hata
            End If
        End If
    Next i
    MsgBox "Kaldirilan eski gorunum: " & adet & vbCrLf & _
           "Takip gorevleri degistirilmedi." & metin, _
           vbInformation, "Is ve Yanit Radari"
    Exit Sub
Hata:
    MsgBox "Gorunum temizleme hatasi " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Is ve Yanit Radari"
End Sub

Private Function RadarSecilenGorev() As Outlook.TaskItem
    Dim pencere As Outlook.Explorer, nesne As Object
    Dim klasor As Outlook.MAPIFolder, kayitKlasoru As Outlook.MAPIFolder
    On Error GoTo Hata
    Set pencere = Application.ActiveExplorer
    If pencere Is Nothing Then GoTo SecimYok
    If pencere.Selection.Count <> 1 Then GoTo SecimYok
    Set nesne = pencere.Selection.Item(1)
    If Not TypeOf nesne Is Outlook.TaskItem Then GoTo SecimYok
    Set klasor = nesne.Parent
    Set kayitKlasoru = RadarKlasoru()
    If StrComp(klasor.EntryID, kayitKlasoru.EntryID, vbBinaryCompare) <> 0 Or _
       StrComp(klasor.StoreID, kayitKlasoru.StoreID, vbBinaryCompare) <> 0 Then
        GoTo SecimYok
    End If
    Set RadarSecilenGorev = nesne
    Exit Function
SecimYok:
    MsgBox "Once Is Masam (Takip Kolay) klasorunden tek bir gorev secin.", _
           vbInformation, "Is ve Yanit Radari"
    Exit Function
Hata:
    MsgBox "Secili gorev okunamadi. Hata " & Err.Number & _
           ": " & Err.Description, vbExclamation, "Is ve Yanit Radari"
End Function

Private Function RadarAlanOku(ByVal gorev As Outlook.TaskItem, _
                              ByVal alan As String) As String
    Dim bilgi As Outlook.UserProperty
    On Error GoTo Bitti
    Set bilgi = gorev.UserProperties.Find(alan)
    If Not bilgi Is Nothing Then
        If Not IsNull(bilgi.Value) Then RadarAlanOku = CStr(bilgi.Value)
    End If
Bitti:
End Function
