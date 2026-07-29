Attribute VB_Name = "modKurulum"
'===============================================================================
' modKurulum - Calisma kitabini sifirdan insa eder
'
' CALISTIRILACAK MAKRO: Kurulum
'
' Bos bir kitapta bu modulleri ice aktarip Kurulum'u calistir. Gerekli tum
' sayfalari, sutun genisliklerini, satir yuksekliklerini ve ornek veriyi kurar.
' Ikinci kez calistirmak var olan veriyi SILMEZ; sadece eksikleri tamamlar.
' Her seyi bastan kurmak icin KurulumSifirla kullan (onay ister).
'===============================================================================
Option Explicit

'===============================================================================
' Giris noktalari
'===============================================================================
Public Sub Kurulum()
    On Error GoTo Hata
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    VeriSayfalariniKur
    AyarlariKur
    If modVeri.KullaniciSayisi() = 0 Then OrnekVeriYukle
    CizelgeSayfasiniKur
    GirisSayfasiniKur

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    modGiris.GirisiGoster
    MsgBox "Kurulum tamamlandi." & vbCrLf & vbCrLf & _
           "Ornek giris:" & vbCrLf & _
           "  Kullanici adi : admin" & vbCrLf & _
           "  Parola        : 1234" & vbCrLf & vbCrLf & _
           "Dosyayi 'Excel Makro Etkin Calisma Kitabi (*.xlsm)' olarak kaydetmeyi unutma.", _
           vbInformation, APP_ADI
    Exit Sub
Hata:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Kurulum sirasinda hata:" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, APP_ADI
End Sub

Public Sub KurulumSifirla()
    Dim cevap As VbMsgBoxResult
    cevap = MsgBox("TUM veriler silinecek ve her sey bastan kurulacak." & vbCrLf & _
                   "Devam edilsin mi?", vbExclamation + vbYesNo + vbDefaultButton2, APP_ADI)
    If cevap <> vbYes Then Exit Sub

    Application.DisplayAlerts = False
    Dim ws As Worksheet, adlar As Variant, i As Long
    adlar = Array(SH_GIRIS, SH_CIZELGE, SH_KULLANICILAR, SH_PROJELER, SH_KATEGORILER, _
                  SH_GOREVLER, SH_YORUMLAR, SH_GECIKMELER, SH_OFFGUNLER, SH_AYARLAR, SH_FOTOLAR)
    ' En az bir sayfa kalmali; gecici bir sayfa ac
    ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = "_gecici"
    For i = LBound(adlar) To UBound(adlar)
        On Error Resume Next
        ThisWorkbook.Worksheets(adlar(i)).Delete
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True

    Kurulum

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("_gecici").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

'===============================================================================
' Veri sayfalari
'===============================================================================
Private Sub VeriSayfalariniKur()
    SayfaKur SH_KULLANICILAR, Array("ID", "AdSoyad", "KullaniciAdi", "Parola", "Rol", _
             "Foto", "YGorevEkle", "YYorum", "YGecikme", "YKullaniciYonet", "Sayfa", "Aktif", "Renk")
    SayfaKur SH_PROJELER, Array("ID", "Ad", "Renk", "Sira", "Aktif")
    SayfaKur SH_KATEGORILER, Array("ID", "Ad", "Renk", "Sira", "Aktif")
    SayfaKur SH_GOREVLER, Array("ID", "ProjeID", "KategoriID", "Sira", "Ad", "Sorumlular", _
             "OncekiAyNot", "PlanBaslangic", "PlanBitis", "GerceklesenBitis", "Durum", _
             "CubukYazisi", "Yuzde", "Aciklama", "OlusturanID", "OlusturmaTarihi")
    SayfaKur SH_YORUMLAR, Array("ID", "GorevID", "KullaniciID", "Tarih", "Metin")
    SayfaKur SH_GECIKMELER, Array("ID", "GorevID", "Neden", "Sorumluluk", "KaydedenID", "Tarih")
    SayfaKur SH_OFFGUNLER, Array("Tarih", "Aciklama")
    SayfaKur SH_AYARLAR, Array("Anahtar", "Deger")
    SayfaKur SH_FOTOLAR, Array("KullaniciID", "Not")
End Sub

Private Sub SayfaKur(ByVal ad As String, ByVal basliklar As Variant)
    Dim ws As Worksheet, i As Long, yeni As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ad
        yeni = True
    End If

    If yeni Or ws.Cells(1, 1).Value = "" Then
        For i = LBound(basliklar) To UBound(basliklar)
            ws.Cells(1, i + 1).Value = basliklar(i)
        Next i
        With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(basliklar) + 1))
            .Font.Bold = True
            .Interior.Color = CLR_BASLIK_ZEMIN()
        End With
        ws.Rows(1).AutoFilter
    End If

    ws.Visible = xlSheetVeryHidden
End Sub

Private Sub AyarlariKur()
    If modVeri.Ayar(AYAR_AKTIF_YIL, "") = "" Then
        modVeri.AyarYaz AYAR_AKTIF_YIL, Year(Date)
        modVeri.AyarYaz AYAR_AKTIF_AY, Month(Date)
    End If
    If modVeri.Ayar(AYAR_OTURUM, "") = "" Then modVeri.AyarYaz AYAR_OTURUM, 0
End Sub

'===============================================================================
' Cizelge sayfasi iskeleti
'===============================================================================
Private Sub CizelgeSayfasiniKur()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CIZELGE)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SH_CIZELGE
    End If

    ws.Visible = xlSheetVisible
    ws.Cells.Clear
    ws.Cells.Interior.Color = CLR_ZEMIN()
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 9
    ws.Cells.Font.Color = CLR_METIN()

    ' Sabit sol sutunlar
    SutunPt ws, C_SERIT, W_SERIT
    SutunPt ws, C_KATEGORI, W_KATEGORI
    SutunPt ws, C_SIRA, W_SIRA
    SutunPt ws, C_SORUMLU, W_SORUMLU
    SutunPt ws, C_GOREV, W_GOREV
    SutunPt ws, C_ONCEKI, W_ONCEKI

    ' Gun sutunlari + sagdaki iki sutun (en genis ay olan 31 gune gore kur)
    Dim d As Long
    For d = 0 To GUN_SAYISI_MAX - 1
        SutunPt ws, C_GUN1 + d, W_GUN
    Next d
    SutunPt ws, C_GUN1 + GUN_SAYISI_MAX, W_ILERLEME
    SutunPt ws, C_GUN1 + GUN_SAYISI_MAX + 1, W_ACIKLAMA

    ' Satir yukseklikleri
    ws.Rows(R_BASLIK).RowHeight = 46
    ws.Rows(R_BASLIK + 1).RowHeight = 6
    ws.Rows(R_EFSANE).RowHeight = 20
    ws.Rows(R_EFSANE + 1).RowHeight = 6
    ws.Rows(R_AYADI).RowHeight = H_AYADI
    ws.Rows(R_GUNHARF).RowHeight = H_GUNHARF
    ws.Rows(R_GUNNO).RowHeight = H_GUNNO

    ' Gorunum: kilavuz cizgileri ve basliklar kapali
    Dim wnd As Window
    Set wnd = ActiveWindow
    If Not wnd Is Nothing Then
        ws.Activate
        wnd.DisplayGridlines = False
        wnd.DisplayHeadings = False
        wnd.Zoom = 100
    End If

    ' Baslik metni
    With ws.Cells(R_BASLIK, C_KATEGORI)
        .Value = APP_ADI
        .Font.Size = 15
        .Font.Bold = True
        .Font.Color = Hx("0F1622")
    End With

    ws.Cells(1, 1).Select
End Sub

' Excel sutun genisligi "karakter" cinsindendir, nokta degil. Hedef nokta
' genisligine ulasmak icin olcup oranla duzeltiyoruz (3 tur yeterli).
Private Sub SutunPt(ByVal ws As Worksheet, ByVal sutun As Long, ByVal hedefPt As Double)
    Dim i As Long, mevcutPt As Double, cw As Double
    ws.Columns(sutun).ColumnWidth = 8
    For i = 1 To 3
        mevcutPt = ws.Columns(sutun).Width
        If mevcutPt <= 0 Then Exit Sub
        cw = ws.Columns(sutun).ColumnWidth * hedefPt / mevcutPt
        If cw < 0.1 Then cw = 0.1
        If cw > 255 Then cw = 255
        ws.Columns(sutun).ColumnWidth = cw
    Next i
End Sub

'===============================================================================
' Giris sayfasi
'===============================================================================
Private Sub GirisSayfasiniKur()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_GIRIS)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SH_GIRIS
    End If

    ws.Visible = xlSheetVisible
    ws.Cells.Clear
    On Error Resume Next
    ws.Shapes.SelectAll
    ws.Shapes.Range(1).Delete
    Dim o As OLEObject
    For Each o In ws.OLEObjects
        o.Delete
    Next o
    On Error GoTo 0

    ws.Cells.Interior.Color = Hx("0F1115")     ' koyu zemin
    ws.Cells.Font.Name = "Segoe UI"
    modGiris.GirisTuvaliniCiz ws
End Sub

'===============================================================================
' Ornek veri - tasarimda onayladigimiz tablonun birebir ayni
'===============================================================================
Private Sub OrnekVeriYukle()
    Dim ws As Worksheet, r As Long, i As Long

    '--- Kullanicilar ---
    Dim kisiler As Variant
    kisiler = Array( _
        Array("Yigit Bayir", "admin", "1234", ROL_ADMIN, "4F46E5"), _
        Array("Adem Kaya", "adem.kaya", "1234", ROL_KULLANICI, "0891B2"), _
        Array("Merve Tunc", "merve.tunc", "1234", ROL_KULLANICI, "DB2777"), _
        Array("Selim Er", "selim.er", "1234", ROL_KULLANICI, "16A34A"), _
        Array("Naz Demir", "naz.demir", "1234", ROL_ADMIN, "F59E0B"), _
        Array("Kaan Oz", "kaan.oz", "1234", ROL_KULLANICI, "7C3AED"), _
        Array("Ece Gur", "ece.gur", "1234", ROL_KULLANICI, "0EA5E9"))
    For i = LBound(kisiler) To UBound(kisiler)
        modVeri.KullaniciEkle kisiler(i)(0), kisiler(i)(1), kisiler(i)(2), _
                              kisiler(i)(3), kisiler(i)(4)
    Next i

    '--- Projeler ---
    Set ws = modVeri.Sayfa(SH_PROJELER)
    Dim projeler As Variant
    projeler = Array(Array("Veri Merkezi Tasima", "4F46E5"), _
                     Array("Son Kullanici Donusumu", "0891B2"), _
                     Array("Is Surekliligi", "7C3AED"))
    For i = LBound(projeler) To UBound(projeler)
        r = i + 2
        ws.Cells(r, P_ID).Value = i + 1
        ws.Cells(r, P_AD).Value = projeler(i)(0)
        ws.Cells(r, P_RENK).Value = projeler(i)(1)
        ws.Cells(r, P_SIRA).Value = i + 1
        ws.Cells(r, P_AKTIF).Value = True
    Next i

    '--- Kategoriler ---
    Set ws = modVeri.Sayfa(SH_KATEGORILER)
    Dim katlar As Variant
    katlar = Array(Array("MAJOR", "4F46E5"), Array("MINOR", "F59E0B"), _
                   Array("VDI", "0891B2"), Array("BAKIM", "7C3AED"), _
                   Array("ACIL", "DC2626"))
    For i = LBound(katlar) To UBound(katlar)
        r = i + 2
        ws.Cells(r, T_ID).Value = i + 1
        ws.Cells(r, T_AD).Value = katlar(i)(0)
        ws.Cells(r, T_RENK).Value = katlar(i)(1)
        ws.Cells(r, T_SIRA).Value = i + 1
        ws.Cells(r, T_AKTIF).Value = True
    Next i

    '--- Gorevler (tasarim mockupindaki tablo) ---
    Dim y As Long, ay As Long
    y = Year(Date): ay = Month(Date)

    GorevYaz 1, 1, 1, 1, "Sunucu gocu - faz 2", "1,2,4", "Onceki aydan 2 gun devretti", _
             DateSerial(y, ay, 1), DateSerial(y, ay, 8), 0, DURUM_TAMAM, "Tooling", 100, _
             "Devir teslim tamamlandi"
    GorevYaz 2, 1, 1, 2, "Omurga switch degisimi", "3,6", "", _
             DateSerial(y, ay, 6), DateSerial(y, ay, 13), DateSerial(y, ay, 18), _
             DURUM_TAMAM, "Malzeme alimi", 100, "Tutanak imzalandi"
    GorevYaz 3, 1, 2, 3, "Kablolama denetimi", "5", "", _
             DateSerial(y, ay, 14), DateSerial(y, ay, 17), 0, DURUM_DEVAM, "Saha", 70, _
             "3. kat kaldi"
    GorevYaz 4, 2, 3, 4, "Kullanici profil tasima", "4,5,7,2", "", _
             DateSerial(y, ay, 20), DateSerial(y, ay, 31), 0, DURUM_DEVAM, "Migrasyon", 45, _
             "120 profilin 54'u tasindi"
    GorevYaz 5, 2, 2, 5, "Lisans yenileme", "1", "Sozlesme onayi bekleniyor", _
             DateSerial(y, ay, 15), DateSerial(y, ay, 20), 0, DURUM_DEVAM, "Satin alma", 60, _
             "Onay bekleniyor"
    GorevYaz 6, 3, 1, 6, "Yedekleme geri donus testi", "2,6", "", _
             DateSerial(y, ay, 27), DateSerial(y, ay, 31), 0, DURUM_PLAN, "Test", 0, _
             "Test ortami kuruluyor"
    GorevYaz 7, 3, 1, 7, "Felaket tatbikati", "7", "", _
             DateSerial(y, ay, 30), DateSerial(y, ay + 1, 6), 0, DURUM_PLAN, "Tatbikat", 0, _
             "Sonraki aya sarkiyor"

    modVeri.AyarYaz AYAR_SON_GOREV_ID, 7

    '--- Ornek yorumlar ---
    modVeri.AyarYaz AYAR_OTURUM, 4      ' Selim Er adina
    modVeri.YorumEkle 4, "3. kattaki 18 profil tamam, kalanlar yarin. Agda yavaslama var."
    modVeri.AyarYaz AYAR_OTURUM, 5      ' Naz Demir
    modVeri.YorumEkle 4, "Yavaslamayi ag ekibine ilettim, switch degisiminden olabilir."
    modVeri.AyarYaz AYAR_OTURUM, 1      ' Yigit Bayir
    modVeri.YorumEkle 4, "Go-Live tarihini sabit tutalim, gerekirse hafta sonu vardiya acariz."
    modVeri.YorumEkle 3, "Izin sureci uzadi, 3 gun ek sure gerekebilir."
    modVeri.AyarYaz AYAR_OTURUM, 0

    '--- Ornek gecikme kaydi ---
    modVeri.GecikmeKaydet 2, "Tedarikci sevkiyati 4 gun gecikti, gumrukte bekledi.", "Dis yuklenici"

    '--- Ornek tatil gunu ---
    modVeri.OffGunDegistir DateSerial(y, ay, 16), "Idari tatil"
End Sub

Private Sub GorevYaz(ByVal id As Long, ByVal projeID As Long, ByVal katID As Long, _
                     ByVal sira As Long, ByVal ad As String, ByVal sorumlular As String, _
                     ByVal onceki As String, ByVal planBas As Date, ByVal planBit As Date, _
                     ByVal gercek As Variant, ByVal durum As String, ByVal cubuk As String, _
                     ByVal yuzde As Double, ByVal aciklama As String)
    Dim ws As Worksheet, r As Long
    Set ws = modVeri.Sayfa(SH_GOREVLER)
    r = modVeri.SonSatir(ws) + 1
    ws.Cells(r, G_ID).Value = id
    ws.Cells(r, G_PROJE).Value = projeID
    ws.Cells(r, G_KATEGORI).Value = katID
    ws.Cells(r, G_SIRA).Value = sira
    ws.Cells(r, G_AD).Value = ad
    ws.Cells(r, G_SORUMLU).Value = sorumlular
    ws.Cells(r, G_ONCEKI).Value = onceki
    ws.Cells(r, G_PLANBAS).Value = planBas
    ws.Cells(r, G_PLANBIT).Value = planBit
    If CDbl(gercek) > 0 Then ws.Cells(r, G_GERCEK).Value = CDate(gercek)
    ws.Cells(r, G_DURUM).Value = durum
    ws.Cells(r, G_CUBUK).Value = cubuk
    ws.Cells(r, G_YUZDE).Value = yuzde
    ws.Cells(r, G_ACIKLAMA).Value = aciklama
    ws.Cells(r, G_OLUSTURAN).Value = 1
    ws.Cells(r, G_TARIH).Value = Now
End Sub
