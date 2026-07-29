Attribute VB_Name = "modModal"
'===============================================================================
' modModal - Genel amacli panel motoru
'
' Projedeki TUM modaller (gorev ekle, kullanici ekle, gecikme nedeni, baski
' ayarlari...) bu tek motoru kullanir. Her modal icin ayri cizim kodu yazmak
' yerine alan listesi tanimlanir, panel kendini kurar.
'
' Kullanim:
'   modModal.Temizle
'   modModal.Alan "Gorev adi", MT_METIN, "Sunucu gocu"
'   modModal.Alan "Kategori", MT_LISTE, "MAJOR", "MAJOR,MINOR,VDI"
'   modModal.Goster "gorev_ekle", "Gorev ekle"
'
' Tamam'a basildiginda TamamTiklandi calisir ve amaca gore ilgili modulun
' kaydetme yordamina yonlendirir. Iptal'de cizelgeye donulur.
'
' Deger okuma:  modModal.Deger(1), modModal.Deger(2) ...
'===============================================================================
Option Explicit

Public Const SH_MODAL As String = "_Panel"

' Alan tipleri
Public Const MT_METIN    As Long = 0
Public Const MT_COKSATIR As Long = 1
Public Const MT_SAYI     As Long = 2
Public Const MT_TARIH    As Long = 3
Public Const MT_LISTE    As Long = 4
Public Const MT_SALTOKUR As Long = 5
Public Const MT_BASLIK   As Long = 6      ' bolum ayirici, giris alani degil
Public Const MT_ONAY     As Long = 7      ' Evet/Hayir

Private Const ILK_SATIR   As Long = 4
Private Const C_ETIKET    As Long = 2
Private Const C_GIRIS     As Long = 3

Private mAmac      As String
Private mAlanAdi() As String
Private mAlanTip() As Long
Private mAlanSecim() As String
Private mAlanDeger() As String
Private mSatir(1 To 40) As Long    ' alan no -> panel satiri (Deger okurken lazim)
Private mSayi      As Long
Private mBaglam    As String       ' modalin uzerinde calistigi kayit (orn. gorev ID)

'===============================================================================
' Tanimlama
'===============================================================================
Public Sub Temizle()
    Dim i As Long
    mSayi = 0
    mAmac = ""
    mBaglam = ""
    ReDim mAlanAdi(1 To 40)
    ReDim mAlanTip(1 To 40)
    ReDim mAlanSecim(1 To 40)
    ReDim mAlanDeger(1 To 40)
    For i = 1 To 40: mSatir(i) = 0: Next i
End Sub

Public Sub Alan(ByVal etiket As String, ByVal tip As Long, _
                Optional ByVal deger As String = "", _
                Optional ByVal secenekler As String = "")
    If mSayi >= 40 Then Exit Sub
    mSayi = mSayi + 1
    mAlanAdi(mSayi) = etiket
    mAlanTip(mSayi) = tip
    mAlanDeger(mSayi) = deger
    mAlanSecim(mSayi) = secenekler
End Sub

Public Property Let Baglam(ByVal v As String)
    mBaglam = v
End Property

Public Property Get Baglam() As String
    Baglam = mBaglam
End Property

'===============================================================================
' Gosterim
'===============================================================================
Public Sub Goster(ByVal amac As String, ByVal baslik As String, _
                  Optional ByVal tamamMetni As String = "Kaydet", _
                  Optional ByVal tehlikeli As Boolean = False)
    Dim ws As Worksheet, i As Long, r As Long
    Dim onceki As Boolean

    mAmac = amac
    onceki = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set ws = PanelSayfasi()
    ws.Visible = xlSheetVisible
    ws.Unprotect
    ws.Cells.Clear
    SekilleriSil ws

    ws.Cells.Interior.Color = Hx("0E1116")
    ws.Cells.Font.Name = "Segoe UI"
    ws.Columns(1).ColumnWidth = 3
    ws.Columns(C_ETIKET).ColumnWidth = 22
    ws.Columns(C_GIRIS).ColumnWidth = 46
    ws.Columns(4).ColumnWidth = 3

    ' --- baslik ---
    With ws.Cells(2, C_ETIKET)
        .Value = baslik
        .Font.Size = 15
        .Font.Bold = True
        .Font.Color = Hx("E8EAF0")
    End With
    ws.Rows(2).RowHeight = 34
    ws.Rows(3).RowHeight = 8

    ' --- alanlar ---
    r = ILK_SATIR
    For i = 1 To mSayi
        If mAlanTip(i) = MT_BASLIK Then
            ws.Rows(r).RowHeight = 26
            With ws.Cells(r, C_ETIKET)
                .Value = UCase$(mAlanAdi(i))
                .Font.Size = 7.5
                .Font.Bold = True
                .Font.Color = Hx("6B7588")
            End With
            r = r + 1
        Else
            ' Etiket satiri - giris hucresinin tam ustunde, ayni sutunda
            ws.Rows(r).RowHeight = 15
            With ws.Cells(r, C_GIRIS)
                .Value = UCase$(mAlanAdi(i))
                .Font.Size = 7
                .Font.Bold = True
                .Font.Color = Hx("7A8394")
                .VerticalAlignment = xlBottom
                .HorizontalAlignment = xlLeft
            End With
            r = r + 1

            ' Giris satiri
            ws.Rows(r).RowHeight = IIf(mAlanTip(i) = MT_COKSATIR, 52, 22)
            GirisHucresiKur ws.Cells(r, C_GIRIS), i
            mSatir(i) = r
            r = r + 1

            ' Alanlar arasi bosluk
            ws.Rows(r).RowHeight = 5
            r = r + 1
        End If
    Next i

    ' --- butonlar ---
    Dim y As Double, x As Double
    y = ws.Cells(r + 1, C_GIRIS).Top
    x = ws.Cells(r + 1, C_GIRIS).Left
    ButonCiz ws, "iptal", "Iptal", x, y, 90, Hx("232936"), Hx("9AA3B2"), "modModal.IptalTiklandi"
    ButonCiz ws, "tamam", tamamMetni, x + 98, y, 130, _
             IIf(tehlikeli, Hx("DC2626"), Hx("4F46E5")), Hx("FFFFFF"), "modModal.TamamTiklandi"

    ' --- telif ---
    With ws.Cells(r + 4, C_GIRIS)
        .Value = APP_TELIF
        .Font.Size = 7
        .Font.Color = Hx("3C4454")
    End With

    ' Sadece giris hucreleri duzenlenebilir olsun; Tab ile alanlar arasi gecis
    ' bu sayede calisir.
    ws.Protect UserInterfaceOnly:=True, DrawingObjects:=False, Contents:=True

    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    On Error GoTo 0
    IlkGirisHucresiniSec ws

    Application.ScreenUpdating = onceki
End Sub

Private Sub GirisHucresiKur(ByVal hc As Range, ByVal alanNo As Long)
    With hc
        .Value = mAlanDeger(alanNo)
        .Font.Size = 10
        .Font.Color = Hx("D7DDE8")
        .Interior.Color = Hx("171C24")
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .WrapText = (mAlanTip(alanNo) = MT_COKSATIR)
        .Locked = (mAlanTip(alanNo) = MT_SALTOKUR)
        With .Borders
            .LineStyle = xlContinuous
            .Color = Hx("2C3340")
            .Weight = xlThin
        End With
    End With

    Select Case mAlanTip(alanNo)
        Case MT_SALTOKUR
            hc.Font.Color = Hx("8A93A3")
            hc.Interior.Color = Hx("12161C")
        Case MT_TARIH
            hc.NumberFormat = "dd.mm.yyyy"
        Case MT_SAYI
            hc.NumberFormat = "0"
            hc.HorizontalAlignment = xlLeft
        Case MT_LISTE, MT_ONAY
            Dim liste As String
            liste = IIf(mAlanTip(alanNo) = MT_ONAY, "Evet,Hayir", mAlanSecim(alanNo))
            On Error Resume Next
            hc.Validation.Delete
            hc.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                              Operator:=xlBetween, Formula1:=liste
            hc.Validation.InCellDropdown = True
            On Error GoTo 0
    End Select
End Sub

'===============================================================================
' Deger okuma
'===============================================================================
Public Function Deger(ByVal alanNo As Long) As String
    If alanNo < 1 Or alanNo > mSayi Then Exit Function
    If mSatir(alanNo) = 0 Then Exit Function
    Dim v As Variant
    v = PanelSayfasi().Cells(mSatir(alanNo), C_GIRIS).Value
    If IsDate(v) And mAlanTip(alanNo) = MT_TARIH Then
        Deger = Format$(v, "dd.mm.yyyy")
    Else
        Deger = CStr(v)
    End If
End Function

Public Function DegerTarih(ByVal alanNo As Long) As Date
    Dim v As Variant
    v = PanelSayfasi().Cells(mSatir(alanNo), C_GIRIS).Value
    If IsDate(v) Then DegerTarih = CDate(v)
End Function

Public Function DegerSayi(ByVal alanNo As Long) As Double
    DegerSayi = Val(Replace(Deger(alanNo), ",", "."))
End Function

Public Function DegerOnay(ByVal alanNo As Long) As Boolean
    DegerOnay = (StrComp(Deger(alanNo), "Evet", vbTextCompare) = 0)
End Function

Public Function Amac() As String
    Amac = mAmac
End Function

'===============================================================================
' Buton olaylari
'===============================================================================
Public Sub TamamTiklandi()
    Dim ok As Boolean
    On Error GoTo Hata

    Select Case mAmac
        Case "gorev_ekle":      ok = modGorev.GorevKaydet(False)
        Case "gorev_duzenle":   ok = modGorev.GorevKaydet(True)
        Case "kullanici_ekle":  ok = modKullanici.KullaniciKaydet()
        Case "kategori_ekle":   ok = modKullanici.KategoriKaydet()
        Case "gecikme":         ok = modGecikme.GecikmeKaydetPanel()
        Case "yorum":           ok = modYorum.YorumKaydet()
        Case "baski":           ok = modBaski.BaskiUygula()
        Case "giris":           ok = modGiris.GirisDogrula()
        Case Else:              ok = True
    End Select

    If ok Then Kapat
    Exit Sub
Hata:
    MsgBox "Kaydetme hatasi: " & Err.Number & " - " & Err.Description, vbExclamation, APP_ADI
End Sub

Public Sub IptalTiklandi()
    Kapat
End Sub

Public Sub Kapat()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = PanelSayfasi()
    ws.Visible = xlSheetVeryHidden
    modVeri.Sayfa(SH_CIZELGE).Activate
    On Error GoTo 0
End Sub

'===============================================================================
' Yardimcilar
'===============================================================================
Public Function PanelSayfasi() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_MODAL)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_MODAL
    End If
    Set PanelSayfasi = ws
End Function

Private Sub SekilleriSil(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

Private Sub IlkGirisHucresiniSec(ByVal ws As Worksheet)
    Dim i As Long
    For i = 1 To mSayi
        If mAlanTip(i) <> MT_SALTOKUR And mAlanTip(i) <> MT_BASLIK And mSatir(i) > 0 Then
            ws.Cells(mSatir(i), C_GIRIS).Select
            Exit Sub
        End If
    Next i
End Sub

Public Sub ButonCiz(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, _
                    ByVal x As Double, ByVal y As Double, ByVal gen As Double, _
                    ByVal zemin As Long, ByVal yazi As Long, ByVal makro As String)
    Dim sk As Shape
    Set sk = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, gen, 30)
    sk.Name = "btn_" & ad
    sk.Adjustments(1) = 0.2
    modCizelge.SekilBicimle sk, zemin, 0
    With sk.TextFrame2
        .TextRange.Text = metin
        .TextRange.Font.Size = 9.5
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Fill.ForeColor.RGB = yazi
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
    End With
    sk.OnAction = makro
End Sub

' Panelin uzerine salt-okunur bilgi blogu basar (gecikme kutusundaki
' planlanan/gerceklesen/sapma ucglusu gibi).
Public Sub BilgiSeridi(ByVal ws As Worksheet, ByVal satir As Long, _
                       ByVal e1 As String, ByVal d1 As String, _
                       ByVal e2 As String, ByVal d2 As String, _
                       ByVal e3 As String, ByVal d3 As String, _
                       ByVal vurguRenk As Long)
    Dim metin As String
    metin = e1 & ": " & d1 & "        " & e2 & ": " & d2 & "        " & e3 & ": " & d3
    With ws.Cells(satir, C_GIRIS)
        .Value = metin
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = vurguRenk
        .Interior.Color = Hx("12161C")
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(satir).RowHeight = 26
End Sub
