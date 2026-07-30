Attribute VB_Name = "modGiris"
'===============================================================================
' modGiris - Giris ekrani ve oturum yonetimi
'
' Giris sayfasi tamamen sayfaya cizilir (secilen "C" yaklasimi). Tek istisna
' parola kutusudur: hucreye yazilan metin yildizla gizlenemedigi icin oraya
' bir ActiveX metin kutusu konur. ActiveX olusturulamazsa (nadir) normal
' hucreye duser ve kullaniciyi uyarir.
'
' UYARI - guvenlik: Bu bir erisim duzenidir, kasa degildir. Parolalar basit
' bir karistirmayla saklanir ve VBA projesi/sayfa gizleme kolayca asilabilir.
' Gercek gizlilik gerektiren veriyi bu dosyada tutma.
'===============================================================================
Option Explicit

Private Const TXT_PAROLA As String = "txtParola"
Private Const TXT_KADI   As String = "txtKullanici"

'===============================================================================
' Giris tuvali
'===============================================================================
Public Sub GirisTuvaliniCiz(ByVal ws As Worksheet)
    Dim sk As Shape
    Dim panelX As Double, panelY As Double, panelW As Double, panelH As Double

    ws.Cells.Interior.Color = Hx("0E1116")
    ws.Columns(1).ColumnWidth = 2
    ws.Rows(1).RowHeight = 40

    panelX = 60: panelY = 50: panelW = 340: panelH = 320

    ' --- panel govdesi ---
    Set sk = ws.Shapes.AddShape(msoShapeRoundedRectangle, panelX, panelY, panelW, panelH)
    sk.Name = "gs_panel"
    sk.Adjustments(1) = 0.05
    modCizelge.SekilBicimle sk, Hx("161A21"), 0
    sk.Line.Visible = msoTrue
    sk.Line.ForeColor.RGB = Hx("232833")
    sk.Line.Weight = 1

    ' --- logo karesi ---
    Set sk = ws.Shapes.AddShape(msoShapeRoundedRectangle, panelX + 32, panelY + 34, 44, 44)
    sk.Name = "gs_logo"
    sk.Adjustments(1) = 0.22
    modCizelge.SekilBicimle sk, Hx("4F46E5"), 0
    YaziKur sk, "GX", 15, Hx("FFFFFF"), msoAlignCenter

    ' --- basliklar ---
    MetinKutusu ws, "gs_baslik", panelX + 32, panelY + 92, 280, 26, _
                APP_ADI, 16, Hx("E8EAF0"), True
    MetinKutusu ws, "gs_altbaslik", panelX + 32, panelY + 116, 280, 18, _
                "Devam etmek icin oturum ac", 9, Hx("7A8394"), False

    ' --- alan etiketleri ---
    MetinKutusu ws, "gs_lbl1", panelX + 32, panelY + 150, 200, 14, _
                "KULLANICI ADI", 7, Hx("7A8394"), True
    MetinKutusu ws, "gs_lbl2", panelX + 32, panelY + 208, 200, 14, _
                "PAROLA", 7, Hx("7A8394"), True

    ' --- giris kutulari (ActiveX) ---
    KutuKur ws, TXT_KADI, panelX + 32, panelY + 166, panelW - 64, 28, False
    KutuKur ws, TXT_PAROLA, panelX + 32, panelY + 224, panelW - 64, 28, True

    ' --- giris butonu ---
    Set sk = ws.Shapes.AddShape(msoShapeRoundedRectangle, panelX + 32, panelY + 264, _
                                panelW - 64, 34)
    sk.Name = "gs_btn"
    sk.Adjustments(1) = 0.2
    modCizelge.SekilBicimle sk, Hx("4F46E5"), 0
    YaziKur sk, "Giris yap", 10, Hx("FFFFFF"), msoAlignCenter
    sk.OnAction = "modGiris.GirisDenemesi"

    ' --- telif ---
    MetinKutusu ws, "gs_telif", panelX, panelY + panelH + 16, panelW, 16, _
                APP_TELIF, 8, Hx("3C4454"), False, msoAlignCenter

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    On Error GoTo 0
End Sub

Private Sub KutuKur(ByVal ws As Worksheet, ByVal ad As String, ByVal x As Double, _
                    ByVal y As Double, ByVal g As Double, ByVal h As Double, _
                    ByVal gizli As Boolean)
    Dim o As OLEObject
    On Error GoTo YedekPlan

    Set o = ws.OLEObjects.Add(ClassType:="Forms.TextBox.1", Link:=False, _
                              DisplayAsIcon:=False, Left:=x, Top:=y, Width:=g, Height:=h)
    o.Name = ad
    With o.Object
        .BackColor = RGB(15, 19, 25)
        .ForeColor = RGB(215, 221, 232)
        .BorderStyle = 1
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        If gizli Then .PasswordChar = ChrW(8226)
    End With
    Exit Sub

YedekPlan:
    ' ActiveX eklenemedi - hucreye dus
    On Error Resume Next
    Dim hc As Range
    Set hc = ws.Range("H5")
    If gizli Then Set hc = ws.Range("H7")
    hc.Value = ""
    hc.Name = ad & "_yedek"
    MetinKutusu ws, "gs_uyari", x, y, g, h, _
        "(ActiveX kutusu olusturulamadi - " & ad & ")", 8, Hx("F59E0B"), False
End Sub

Private Sub MetinKutusu(ByVal ws As Worksheet, ByVal ad As String, ByVal x As Double, _
                        ByVal y As Double, ByVal g As Double, ByVal h As Double, _
                        ByVal metin As String, ByVal boyut As Double, ByVal renk As Long, _
                        ByVal kalin As Boolean, Optional ByVal hiza As Long = msoAlignLeft)
    Dim sk As Shape
    Set sk = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, g, h)
    sk.Name = ad
    sk.Fill.Visible = msoFalse
    sk.Line.Visible = msoFalse
    With sk.TextFrame2
        .TextRange.Text = metin
        .TextRange.Font.Size = boyut
        .TextRange.Font.Bold = IIf(kalin, msoTrue, msoFalse)
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Fill.ForeColor.RGB = renk
        .TextRange.ParagraphFormat.Alignment = hiza
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoTrue
        .AutoSize = msoAutoSizeNone
    End With
End Sub

Private Sub YaziKur(ByVal sk As Shape, ByVal metin As String, ByVal boyut As Double, _
                    ByVal renk As Long, ByVal hiza As Long)
    With sk.TextFrame2
        .TextRange.Text = metin
        .TextRange.Font.Size = boyut
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Fill.ForeColor.RGB = renk
        .TextRange.ParagraphFormat.Alignment = hiza
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
    End With
End Sub

'===============================================================================
' Oturum
'===============================================================================
Public Sub GirisiGoster()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_GIRIS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    modVeri.AyarYaz AYAR_OTURUM, 0
    ws.Visible = xlSheetVisible
    ws.Activate

    ' Diger sayfalari gizle. Panel ve gecici rapor sayfalari "_" ile baslar ama
    ' acik kalmis olabilirler; onlari da kapatiyoruz.
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If s.Name <> SH_GIRIS Then
            If Left$(s.Name, 1) <> "_" Then
                s.Visible = xlSheetHidden
            ElseIf s.Visible = xlSheetVisible Then
                s.Visible = xlSheetVeryHidden
            End If
        End If
    Next s

    KutuTemizle ws
End Sub

Private Sub KutuTemizle(ByVal ws As Worksheet)
    On Error Resume Next
    ws.OLEObjects(TXT_KADI).Object.Text = ""
    ws.OLEObjects(TXT_PAROLA).Object.Text = ""
    On Error GoTo 0
End Sub

Public Sub GirisDenemesi()
    Dim ws As Worksheet, kadi As String, parola As String
    Dim kid As Long, u As TKullanici

    Set ws = ThisWorkbook.Worksheets(SH_GIRIS)

    On Error Resume Next
    kadi = Trim$(CStr(ws.OLEObjects(TXT_KADI).Object.Text))
    parola = CStr(ws.OLEObjects(TXT_PAROLA).Object.Text)
    On Error GoTo 0

    If Len(kadi) = 0 Then
        MsgBox "Kullanici adi bos birakilamaz.", vbExclamation, APP_ADI
        Exit Sub
    End If

    kid = modVeri.KullaniciAdiylaBul(kadi)
    If kid = 0 Then
        MsgBox "Kullanici bulunamadi.", vbExclamation, APP_ADI
        Exit Sub
    End If

    u = modVeri.KullaniciGetir(kid)
    If Not u.Aktif Then
        MsgBox "Bu kullanici pasif durumda.", vbExclamation, APP_ADI
        Exit Sub
    End If

    If Not modVeri.ParolaDogru(kid, parola) Then
        MsgBox "Parola hatali.", vbExclamation, APP_ADI
        Exit Sub
    End If

    modVeri.AyarYaz AYAR_OTURUM, kid
    OturumBaslat u
End Sub

' modModal uzerinden cagrilirsa (yedek yol)
Public Function GirisDogrula() As Boolean
    GirisDenemesi
    GirisDogrula = (CLng(Val(modVeri.Ayar(AYAR_OTURUM, 0))) > 0)
End Function

Private Sub OturumBaslat(ByRef u As TKullanici)
    Dim ws As Worksheet

    Application.ScreenUpdating = False

    ' Yetkisine gore sayfalari ac
    Set ws = modVeri.Sayfa(SH_CIZELGE)
    ws.Visible = xlSheetVisible

    ' Excel etkin sayfayi gizlemeye izin vermez. Giris sayfasi su anda etkin
    ' oldugu icin once cizelgeye geciyoruz; ters sirada giriste hata veriyordu.
    ws.Activate
    ThisWorkbook.Worksheets(SH_GIRIS).Visible = xlSheetVeryHidden

    modCizelge.Ciz
    KullaniciRozetiniGuncelle u

    ws.Activate
    Application.ScreenUpdating = True
End Sub

Private Sub KullaniciRozetiniGuncelle(ByRef u As TKullanici)
    Dim ws As Worksheet, sk As Shape, ad As String
    Set ws = modVeri.Sayfa(SH_CIZELGE)
    ad = SP_UI & "oturum"

    On Error Resume Next
    ws.Shapes(ad).Delete
    On Error GoTo 0

    Set sk = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
             ws.Cells(R_BASLIK, C_KATEGORI).Left, ws.Cells(R_BASLIK, C_KATEGORI).Top + 26, 320, 16)
    sk.Name = ad
    sk.Fill.Visible = msoFalse
    sk.Line.Visible = msoFalse
    With sk.TextFrame2
        .TextRange.Text = u.Ad & "  " & ChrW(183) & "  " & _
                          IIf(u.Rol = ROL_ADMIN, "Yonetici", "Kullanici") & _
                          "     " & APP_TELIF
        .TextRange.Font.Size = 8
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Fill.ForeColor.RGB = CLR_METIN_SOLUK()
        .MarginLeft = 0: .MarginTop = 0
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
    End With
End Sub

Public Sub OturumuKapat()
    modVeri.AyarYaz AYAR_OTURUM, 0
    GirisiGoster
End Sub

Public Function OturumVarMi() As Boolean
    OturumVarMi = (CLng(Val(modVeri.Ayar(AYAR_OTURUM, 0))) > 0)
End Function

Public Function OturumGerekli() As Boolean
    If Not OturumVarMi() Then
        MsgBox "Bu islem icin once oturum acmalisin.", vbExclamation, APP_ADI
        GirisiGoster
        Exit Function
    End If
    OturumGerekli = True
End Function
