Attribute VB_Name = "modBaski"
'===============================================================================
' modBaski - Sayfa duzeni, A3/A1 sigdirma ve PDF cikti
'
' A1/A2 HAKKINDA DURUST NOT:
' Excel'in hazir kagit sabitleri A4 ve A3'u kapsar. A2 ve A1 icin kagit
' boyutunu YAZICI SURUCUSU sunmak zorundadir; sistemde A1 destekleyen bir
' surucu (plotter ya da ozel boyutlu PDF yazicisi) yoksa Excel bu boyuta
' gecemez. Kod once dogrudan denemeyi yapar, olmazsa seni uyarir ve
' "A3'e sigdir, matbaada buyut" yoluna duser. Bu Excel'in siniri, kodun degil.
'===============================================================================
Option Explicit

Private Const B_KAGIT = 1, B_YON = 2, B_SIGDIR = 3, B_BASLIKTEKRAR = 4
Private Const B_RENKLI = 5, B_YORUMLAR = 6, B_GECIKMELER = 7

'===============================================================================
' Panel
'===============================================================================
Public Sub BaskiPaneliAc()
    If Not modGiris.OturumGerekli() Then Exit Sub

    modModal.Temizle
    modModal.Alan "Kagit boyutu", MT_LISTE, "A3", "A4,A3,A2,A1"
    modModal.Alan "Yon", MT_LISTE, "Yatay", "Yatay,Dikey"
    modModal.Alan "Sigdirma", MT_LISTE, "1 sayfaya sigdir", _
                  "1 sayfaya sigdir,Genisligi 1 sayfa,Olcekleme yok"
    modModal.Alan "Baslik satirlari her sayfada tekrarlansin", MT_ONAY, "Evet"
    modModal.Alan "Renkli bas", MT_ONAY, "Evet"
    modModal.Alan "Yorumlari ek sayfa olarak bas", MT_ONAY, "Hayir"
    modModal.Alan "Gecikme nedenlerini ek sayfa olarak bas", MT_ONAY, "Evet"

    modModal.Goster "baski", "Yazdir / PDF'e aktar", "PDF olarak kaydet"
End Sub

'===============================================================================
' Uygulama
'===============================================================================
Public Function BaskiUygula() As Boolean
    Dim ws As Worksheet
    Dim kagit As String, yon As String, sigdir As String
    Dim uyari As String, dosya As String

    Set ws = modVeri.Sayfa(SH_CIZELGE)
    kagit = modModal.Deger(B_KAGIT)
    yon = modModal.Deger(B_YON)
    sigdir = modModal.Deger(B_SIGDIR)

    Application.ScreenUpdating = False
    SayfaDuzeniniKur ws, kagit, yon, sigdir, _
                     modModal.DegerOnay(B_BASLIKTEKRAR), _
                     modModal.DegerOnay(B_RENKLI), uyari

    If modModal.DegerOnay(B_GECIKMELER) Then EkSayfaKur "Gecikmeler", GecikmeRaporu()
    If modModal.DegerOnay(B_YORUMLAR) Then EkSayfaKur "Yorumlar", YorumRaporu()

    Application.ScreenUpdating = True

    If Len(uyari) > 0 Then
        MsgBox uyari, vbExclamation, APP_ADI
    End If

    dosya = PdfYoluSor()
    If Len(dosya) = 0 Then
        BaskiUygula = True       ' kullanici vazgecti, panel yine de kapansin
        Exit Function
    End If

    On Error GoTo Hata
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=dosya, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, _
                           IgnorePrintAreas:=False, OpenAfterPublish:=True
    MsgBox "PDF olusturuldu:" & vbCrLf & dosya, vbInformation, APP_ADI
    BaskiUygula = True
    Exit Function

Hata:
    MsgBox "PDF olusturulamadi: " & Err.Description, vbCritical, APP_ADI
    BaskiUygula = True
End Function

'===============================================================================
' Sayfa duzeni
'===============================================================================
Public Sub SayfaDuzeniniKur(ByVal ws As Worksheet, ByVal kagit As String, _
                            ByVal yon As String, ByVal sigdir As String, _
                            ByVal baslikTekrar As Boolean, ByVal renkli As Boolean, _
                            ByRef uyari As String)
    Dim sonSatir As Long, sonSutun As Long

    sonSatir = ws.Cells(ws.Rows.Count, C_GOREV).End(xlUp).Row
    If sonSatir < R_VERI_BAS Then sonSatir = R_VERI_BAS
    sonSutun = C_GUN1 + GUN_SAYISI_MAX + 1

    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(R_AYADI, 1), ws.Cells(sonSatir, sonSutun)).Address
        .Orientation = IIf(yon = "Dikey", xlPortrait, xlLandscape)
        .LeftMargin = Application.CentimetersToPoints(0.6)
        .RightMargin = Application.CentimetersToPoints(0.6)
        .TopMargin = Application.CentimetersToPoints(0.8)
        .BottomMargin = Application.CentimetersToPoints(0.8)
        .HeaderMargin = Application.CentimetersToPoints(0.3)
        .FooterMargin = Application.CentimetersToPoints(0.3)
        .CenterHorizontally = True
        .BlackAndWhite = Not renkli

        If baslikTekrar Then
            .PrintTitleRows = ws.Rows(R_AYADI).Address & ":" & ws.Rows(R_GUNNO).Address
            .PrintTitleColumns = ws.Columns(C_SERIT).Address & ":" & ws.Columns(C_ONCEKI).Address
        Else
            .PrintTitleRows = ""
            .PrintTitleColumns = ""
        End If

        .CenterFooter = APP_ADI & "  -  " & AyAdi(CLng(Val(modVeri.Ayar(AYAR_AKTIF_AY, Month(Date))))) & _
                        " " & modVeri.Ayar(AYAR_AKTIF_YIL, Year(Date)) & _
                        "   |   " & APP_TELIF & "   |   Sayfa &P / &N"
        .RightHeader = "Cikti: " & Format$(Now, "dd.mm.yyyy hh:nn")

        Select Case sigdir
            Case "1 sayfaya sigdir"
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
            Case "Genisligi 1 sayfa"
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
            Case Else
                .Zoom = 100
        End Select
    End With
    Application.PrintCommunication = True

    KagitBoyutuAyarla ws, kagit, uyari
End Sub

' A4/A3 hazir sabitlerle; A2/A1 Windows kagit koduyla denenir.
Private Sub KagitBoyutuAyarla(ByVal ws As Worksheet, ByVal kagit As String, ByRef uyari As String)
    Dim kod As Long, tamam As Boolean

    Select Case UCase$(kagit)
        Case "A4": kod = xlPaperA4      ' 9
        Case "A3": kod = xlPaperA3      ' 8
        Case "A2": kod = 66             ' DMPAPER_A2
        Case "A1": kod = 67             ' bazi suruculerde A1
        Case Else: kod = xlPaperA3
    End Select

    Application.PrintCommunication = False
    On Error Resume Next
    Err.Clear
    ws.PageSetup.PaperSize = kod
    tamam = (Err.Number = 0)
    On Error GoTo 0
    Application.PrintCommunication = True

    ' Gercekten uygulandi mi diye geri oku
    If tamam Then
        On Error Resume Next
        tamam = (ws.PageSetup.PaperSize = kod)
        On Error GoTo 0
    End If

    If Not tamam Then
        Application.PrintCommunication = False
        On Error Resume Next
        ws.PageSetup.PaperSize = xlPaperA3
        On Error GoTo 0
        Application.PrintCommunication = True

        uyari = kagit & " kagit boyutu bu bilgisayardaki yazici surucusunde yok." & vbCrLf & vbCrLf & _
                "Cizelge A3 yatay olarak tek sayfaya sigdirildi. " & kagit & " cikti icin iki secenek var:" & vbCrLf & _
                "  1) A1 destekleyen bir yazici/plotter surucusu kur, sonra tekrar dene." & vbCrLf & _
                "  2) Bu A3 PDF'i matbaada " & kagit & "'e buyutt. Vektorel oldugu icin kalite kaybi olmaz."
    End If
End Sub

Private Function PdfYoluSor() As String
    Dim d As FileDialog, varsayilan As String
    varsayilan = "Cizelge_" & modVeri.Ayar(AYAR_AKTIF_YIL, Year(Date)) & "_" & _
                 Format$(CLng(Val(modVeri.Ayar(AYAR_AKTIF_AY, Month(Date)))), "00") & ".pdf"

    On Error GoTo Basit
    Set d = Application.FileDialog(msoFileDialogSaveAs)
    d.Title = "PDF'i kaydet"
    d.InitialFileName = varsayilan
    If d.Show = -1 Then
        PdfYoluSor = d.SelectedItems(1)
        If LCase$(Right$(PdfYoluSor, 4)) <> ".pdf" Then PdfYoluSor = PdfYoluSor & ".pdf"
    End If
    Exit Function

Basit:
    PdfYoluSor = Application.GetSaveAsFilename(varsayilan, "PDF Dosyasi (*.pdf), *.pdf")
    If PdfYoluSor = "False" Then PdfYoluSor = ""
End Function

'===============================================================================
' Ek rapor sayfalari
'===============================================================================
Private Sub EkSayfaKur(ByVal ad As String, ByVal icerik As String)
    Dim ws As Worksheet, satirlar As Variant, i As Long

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("_Rapor" & ad).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add( _
             After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "_Rapor" & ad
    ws.Cells.Font.Name = "Segoe UI"
    ws.Columns(1).ColumnWidth = 3
    ws.Columns(2).ColumnWidth = 110

    With ws.Cells(2, 2)
        .Value = UCase$(ad) & " - " & AyAdi(CLng(Val(modVeri.Ayar(AYAR_AKTIF_AY, Month(Date))))) & _
                 " " & modVeri.Ayar(AYAR_AKTIF_YIL, Year(Date))
        .Font.Size = 14
        .Font.Bold = True
    End With

    satirlar = Split(icerik, vbLf)
    For i = LBound(satirlar) To UBound(satirlar)
        With ws.Cells(4 + i, 2)
            .Value = satirlar(i)
            .Font.Size = 9
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    Next i

    ws.PageSetup.Orientation = xlPortrait
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    ws.Visible = xlSheetVisible
End Sub

Private Function GecikmeRaporu() As String
    Dim ws As Worksheet, i As Long, s As String
    Dim g As TGorev, u As TKullanici, gecikme As Long

    Set ws = modVeri.Sayfa(SH_GECIKMELER)
    For i = 2 To modVeri.SonSatir(ws)
        If ws.Cells(i, D_GOREV).Value <> "" Then
            g = modVeri.GorevGetir(CLng(Val(ws.Cells(i, D_GOREV).Value)))
            If g.ID > 0 Then
                gecikme = modVeri.GecikmeGunu(g)
                u = modVeri.KullaniciGetir(CLng(Val(ws.Cells(i, D_KAYDEDEN).Value)))
                s = s & IIf(Len(s) > 0, vbLf & vbLf, "") & _
                    g.Ad & "   (+" & gecikme & " gun)" & vbLf & _
                    "Planlanan: " & Format$(g.PlanBitis, "dd.mm.yyyy") & _
                    IIf(g.GerceklesenBitis > 0, _
                        "   Gerceklesen: " & Format$(CDate(g.GerceklesenBitis), "dd.mm.yyyy"), _
                        "   (devam ediyor)") & vbLf & _
                    "Sorumluluk: " & ws.Cells(i, D_SORUMLULUK).Value & vbLf & _
                    "Neden: " & ws.Cells(i, D_NEDEN).Value & vbLf & _
                    "Kaydeden: " & u.Ad & "  " & Format$(ws.Cells(i, D_TARIH).Value, "dd.mm.yyyy hh:nn")
            End If
        End If
    Next i
    If Len(s) = 0 Then s = "Bu donemde kayitli gecikme yok."
    GecikmeRaporu = s
End Function

Private Function YorumRaporu() As String
    Dim wsG As Worksheet, i As Long, s As String, g As TGorev
    Set wsG = modVeri.Sayfa(SH_GOREVLER)
    For i = 2 To modVeri.SonSatir(wsG)
        If wsG.Cells(i, G_ID).Value <> "" Then
            g = modVeri.GorevOku(i)
            If modVeri.YorumSayisi(g.ID) > 0 Then
                s = s & IIf(Len(s) > 0, vbLf & vbLf, "") & _
                    g.Ad & vbLf & modYorum.YorumGecmisi(g.ID)
            End If
        End If
    Next i
    If Len(s) = 0 Then s = "Kayitli yorum yok."
    YorumRaporu = s
End Function
