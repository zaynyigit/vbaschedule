Attribute VB_Name = "modCizelge"
'===============================================================================
' modCizelge - Cizelgeyi ekrana cizer
'
' Tek giris noktasi: Ciz. Veri her degistiginde bu cagrilir; sayfa bastan
' cizilir. Boylece "ekranda kalan eski sekil" sorunu hic olusmaz.
'
' Cizim sirasi: hucre bicimleri -> proje/gorev satirlari -> sekiller.
' Sekiller (bar, gecikme seridi, halka, avatar) gx_ onekiyle adlandirilir ve
' her cizimde once hepsi silinir. UI_ onekli kalici butonlara dokunulmaz.
'===============================================================================
Option Explicit

Private Const AV_CAP As Long = 3          ' en fazla kac avatar gosterilsin

'===============================================================================
' Giris noktasi
'===============================================================================
Public Sub Ciz()
    Dim ws As Worksheet
    Dim yil As Long, ay As Long, gunSayisi As Long
    Dim hizli As Boolean

    On Error GoTo Hata
    hizli = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set ws = modVeri.Sayfa(SH_CIZELGE)
    yil = CLng(Val(modVeri.Ayar(AYAR_AKTIF_YIL, Year(Date))))
    ay = CLng(Val(modVeri.Ayar(AYAR_AKTIF_AY, Month(Date))))
    gunSayisi = Day(DateSerial(yil, ay + 1, 0))

    ws.Unprotect
    EskiSekilleriSil ws
    GovdeyiTemizle ws
    TakvimBasliginiCiz ws, yil, ay, gunSayisi
    SatirlariCiz ws, yil, ay, gunSayisi
    BugunCizgisiniCiz ws, yil, ay, gunSayisi
    AracCubuguKur ws
    EfsaneyiCiz ws
    BolmeyiDondur ws

    Application.ScreenUpdating = True
    Exit Sub
Hata:
    Application.ScreenUpdating = True
    MsgBox "Cizim hatasi: " & Err.Number & " - " & Err.Description, vbCritical, APP_ADI
End Sub

'===============================================================================
' Temizlik
'===============================================================================
Private Sub EskiSekilleriSil(ByVal ws As Worksheet)
    Dim i As Long, n As String
    For i = ws.Shapes.Count To 1 Step -1
        n = ws.Shapes(i).Name
        If Left$(n, 3) = "gx_" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub GovdeyiTemizle(ByVal ws As Worksheet)
    Dim sonSatir As Long
    sonSatir = ws.Cells(ws.Rows.Count, C_GOREV).End(xlUp).Row
    If sonSatir < R_VERI_BAS Then sonSatir = R_VERI_BAS
    With ws.Range(ws.Cells(R_AYADI, 1), ws.Cells(sonSatir + 60, C_GUN1 + GUN_SAYISI_MAX + 1))
        .ClearContents
        .Interior.Color = CLR_ZEMIN()
        .Borders.LineStyle = xlNone
        .UnMerge
    End With
    ws.Rows(R_VERI_BAS & ":" & (sonSatir + 60)).RowHeight = H_SATIR
End Sub

'===============================================================================
' Takvim basligi: buyuk ay adi + gun harfleri + gun numaralari
'===============================================================================
Private Sub TakvimBasliginiCiz(ByVal ws As Worksheet, ByVal yil As Long, _
                               ByVal ay As Long, ByVal gunSayisi As Long)
    Dim d As Long, c As Long, t As Date

    ' Kullanilmayan gun sutunlarini gizle (28/29/30 gunluk aylar icin)
    For d = 1 To GUN_SAYISI_MAX
        ws.Columns(C_GUN1 + d - 1).Hidden = (d > gunSayisi)
    Next d

    '--- Buyuk AY / YIL bandi ---
    With ws.Range(ws.Cells(R_AYADI, C_GUN1), ws.Cells(R_AYADI, C_GUN1 + gunSayisi - 1))
        .Merge
        .Value = AyAdi(ay) & "   " & yil
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = CLR_KOYU_BANT()
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = CLR_BEYAZ()
        .Font.Name = "Segoe UI Semibold"
    End With

    '--- Sol blok basliklari ---
    BaslikHucresi ws, R_AYADI, C_KATEGORI, "KATEGORI", R_GUNNO
    BaslikHucresi ws, R_AYADI, C_SIRA, "#", R_GUNNO
    BaslikHucresi ws, R_AYADI, C_SORUMLU, "SORUMLU", R_GUNNO
    BaslikHucresi ws, R_AYADI, C_GOREV, "GOREV", R_GUNNO
    BaslikHucresi ws, R_AYADI, C_ONCEKI, "ONCEKI AYDAN KALAN", R_GUNNO
    BaslikHucresi ws, R_AYADI, C_GUN1 + GUN_SAYISI_MAX, "ILERLEME", R_GUNNO
    BaslikHucresi ws, R_AYADI, C_GUN1 + GUN_SAYISI_MAX + 1, "ACIKLAMA", R_GUNNO

    '--- Gun harfleri ve numaralari ---
    For d = 1 To gunSayisi
        c = C_GUN1 + d - 1
        t = DateSerial(yil, ay, d)

        With ws.Cells(R_GUNHARF, c)
            .Value = GunHarfi(Weekday(t, vbMonday))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 7
            .Font.Bold = True
            .Font.Color = Hx("6B7588")
            .Interior.Color = CLR_BASLIK_ZEMIN()
        End With

        With ws.Cells(R_GUNNO, c)
            .Value = d
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = Hx("28324A")
            .Interior.Color = CLR_ZEMIN_ALT()
        End With

        ' Hafta sonu / tatil / bugun vurgusu
        If modVeri.OffGunMu(t) Then
            ws.Cells(R_GUNHARF, c).Interior.Color = CLR_TATIL()
            ws.Cells(R_GUNNO, c).Interior.Color = CLR_TATIL()
        ElseIf modVeri.HaftaSonuMu(t) Then
            ws.Cells(R_GUNHARF, c).Interior.Color = CLR_HAFTASONU()
            ws.Cells(R_GUNNO, c).Interior.Color = CLR_HAFTASONU()
        End If

        If t = Date Then
            With ws.Cells(R_GUNNO, c)
                .Interior.Color = Hx("FEE2E2")
                .Font.Color = Hx("B91C1C")
            End With
        End If

        ' Gun basligina cift tiklayinca tatil isaretlenir (bkz. sayfa olayi)
        ws.Cells(R_GUNNO, c).Locked = False
    Next d
End Sub

Private Sub BaslikHucresi(ByVal ws As Worksheet, ByVal r1 As Long, ByVal c As Long, _
                          ByVal metin As String, ByVal r2 As Long)
    With ws.Range(ws.Cells(r1, c), ws.Cells(r2, c))
        .Merge
        .Value = metin
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Font.Size = 7.5
        .Font.Bold = True
        .Font.Color = Hx("59637A")
        .Interior.Color = CLR_BASLIK_ZEMIN()
        .WrapText = False
    End With
End Sub

'===============================================================================
' Proje ve gorev satirlari
'===============================================================================
Private Sub SatirlariCiz(ByVal ws As Worksheet, ByVal yil As Long, _
                         ByVal ay As Long, ByVal gunSayisi As Long)
    Dim wsP As Worksheet, wsG As Worksheet
    Dim satir As Long, i As Long, j As Long
    Dim projeID As Long, p As TProje
    Dim gorevSatirlari As Collection, g As TGorev
    Dim projeGorevleri As Collection
    Dim toplamYuzde As Double, sayac As Long
    Dim ilkGun As Long, sonGun As Long, pIlk As Long, pSon As Long

    Set wsP = modVeri.Sayfa(SH_PROJELER)
    Set gorevSatirlari = modVeri.AyinGorevleri(yil, ay)
    satir = R_VERI_BAS

    For i = 2 To modVeri.SonSatir(wsP)
        projeID = CLng(Val(wsP.Cells(i, P_ID).Value))
        If projeID > 0 Then
            ' Bu projenin bu aya dusen gorevlerini topla
            Set projeGorevleri = New Collection
            For j = 1 To gorevSatirlari.Count
                g = modVeri.GorevOku(gorevSatirlari(j))
                If g.ProjeID = projeID Then projeGorevleri.Add gorevSatirlari(j)
            Next j

            If projeGorevleri.Count > 0 Then
                p = modVeri.ProjeGetir(projeID)

                ' --- proje grup satiri ---
                toplamYuzde = 0: pIlk = 32: pSon = 0
                For j = 1 To projeGorevleri.Count
                    g = modVeri.GorevOku(projeGorevleri(j))
                    toplamYuzde = toplamYuzde + g.Yuzde
                    AyIcindekiAralik g, yil, ay, gunSayisi, ilkGun, sonGun
                    If ilkGun > 0 Then
                        If ilkGun < pIlk Then pIlk = ilkGun
                        If sonGun > pSon Then pSon = sonGun
                    End If
                Next j

                ProjeSatiriCiz ws, satir, p, projeGorevleri.Count, gunSayisi, pIlk, pSon, _
                               toplamYuzde / projeGorevleri.Count
                satir = satir + 1

                ' --- gorev satirlari ---
                For j = 1 To projeGorevleri.Count
                    g = modVeri.GorevOku(projeGorevleri(j))
                    GorevSatiriCiz ws, satir, g, p, yil, ay, gunSayisi
                    satir = satir + 1
                Next j
            End If
        End If
    Next i

    ' Hicbir gorev yoksa bilgi ver
    If satir = R_VERI_BAS Then
        With ws.Cells(satir, C_GOREV)
            .Value = "Bu ayda gorev yok. Bir gun hucresine cift tiklayarak ekleyebilirsin."
            .Font.Italic = True
            .Font.Color = CLR_METIN_SOLUK()
        End With
    End If
End Sub

' Gorevin bu ay icindeki gun araligini bulur (1..gunSayisi). Ay disindaysa 0.
' Bitis olarak gecikme varsa gerceklesen/bugun alinir; bar + gecikme seridi
' birlikte bu araliga sigar.
Private Sub AyIcindekiAralik(ByRef g As TGorev, ByVal yil As Long, ByVal ay As Long, _
                             ByVal gunSayisi As Long, ByRef ilkGun As Long, ByRef sonGun As Long)
    Dim ayBas As Date, aySon As Date, b As Date, s As Date
    ayBas = DateSerial(yil, ay, 1)
    aySon = DateSerial(yil, ay, gunSayisi)

    b = g.PlanBaslangic
    s = g.PlanBitis
    If g.GerceklesenBitis > 0 Then
        If CDate(g.GerceklesenBitis) > s Then s = CDate(g.GerceklesenBitis)
    ElseIf g.Durum <> DURUM_TAMAM And Date > s Then
        s = Date
    End If

    If b > aySon Or s < ayBas Then
        ilkGun = 0: sonGun = 0
        Exit Sub
    End If
    If b < ayBas Then b = ayBas
    If s > aySon Then s = aySon
    ilkGun = Day(b)
    sonGun = Day(s)
End Sub

Private Sub ProjeSatiriCiz(ByVal ws As Worksheet, ByVal satir As Long, ByRef p As TProje, _
                           ByVal gorevSayisi As Long, ByVal gunSayisi As Long, _
                           ByVal ilkGun As Long, ByVal sonGun As Long, ByVal yuzde As Double)
    ws.Rows(satir).RowHeight = H_PROJE

    With ws.Range(ws.Cells(satir, C_SERIT), ws.Cells(satir, C_GUN1 + GUN_SAYISI_MAX + 1))
        .Interior.Color = Hx("F8FAFD")
    End With

    With ws.Range(ws.Cells(satir, C_KATEGORI), ws.Cells(satir, C_ONCEKI))
        .Merge
        .Value = "  " & UCase$(p.Ad) & "     " & gorevSayisi & " gorev"
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = Hx("111827")
        .VerticalAlignment = xlCenter
    End With

    ' Proje renk seridi
    ws.Cells(satir, C_SERIT).Interior.Color = Hx(p.Renk)

    ' Projenin sure bandi
    If ilkGun > 0 And sonGun >= ilkGun Then
        Dim sk As Shape, r1 As Range, r2 As Range
        Set r1 = ws.Cells(satir, C_GUN1 + ilkGun - 1)
        Set r2 = ws.Cells(satir, C_GUN1 + sonGun - 1)
        Set sk = ws.Shapes.AddShape(msoShapeRoundedRectangle, r1.Left + 1, _
                 r1.Top + r1.Height / 2 - 3, (r2.Left + r2.Width) - r1.Left - 2, 6)
        sk.Name = SP_PROJE & satir
        SekilBicimle sk, Hx(p.Renk), 0.62
        sk.Adjustments(1) = 0.5
    End If

    ' Proje ilerleme halkasi
    HalkaCiz ws, satir, C_GUN1 + GUN_SAYISI_MAX, yuzde, Hx(p.Renk), 15
End Sub

Private Sub GorevSatiriCiz(ByVal ws As Worksheet, ByVal satir As Long, ByRef g As TGorev, _
                           ByRef p As TProje, ByVal yil As Long, ByVal ay As Long, _
                           ByVal gunSayisi As Long)
    Dim kat As TKategori
    Dim d As Long, c As Long, t As Date
    Dim ilkGun As Long, sonGun As Long
    Dim planIlk As Long, planSon As Long
    Dim renk As Long

    ws.Rows(satir).RowHeight = H_SATIR
    kat = modVeri.KategoriGetir(g.KategoriID)
    renk = DurumRengi(g.Durum)

    ' --- sol blok ---
    ws.Cells(satir, C_SERIT).Interior.Color = Hx(p.Renk)

    With ws.Cells(satir, C_KATEGORI)
        .Value = kat.Ad
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 7.5
        .Font.Bold = True
        .Font.Color = CLR_BEYAZ()
        .Interior.Color = Hx(kat.Renk)
    End With

    With ws.Cells(satir, C_SIRA)
        .Value = g.Sira
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Color = CLR_METIN_SOLUK()
        .Font.Bold = True
    End With

    AvatarlariCiz ws, satir, g

    With ws.Cells(satir, C_GOREV)
        .Value = g.Ad
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Font.Bold = True
        .Font.Size = 9
    End With

    With ws.Cells(satir, C_ONCEKI)
        .Value = IIf(Len(g.OncekiAyNot) = 0, ChrW(8212), g.OncekiAyNot)
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Font.Size = 8
        .Font.Color = CLR_METIN_SOLUK()
    End With

    ' --- gun hucrelerinin zemini ---
    For d = 1 To gunSayisi
        c = C_GUN1 + d - 1
        t = DateSerial(yil, ay, d)
        If modVeri.OffGunMu(t) Then
            ws.Cells(satir, c).Interior.Color = Hx("E6EAF1")
        ElseIf modVeri.HaftaSonuMu(t) Then
            ws.Cells(satir, c).Interior.Color = Hx("F3F6FA")
        ElseIf t = Date Then
            ws.Cells(satir, c).Interior.Color = CLR_BUGUN_ZEMIN()
        End If
        ws.Cells(satir, c).Locked = False       ' cift tiklamayla gorev eklenebilsin
    Next d

    ' --- planlanan bar ---
    GunAraligi g.PlanBaslangic, g.PlanBitis, yil, ay, gunSayisi, planIlk, planSon
    If planIlk > 0 Then
        BarCiz ws, satir, planIlk, planSon, renk, g, (g.Durum = DURUM_PLAN)
    End If

    ' --- gecikme seridi ---
    GecikmeSeridiCiz ws, satir, g, yil, ay, gunSayisi, planSon

    ' --- sag blok ---
    HalkaCiz ws, satir, C_GUN1 + GUN_SAYISI_MAX, g.Yuzde, renk, 19

    With ws.Cells(satir, C_GUN1 + GUN_SAYISI_MAX + 1)
        .Value = g.Aciklama
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Font.Size = 8
        .Font.Color = Hx("5D677A")
        .Interior.Color = Hx("FBFCFE")
    End With
End Sub

' Verilen tarih araliginin ay icindeki gun karsiligi (0 = ay disinda)
Private Sub GunAraligi(ByVal bas As Date, ByVal bit As Date, ByVal yil As Long, _
                       ByVal ay As Long, ByVal gunSayisi As Long, _
                       ByRef ilkGun As Long, ByRef sonGun As Long)
    Dim ayBas As Date, aySon As Date
    ayBas = DateSerial(yil, ay, 1)
    aySon = DateSerial(yil, ay, gunSayisi)
    If bas > aySon Or bit < ayBas Then
        ilkGun = 0: sonGun = 0
        Exit Sub
    End If
    If bas < ayBas Then bas = ayBas
    If bit > aySon Then bit = aySon
    ilkGun = Day(bas)
    sonGun = Day(bit)
End Sub

'===============================================================================
' Sekiller: bar, gecikme, halka, avatar
'===============================================================================
Private Sub BarCiz(ByVal ws As Worksheet, ByVal satir As Long, ByVal ilkGun As Long, _
                   ByVal sonGun As Long, ByVal renk As Long, ByRef g As TGorev, _
                   ByVal planlanan As Boolean)
    Dim r1 As Range, r2 As Range, sk As Shape
    Dim sol As Double, gen As Double, yuk As Double, ust As Double

    Set r1 = ws.Cells(satir, C_GUN1 + ilkGun - 1)
    Set r2 = ws.Cells(satir, C_GUN1 + sonGun - 1)
    yuk = 13
    sol = r1.Left + 1.5
    gen = (r2.Left + r2.Width) - r1.Left - 3
    ust = r1.Top + (r1.Height - yuk) / 2
    If gen < 4 Then gen = 4

    Set sk = ws.Shapes.AddShape(msoShapeRoundedRectangle, sol, ust, gen, yuk)
    sk.Name = SP_BAR & g.ID
    sk.Adjustments(1) = 0.5

    If planlanan Then
        ' Planlanan: saydam grimsi govde, iki ucta kisa gri cizgi
        With sk.Fill
            .ForeColor.RGB = CLR_PLAN()
            .Transparency = 0.78
        End With
        With sk.Line
            .ForeColor.RGB = CLR_PLAN()
            .Weight = 1
            .Transparency = 0.35
        End With
        UcCizgisiCiz ws, g.ID & "a", sol, ust, yuk
        UcCizgisiCiz ws, g.ID & "b", sol + gen - 2, ust, yuk
    Else
        SekilBicimle sk, renk, 0
    End If

    ' Cubuk yazisi - ancak sigacaksa
    If Len(g.CubukYazisi) > 0 And gen > 40 Then
        With sk.TextFrame2
            .TextRange.Text = g.CubukYazisi
            .TextRange.Font.Size = 7.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Name = "Segoe UI"
            If planlanan Then
                .TextRange.Font.Fill.ForeColor.RGB = Hx("64748B")
            Else
                .TextRange.Font.Fill.ForeColor.RGB = CLR_BEYAZ()
            End If
            .MarginLeft = 3: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorNone
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .WordWrap = msoFalse
            .AutoSize = msoAutoSizeNone
        End With
    End If

    sk.OnAction = "modYorum.BaraTiklandi"
    sk.AlternativeText = CStr(g.ID)

    ' Yorum rozeti
    Dim n As Long
    n = modVeri.YorumSayisi(g.ID)
    If n > 0 Then RozetCiz ws, g.ID, sol + gen - 9, ust - 4, n
End Sub

Private Sub UcCizgisiCiz(ByVal ws As Worksheet, ByVal etiket As String, ByVal sol As Double, _
                         ByVal ust As Double, ByVal yuk As Double)
    Dim sk As Shape
    Set sk = ws.Shapes.AddShape(msoShapeRectangle, sol, ust - 2, 2, yuk + 4)
    sk.Name = SP_BAR & "cap_" & etiket
    SekilBicimle sk, CLR_PLAN(), 0
End Sub

Private Sub GecikmeSeridiCiz(ByVal ws As Worksheet, ByVal satir As Long, ByRef g As TGorev, _
                             ByVal yil As Long, ByVal ay As Long, ByVal gunSayisi As Long, _
                             ByVal planSon As Long)
    Dim gecikme As Long, bitis As Date
    Dim gIlk As Long, gSon As Long
    Dim r1 As Range, r2 As Range, sk As Shape
    Dim sol As Double, sag As Double, orta As Double

    gecikme = modVeri.GecikmeGunu(g)
    If gecikme <= 0 Then Exit Sub

    If g.GerceklesenBitis > 0 Then bitis = CDate(g.GerceklesenBitis) Else bitis = Date

    ' Serit planlanan bitisin ERTESI gununden gerceklesen bitise kadar
    GunAraligi DateAdd("d", 1, g.PlanBitis), bitis, yil, ay, gunSayisi, gIlk, gSon
    If gIlk = 0 Then Exit Sub

    Set r1 = ws.Cells(satir, C_GUN1 + gIlk - 1)
    Set r2 = ws.Cells(satir, C_GUN1 + gSon - 1)
    sol = r1.Left
    sag = r2.Left + r2.Width - 2
    orta = r1.Top + r1.Height / 2

    ' Kirmizi kesikli cizgi
    Set sk = ws.Shapes.AddLine(sol, orta, sag, orta)
    sk.Name = SP_GECIKME & g.ID
    With sk.Line
        .ForeColor.RGB = CLR_GECIKME()
        .Weight = 2.25
        .DashStyle = msoLineDashDot
    End With
    sk.OnAction = "modGecikme.SerideTiklandi"
    sk.AlternativeText = CStr(g.ID)

    ' Ucta dik kirmizi kapak
    Set sk = ws.Shapes.AddShape(msoShapeRectangle, sag - 1, orta - 7, 2.2, 14)
    sk.Name = SP_GECIKME & "cap_" & g.ID
    SekilBicimle sk, CLR_GECIKME(), 0
    sk.OnAction = "modGecikme.SerideTiklandi"
    sk.AlternativeText = CStr(g.ID)

    ' Planlanan bitisi isaretleyen cemberli nokta
    If planSon > 0 Then
        Dim rp As Range
        Set rp = ws.Cells(satir, C_GUN1 + planSon - 1)
        Set sk = ws.Shapes.AddShape(msoShapeOval, rp.Left + rp.Width - 6, orta - 4.5, 9, 9)
        sk.Name = SP_GECIKME & "dot_" & g.ID
        sk.Fill.ForeColor.RGB = CLR_BEYAZ()
        With sk.Line
            .ForeColor.RGB = CLR_GECIKME()
            .Weight = 2
        End With
        sk.Shadow.Visible = msoFalse
    End If

    ' "+N g" etiketi
    Set sk = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, sag + 2, orta - 7, 30, 12)
    sk.Name = SP_GECIKME & "lbl_" & g.ID
    With sk.TextFrame2
        .TextRange.Text = "+" & gecikme & "g"
        .TextRange.Font.Size = 7
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = CLR_GECIKME()
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
    End With
    sk.Fill.Visible = msoFalse
    sk.Line.Visible = msoFalse
End Sub

' Halka: arka plan cemberi + ilerleme dilimi + ortasi beyaz oyuk + yuzde yazisi
Private Sub HalkaCiz(ByVal ws As Worksheet, ByVal satir As Long, ByVal sutun As Long, _
                     ByVal yuzde As Double, ByVal renk As Long, ByVal cap As Double)
    Dim r As Range, sk As Shape
    Dim sol As Double, ust As Double, ic As Double

    Set r = ws.Cells(satir, sutun)
    sol = r.Left + (r.Width - cap) / 2 - 8
    ust = r.Top + (r.Height - cap) / 2
    If sol < r.Left Then sol = r.Left + 2

    ' 1) arka plan cemberi
    Set sk = ws.Shapes.AddShape(msoShapeOval, sol, ust, cap, cap)
    sk.Name = SP_HALKA & "bg_" & satir
    SekilBicimle sk, Hx("E8EDF4"), 0

    ' 2) ilerleme dilimi
    If yuzde >= 99.5 Then
        Set sk = ws.Shapes.AddShape(msoShapeOval, sol, ust, cap, cap)
        sk.Name = SP_HALKA & "fg_" & satir
        SekilBicimle sk, renk, 0
    ElseIf yuzde > 0.5 Then
        Set sk = ws.Shapes.AddShape(msoShapePie, sol, ust, cap, cap)
        sk.Name = SP_HALKA & "fg_" & satir
        SekilBicimle sk, renk, 0
        ' Pie acilari saat 3 yonunden saat yonunde derece. 270 = saat 12.
        On Error Resume Next
        sk.Adjustments(1) = 270
        sk.Adjustments(2) = NormalizeAci(270 + 3.6 * yuzde)
        On Error GoTo 0
    End If

    ' 3) ortayi oy
    ic = cap * 0.56
    Set sk = ws.Shapes.AddShape(msoShapeOval, sol + (cap - ic) / 2, ust + (cap - ic) / 2, ic, ic)
    sk.Name = SP_HALKA & "in_" & satir
    SekilBicimle sk, CLR_ZEMIN(), 0

    ' 4) yuzde yazisi - halkanin altina
    With ws.Cells(satir, sutun)
        .Value = "%" & Format$(yuzde, "0")
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Font.Size = 8.5
        .Font.Bold = True
        .Font.Color = Hx("28324A")
        .Interior.Color = Hx("FBFCFE")
    End With
End Sub

Private Function NormalizeAci(ByVal a As Double) As Double
    Do While a >= 360
        a = a - 360
    Loop
    Do While a < 0
        a = a + 360
    Loop
    NormalizeAci = a
End Function

Private Sub AvatarlariCiz(ByVal ws As Worksheet, ByVal satir As Long, ByRef g As TGorev)
    Dim ids As Variant, i As Long, n As Long
    Dim r As Range, sk As Shape, u As TKullanici
    Dim cap As Double, sol As Double, ust As Double, adim As Double

    If Len(Trim$(g.Sorumlular)) = 0 Then Exit Sub
    ids = Split(g.Sorumlular, ",")
    n = UBound(ids) - LBound(ids) + 1

    Set r = ws.Cells(satir, C_SORUMLU)
    cap = 15: adim = 11
    ust = r.Top + (r.Height - cap) / 2
    sol = r.Left + 4

    For i = LBound(ids) To UBound(ids)
        If i - LBound(ids) >= AV_CAP Then Exit For
        u = modVeri.KullaniciGetir(CLng(Val(ids(i))))
        If u.ID > 0 Then
            Set sk = ws.Shapes.AddShape(msoShapeOval, sol, ust, cap, cap)
            sk.Name = SP_AVATAR & satir & "_" & i
            SekilBicimle sk, Hx(IIf(Len(u.Renk) = 6, u.Renk, "64748B")), 0
            sk.Line.Visible = msoTrue
            sk.Line.ForeColor.RGB = CLR_BEYAZ()
            sk.Line.Weight = 1.25
            AvatarYazisi sk, BasHarfler(u.Ad), CLR_BEYAZ()
            sol = sol + adim
        End If
    Next i

    If n > AV_CAP Then
        Set sk = ws.Shapes.AddShape(msoShapeOval, sol, ust, cap, cap)
        sk.Name = SP_AVATAR & satir & "_more"
        SekilBicimle sk, Hx("E2E8F0"), 0
        sk.Line.Visible = msoTrue
        sk.Line.ForeColor.RGB = CLR_BEYAZ()
        sk.Line.Weight = 1.25
        AvatarYazisi sk, "+" & (n - AV_CAP), Hx("475569")
    End If

    ' Avatarlarin uzerine gelince sorumlu listesi cikssin.
    ' Var olan aciklamayi once sil: AddComment zaten aciklamasi olan hucrede
    ' hata verir ve cizim her yenilendiginde buraya tekrar gelinir.
    Dim hc As Range
    Set hc = ws.Cells(satir, C_SORUMLU)
    On Error Resume Next
    If Not hc.Comment Is Nothing Then hc.Comment.Delete
    On Error GoTo 0
    hc.AddComment
    hc.Comment.Text Text:=SorumluOzeti(g.Sorumlular)
    hc.Comment.Visible = False
    hc.Comment.Shape.TextFrame.AutoSize = True
End Sub

Private Sub AvatarYazisi(ByVal sk As Shape, ByVal metin As String, ByVal renk As Long)
    With sk.TextFrame2
        .TextRange.Text = metin
        .TextRange.Font.Size = 6.5
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Fill.ForeColor.RGB = renk
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
    End With
End Sub

Public Function BasHarfler(ByVal ad As String) As String
    Dim p As Variant, s As String
    p = Split(Trim$(ad), " ")
    If UBound(p) >= 1 Then
        s = UCase$(Left$(p(0), 1) & Left$(p(UBound(p)), 1))
    Else
        s = UCase$(Left$(Trim$(ad), 2))
    End If
    BasHarfler = s
End Function

Public Function SorumluOzeti(ByVal idListesi As String) As String
    Dim ids As Variant, i As Long, u As TKullanici, s As String
    If Len(Trim$(idListesi)) = 0 Then SorumluOzeti = "Sorumlu atanmamis": Exit Function
    ids = Split(idListesi, ",")
    s = "SORUMLULAR (" & (UBound(ids) + 1) & " kisi)" & vbLf
    For i = LBound(ids) To UBound(ids)
        u = modVeri.KullaniciGetir(CLng(Val(ids(i))))
        If u.ID > 0 Then s = s & vbLf & ChrW(8226) & " " & u.Ad & _
            IIf(u.Rol = ROL_ADMIN, "  (admin)", "")
    Next i
    SorumluOzeti = s
End Function

Private Sub RozetCiz(ByVal ws As Worksheet, ByVal gorevID As Long, ByVal sol As Double, _
                     ByVal ust As Double, ByVal adet As Long)
    Dim sk As Shape
    Set sk = ws.Shapes.AddShape(msoShapeOval, sol, ust, 12, 12)
    sk.Name = SP_ROZET & gorevID
    SekilBicimle sk, Hx("0F172A"), 0
    AvatarYazisi sk, CStr(adet), CLR_BEYAZ()
    sk.OnAction = "modYorum.BaraTiklandi"
    sk.AlternativeText = CStr(gorevID)
End Sub

'===============================================================================
' Bugun cizgisi
'===============================================================================
Private Sub BugunCizgisiniCiz(ByVal ws As Worksheet, ByVal yil As Long, _
                              ByVal ay As Long, ByVal gunSayisi As Long)
    Dim bugun As Date, c As Long, sonSatir As Long
    Dim r1 As Range, r2 As Range, sk As Shape, x As Double

    bugun = Date
    If Year(bugun) <> yil Or Month(bugun) <> ay Then Exit Sub

    c = C_GUN1 + Day(bugun) - 1
    sonSatir = ws.Cells(ws.Rows.Count, C_GOREV).End(xlUp).Row
    If sonSatir < R_VERI_BAS Then sonSatir = R_VERI_BAS + 1

    Set r1 = ws.Cells(R_GUNHARF, c)
    Set r2 = ws.Cells(sonSatir, c)
    x = r1.Left + r1.Width / 2

    Set sk = ws.Shapes.AddLine(x, r1.Top, x, r2.Top + r2.Height)
    sk.Name = SP_BUGUN
    With sk.Line
        .ForeColor.RGB = CLR_BUGUN_CIZGI()
        .Weight = 1.75
    End With
    sk.ZOrder msoBringToFront
End Sub

'===============================================================================
' Arac cubugu, efsane, bolme
'===============================================================================
Private Sub AracCubuguKur(ByVal ws As Worksheet)
    ' Butonlar UI_ onekli ve kalicidir; sadece bir kez olusturulur.
    If SekilVarMi(ws, SP_UI & "onceki") Then Exit Sub

    Dim x As Double, y As Double
    x = ws.Cells(R_BASLIK, C_GUN1).Left
    y = ws.Cells(R_BASLIK, C_GUN1).Top + 10

    x = ButonEkle(ws, "onceki", ChrW(8249) & " Onceki ay", x, y, 78, "modGezinme.OncekiAy")
    x = ButonEkle(ws, "bugun", "Bugun", x, y, 48, "modGezinme.BugunAyi")
    x = ButonEkle(ws, "sonraki", "Sonraki ay " & ChrW(8250), x, y, 78, "modGezinme.SonrakiAy")
    x = x + 10
    x = ButonEkle(ws, "gorevekle", "+ Gorev", x, y, 62, "modGorev.GorevEkleAc")
    x = ButonEkle(ws, "kullanici", "Kullanicilar", x, y, 72, "modKullanici.KullaniciPaneliAc")
    x = ButonEkle(ws, "kategori", "Kategoriler", x, y, 72, "modKullanici.KategoriPaneliAc")
    x = ButonEkle(ws, "pdf", "Yazdir / PDF", x, y, 80, "modBaski.BaskiPaneliAc")
    x = ButonEkle(ws, "cikis", "Cikis", x, y, 46, "modGiris.OturumuKapat")
End Sub

Private Function ButonEkle(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, _
                           ByVal x As Double, ByVal y As Double, ByVal gen As Double, _
                           ByVal makro As String) As Double
    Dim sk As Shape
    Set sk = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, gen, 22)
    sk.Name = SP_UI & ad
    sk.Adjustments(1) = 0.22
    SekilBicimle sk, Hx("F2F5FA"), 0
    sk.Line.Visible = msoTrue
    sk.Line.ForeColor.RGB = Hx("DFE5EE")
    sk.Line.Weight = 0.75
    With sk.TextFrame2
        .TextRange.Text = metin
        .TextRange.Font.Size = 8
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Fill.ForeColor.RGB = Hx("47536B")
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
    End With
    sk.OnAction = makro
    ButonEkle = x + gen + 5
End Function

Private Sub EfsaneyiCiz(ByVal ws As Worksheet)
    Dim ogeler As Variant, i As Long, c As Long
    ogeler = Array(Array("Tamamlandi", "16A34A"), Array("Devam ediyor", "F59E0B"), _
                   Array("Planlandi", "94A3B8"), Array("Hafta sonu", "E9EDF4"), _
                   Array("Idari tatil", "D9DFE9"), Array("Bugun", "DC2626"))
    c = C_GUN1
    For i = LBound(ogeler) To UBound(ogeler)
        With ws.Cells(R_EFSANE, c)
            .Value = ChrW(9632) & " " & ogeler(i)(0)
            .Font.Size = 7.5
            .Font.Bold = True
            .Font.Color = Hx(ogeler(i)(1))
            .HorizontalAlignment = xlLeft
        End With
        c = c + 5
    Next i
End Sub

Private Sub BolmeyiDondur(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(R_VERI_BAS, C_GUN1).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(R_VERI_BAS, C_GUN1).Select
    On Error GoTo 0
End Sub

'===============================================================================
' Ortak sekil yardimcilari
'===============================================================================
Public Sub SekilBicimle(ByVal sk As Shape, ByVal renk As Long, ByVal saydamlik As Double)
    With sk.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = renk
        .Transparency = saydamlik
    End With
    sk.Line.Visible = msoFalse
    sk.Shadow.Visible = msoFalse
End Sub

Public Function SekilVarMi(ByVal ws As Worksheet, ByVal ad As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = ws.Shapes(ad)
    On Error GoTo 0
    SekilVarMi = Not (s Is Nothing)
End Function
