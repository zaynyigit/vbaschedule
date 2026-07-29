Attribute VB_Name = "modVeri"
'===============================================================================
' modVeri - Veri katmani
'
' Tum okuma/yazma buradan gecer. Arayuz kodu asla gizli sayfalara dogrudan
' dokunmaz. Ileride veri baska yere tasinirsa (ayri dosya, SQL) sadece bu
' modul degisir, cizim kodu ellenmez.
'
' Her veri sayfasinin 1. satiri basliktir, veri 2. satirdan baslar.
'===============================================================================
Option Explicit

'--- _Kullanicilar sutunlari ---------------------------------------------------
Public Const K_ID = 1, K_AD = 2, K_KADI = 3, K_PAROLA = 4, K_ROL = 5
Public Const K_FOTO = 6, K_Y_GOREV = 7, K_Y_YORUM = 8, K_Y_GECIKME = 9
Public Const K_Y_KULLANICI = 10, K_SAYFA = 11, K_AKTIF = 12, K_RENK = 13

'--- _Projeler -----------------------------------------------------------------
Public Const P_ID = 1, P_AD = 2, P_RENK = 3, P_SIRA = 4, P_AKTIF = 5

'--- _Kategoriler --------------------------------------------------------------
Public Const T_ID = 1, T_AD = 2, T_RENK = 3, T_SIRA = 4, T_AKTIF = 5

'--- _Gorevler -----------------------------------------------------------------
Public Const G_ID = 1, G_PROJE = 2, G_KATEGORI = 3, G_SIRA = 4, G_AD = 5
Public Const G_SORUMLU = 6, G_ONCEKI = 7, G_PLANBAS = 8, G_PLANBIT = 9
Public Const G_GERCEK = 10, G_DURUM = 11, G_CUBUK = 12, G_YUZDE = 13
Public Const G_ACIKLAMA = 14, G_OLUSTURAN = 15, G_TARIH = 16

'--- _Yorumlar -----------------------------------------------------------------
Public Const Y_ID = 1, Y_GOREV = 2, Y_KULLANICI = 3, Y_TARIH = 4, Y_METIN = 5

'--- _Gecikmeler ---------------------------------------------------------------
Public Const D_ID = 1, D_GOREV = 2, D_NEDEN = 3, D_SORUMLULUK = 4
Public Const D_KAYDEDEN = 5, D_TARIH = 6

'--- _OffGunler ----------------------------------------------------------------
Public Const O_TARIH = 1, O_ACIKLAMA = 2

'===============================================================================
' Tipler
'===============================================================================
Public Type TGorev
    ID              As Long
    ProjeID         As Long
    KategoriID      As Long
    Sira            As Long
    Ad              As String
    Sorumlular      As String      ' virgulle ayrilmis kullanici ID listesi
    OncekiAyNot     As String
    PlanBaslangic   As Date
    PlanBitis       As Date
    GerceklesenBitis As Double     ' 0 = henuz bitmedi
    Durum           As String
    CubukYazisi     As String
    Yuzde           As Double
    Aciklama        As String
    SatirNo         As Long
End Type

Public Type TKullanici
    ID          As Long
    Ad          As String
    KullaniciAdi As String
    Rol         As String
    FotoAdi     As String
    Renk        As String
    YGorev      As Boolean
    YYorum      As Boolean
    YGecikme    As Boolean
    YKullanici  As Boolean
    Sayfa       As String
    Aktif       As Boolean
End Type

Public Type TProje
    ID   As Long
    Ad   As String
    Renk As String
    Sira As Long
End Type

Public Type TKategori
    ID   As Long
    Ad   As String
    Renk As String
    Sira As Long
End Type

'===============================================================================
' Temel yardimcilar
'===============================================================================
Public Function Sayfa(ByVal ad As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 513, "modVeri.Sayfa", _
            "'" & ad & "' sayfasi bulunamadi. Once Kurulum makrosunu calistirin."
    End If
    Set Sayfa = ws
End Function

Public Function SonSatir(ByVal ws As Worksheet) As Long
    SonSatir = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If SonSatir < 1 Then SonSatir = 1
End Function

'--- Ayarlar -------------------------------------------------------------------
Public Function Ayar(ByVal anahtar As String, Optional ByVal varsayilan As Variant = "") As Variant
    Dim ws As Worksheet, i As Long, son As Long
    Set ws = Sayfa(SH_AYARLAR)
    son = SonSatir(ws)
    For i = 2 To son
        If StrComp(CStr(ws.Cells(i, 1).Value), anahtar, vbTextCompare) = 0 Then
            Ayar = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    Ayar = varsayilan
End Function

Public Sub AyarYaz(ByVal anahtar As String, ByVal deger As Variant)
    Dim ws As Worksheet, i As Long, son As Long
    Set ws = Sayfa(SH_AYARLAR)
    son = SonSatir(ws)
    For i = 2 To son
        If StrComp(CStr(ws.Cells(i, 1).Value), anahtar, vbTextCompare) = 0 Then
            ws.Cells(i, 2).Value = deger
            Exit Sub
        End If
    Next i
    ws.Cells(son + 1, 1).Value = anahtar
    ws.Cells(son + 1, 2).Value = deger
End Sub

Public Function YeniID(ByVal sayacAnahtari As String) As Long
    Dim n As Long
    n = CLng(Val(Ayar(sayacAnahtari, 0))) + 1
    AyarYaz sayacAnahtari, n
    YeniID = n
End Function

'===============================================================================
' Kullanicilar
'===============================================================================
Public Function KullaniciSayisi() As Long
    KullaniciSayisi = SonSatir(Sayfa(SH_KULLANICILAR)) - 1
    If KullaniciSayisi < 0 Then KullaniciSayisi = 0
End Function

Public Function KullaniciOku(ByVal satir As Long) As TKullanici
    Dim ws As Worksheet, u As TKullanici
    Set ws = Sayfa(SH_KULLANICILAR)
    u.ID = CLng(Val(ws.Cells(satir, K_ID).Value))
    u.Ad = CStr(ws.Cells(satir, K_AD).Value)
    u.KullaniciAdi = CStr(ws.Cells(satir, K_KADI).Value)
    u.Rol = CStr(ws.Cells(satir, K_ROL).Value)
    u.FotoAdi = CStr(ws.Cells(satir, K_FOTO).Value)
    u.Renk = CStr(ws.Cells(satir, K_RENK).Value)
    u.YGorev = (ws.Cells(satir, K_Y_GOREV).Value = True)
    u.YYorum = (ws.Cells(satir, K_Y_YORUM).Value = True)
    u.YGecikme = (ws.Cells(satir, K_Y_GECIKME).Value = True)
    u.YKullanici = (ws.Cells(satir, K_Y_KULLANICI).Value = True)
    u.Sayfa = CStr(ws.Cells(satir, K_SAYFA).Value)
    u.Aktif = (ws.Cells(satir, K_AKTIF).Value = True)
    KullaniciOku = u
End Function

Public Function KullaniciSatiri(ByVal kullaniciID As Long) As Long
    Dim ws As Worksheet, i As Long
    Set ws = Sayfa(SH_KULLANICILAR)
    For i = 2 To SonSatir(ws)
        If CLng(Val(ws.Cells(i, K_ID).Value)) = kullaniciID Then
            KullaniciSatiri = i
            Exit Function
        End If
    Next i
    KullaniciSatiri = 0
End Function

Public Function KullaniciGetir(ByVal kullaniciID As Long) As TKullanici
    Dim s As Long
    s = KullaniciSatiri(kullaniciID)
    If s > 0 Then KullaniciGetir = KullaniciOku(s)
End Function

Public Function KullaniciAdiylaBul(ByVal kadi As String) As Long
    Dim ws As Worksheet, i As Long
    Set ws = Sayfa(SH_KULLANICILAR)
    For i = 2 To SonSatir(ws)
        If StrComp(CStr(ws.Cells(i, K_KADI).Value), kadi, vbTextCompare) = 0 Then
            KullaniciAdiylaBul = CLng(Val(ws.Cells(i, K_ID).Value))
            Exit Function
        End If
    Next i
    KullaniciAdiylaBul = 0
End Function

Public Function KullaniciEkle(ByVal ad As String, ByVal kadi As String, _
                              ByVal parola As String, ByVal rol As String, _
                              ByVal renk As String) As Long
    Dim ws As Worksheet, r As Long, yeni As Long
    Set ws = Sayfa(SH_KULLANICILAR)
    r = SonSatir(ws) + 1
    yeni = YeniID(AYAR_SON_KUL_ID)
    ws.Cells(r, K_ID).Value = yeni
    ws.Cells(r, K_AD).Value = ad
    ws.Cells(r, K_KADI).Value = kadi
    ws.Cells(r, K_PAROLA).Value = ParolaKaristir(parola)
    ws.Cells(r, K_ROL).Value = rol
    ws.Cells(r, K_FOTO).Value = ""
    ws.Cells(r, K_RENK).Value = renk
    ws.Cells(r, K_Y_GOREV).Value = (rol = ROL_ADMIN)
    ws.Cells(r, K_Y_YORUM).Value = True
    ws.Cells(r, K_Y_GECIKME).Value = (rol = ROL_ADMIN)
    ws.Cells(r, K_Y_KULLANICI).Value = (rol = ROL_ADMIN)
    ws.Cells(r, K_SAYFA).Value = SH_CIZELGE
    ws.Cells(r, K_AKTIF).Value = True
    KullaniciEkle = yeni
End Function

' Parolayi duz metin saklamamak icin basit bir karistirma.
' NOT: Bu gercek guvenlik degildir, sadece dosyayi acan birinin parolayi
' bir bakista okumasini engeller. VBA'da gercek kriptografi yok.
Public Function ParolaKaristir(ByVal p As String) As String
    Dim i As Long, k As Long, s As String
    k = 7
    For i = 1 To Len(p)
        k = (k * 31 + Asc(Mid$(p, i, 1))) Mod 65521
        s = s & Right$("0" & Hex$((Asc(Mid$(p, i, 1)) Xor (k And 255))), 2)
    Next i
    ParolaKaristir = s & "-" & Format$(Len(p), "00")
End Function

Public Function ParolaDogru(ByVal kullaniciID As Long, ByVal parola As String) As Boolean
    Dim s As Long
    s = KullaniciSatiri(kullaniciID)
    If s = 0 Then Exit Function
    ParolaDogru = (Sayfa(SH_KULLANICILAR).Cells(s, K_PAROLA).Value = ParolaKaristir(parola))
End Function

Public Function AktifKullanici() As TKullanici
    AktifKullanici = KullaniciGetir(CLng(Val(Ayar(AYAR_OTURUM, 0))))
End Function

Public Function YetkiVar(ByVal yetki As String) As Boolean
    Dim u As TKullanici
    u = AktifKullanici()
    If u.ID = 0 Then Exit Function
    If u.Rol = ROL_ADMIN Then YetkiVar = True: Exit Function
    Select Case yetki
        Case YETKI_GOREV_EKLE:  YetkiVar = u.YGorev
        Case YETKI_YORUM:       YetkiVar = u.YYorum
        Case YETKI_GECIKME:     YetkiVar = u.YGecikme
        Case YETKI_KULLANICI:   YetkiVar = u.YKullanici
    End Select
End Function

'===============================================================================
' Projeler / Kategoriler
'===============================================================================
Public Function ProjeGetir(ByVal projeID As Long) As TProje
    Dim ws As Worksheet, i As Long, p As TProje
    Set ws = Sayfa(SH_PROJELER)
    For i = 2 To SonSatir(ws)
        If CLng(Val(ws.Cells(i, P_ID).Value)) = projeID Then
            p.ID = projeID
            p.Ad = CStr(ws.Cells(i, P_AD).Value)
            p.Renk = CStr(ws.Cells(i, P_RENK).Value)
            p.Sira = CLng(Val(ws.Cells(i, P_SIRA).Value))
            Exit For
        End If
    Next i
    ProjeGetir = p
End Function

Public Function KategoriGetir(ByVal katID As Long) As TKategori
    Dim ws As Worksheet, i As Long, k As TKategori
    Set ws = Sayfa(SH_KATEGORILER)
    For i = 2 To SonSatir(ws)
        If CLng(Val(ws.Cells(i, T_ID).Value)) = katID Then
            k.ID = katID
            k.Ad = CStr(ws.Cells(i, T_AD).Value)
            k.Renk = CStr(ws.Cells(i, T_RENK).Value)
            k.Sira = CLng(Val(ws.Cells(i, T_SIRA).Value))
            Exit For
        End If
    Next i
    KategoriGetir = k
End Function

Public Function KategoriAdiylaBul(ByVal ad As String) As Long
    Dim ws As Worksheet, i As Long
    Set ws = Sayfa(SH_KATEGORILER)
    For i = 2 To SonSatir(ws)
        If StrComp(CStr(ws.Cells(i, T_AD).Value), ad, vbTextCompare) = 0 Then
            KategoriAdiylaBul = CLng(Val(ws.Cells(i, T_ID).Value))
            Exit Function
        End If
    Next i
End Function

Public Function ProjeAdiylaBul(ByVal ad As String) As Long
    Dim ws As Worksheet, i As Long
    Set ws = Sayfa(SH_PROJELER)
    For i = 2 To SonSatir(ws)
        If StrComp(CStr(ws.Cells(i, P_AD).Value), ad, vbTextCompare) = 0 Then
            ProjeAdiylaBul = CLng(Val(ws.Cells(i, P_ID).Value))
            Exit Function
        End If
    Next i
End Function

'===============================================================================
' Gorevler
'===============================================================================
Public Function GorevOku(ByVal satir As Long) As TGorev
    Dim ws As Worksheet, g As TGorev
    Set ws = Sayfa(SH_GOREVLER)
    g.ID = CLng(Val(ws.Cells(satir, G_ID).Value))
    g.ProjeID = CLng(Val(ws.Cells(satir, G_PROJE).Value))
    g.KategoriID = CLng(Val(ws.Cells(satir, G_KATEGORI).Value))
    g.Sira = CLng(Val(ws.Cells(satir, G_SIRA).Value))
    g.Ad = CStr(ws.Cells(satir, G_AD).Value)
    g.Sorumlular = CStr(ws.Cells(satir, G_SORUMLU).Value)
    g.OncekiAyNot = CStr(ws.Cells(satir, G_ONCEKI).Value)
    g.PlanBaslangic = CDate(ws.Cells(satir, G_PLANBAS).Value)
    g.PlanBitis = CDate(ws.Cells(satir, G_PLANBIT).Value)
    If IsEmpty(ws.Cells(satir, G_GERCEK).Value) Or ws.Cells(satir, G_GERCEK).Value = "" Then
        g.GerceklesenBitis = 0
    Else
        g.GerceklesenBitis = CDbl(ws.Cells(satir, G_GERCEK).Value)
    End If
    g.Durum = CStr(ws.Cells(satir, G_DURUM).Value)
    g.CubukYazisi = CStr(ws.Cells(satir, G_CUBUK).Value)
    g.Yuzde = CDbl(Val(ws.Cells(satir, G_YUZDE).Value))
    g.Aciklama = CStr(ws.Cells(satir, G_ACIKLAMA).Value)
    g.SatirNo = satir
    GorevOku = g
End Function

Public Function GorevSatiri(ByVal gorevID As Long) As Long
    Dim ws As Worksheet, i As Long
    Set ws = Sayfa(SH_GOREVLER)
    For i = 2 To SonSatir(ws)
        If CLng(Val(ws.Cells(i, G_ID).Value)) = gorevID Then
            GorevSatiri = i
            Exit Function
        End If
    Next i
End Function

Public Function GorevGetir(ByVal gorevID As Long) As TGorev
    Dim s As Long
    s = GorevSatiri(gorevID)
    If s > 0 Then GorevGetir = GorevOku(s)
End Function

' Verilen ayla kesisen gorevlerin satir numaralarini dondurur.
' Kesisme testi: gorev [PlanBaslangic .. Bitis] araligi ay araligiyla ortusuyor mu.
' Bitis = gerceklesen bitis varsa o, yoksa planlanan bitis (gecikme seridi de sigsin).
Public Function AyinGorevleri(ByVal yil As Long, ByVal ay As Long) As Collection
    Dim ws As Worksheet, i As Long, g As TGorev
    Dim ayBas As Date, aySon As Date, bitis As Date
    Dim c As New Collection

    Set ws = Sayfa(SH_GOREVLER)
    ayBas = DateSerial(yil, ay, 1)
    aySon = DateSerial(yil, ay + 1, 0)

    For i = 2 To SonSatir(ws)
        If ws.Cells(i, G_ID).Value <> "" Then
            g = GorevOku(i)
            If g.GerceklesenBitis > 0 Then
                bitis = CDate(g.GerceklesenBitis)
            Else
                bitis = g.PlanBitis
            End If
            If bitis < g.PlanBitis Then bitis = g.PlanBitis
            If g.PlanBaslangic <= aySon And bitis >= ayBas Then
                c.Add i
            End If
        End If
    Next i
    Set AyinGorevleri = c
End Function

Public Function GorevEkle(ByRef g As TGorev) As Long
    Dim ws As Worksheet, r As Long, yeni As Long
    Set ws = Sayfa(SH_GOREVLER)
    r = SonSatir(ws) + 1
    yeni = YeniID(AYAR_SON_GOREV_ID)
    g.ID = yeni
    ws.Cells(r, G_ID).Value = yeni
    GorevYaz r, g
    ws.Cells(r, G_OLUSTURAN).Value = CLng(Val(Ayar(AYAR_OTURUM, 0)))
    ws.Cells(r, G_TARIH).Value = Now
    GorevEkle = yeni
End Function

Public Sub GorevGuncelle(ByRef g As TGorev)
    Dim s As Long
    s = GorevSatiri(g.ID)
    If s = 0 Then Err.Raise vbObjectError + 514, "modVeri.GorevGuncelle", _
        "Guncellenecek gorev bulunamadi (ID " & g.ID & ")."
    GorevYaz s, g
End Sub

Private Sub GorevYaz(ByVal satir As Long, ByRef g As TGorev)
    Dim ws As Worksheet
    Set ws = Sayfa(SH_GOREVLER)
    ws.Cells(satir, G_PROJE).Value = g.ProjeID
    ws.Cells(satir, G_KATEGORI).Value = g.KategoriID
    ws.Cells(satir, G_SIRA).Value = g.Sira
    ws.Cells(satir, G_AD).Value = g.Ad
    ws.Cells(satir, G_SORUMLU).Value = g.Sorumlular
    ws.Cells(satir, G_ONCEKI).Value = g.OncekiAyNot
    ws.Cells(satir, G_PLANBAS).Value = g.PlanBaslangic
    ws.Cells(satir, G_PLANBIT).Value = g.PlanBitis
    If g.GerceklesenBitis > 0 Then
        ws.Cells(satir, G_GERCEK).Value = CDate(g.GerceklesenBitis)
    Else
        ws.Cells(satir, G_GERCEK).ClearContents
    End If
    ws.Cells(satir, G_DURUM).Value = g.Durum
    ws.Cells(satir, G_CUBUK).Value = g.CubukYazisi
    ws.Cells(satir, G_YUZDE).Value = g.Yuzde
    ws.Cells(satir, G_ACIKLAMA).Value = g.Aciklama
End Sub

Public Sub GorevSil(ByVal gorevID As Long)
    Dim s As Long
    s = GorevSatiri(gorevID)
    If s > 0 Then Sayfa(SH_GOREVLER).Rows(s).Delete
    YorumlariSil gorevID
End Sub

' Gecikme gun sayisi. 0 = gecikme yok.
Public Function GecikmeGunu(ByRef g As TGorev) As Long
    Dim bitis As Date
    If g.GerceklesenBitis > 0 Then
        bitis = CDate(g.GerceklesenBitis)
    ElseIf g.Durum <> DURUM_TAMAM And Date > g.PlanBitis Then
        bitis = Date                      ' hala bitmedi, gecikme bugune kadar buyuyor
    Else
        Exit Function
    End If
    If bitis > g.PlanBitis Then GecikmeGunu = CLng(bitis - g.PlanBitis)
End Function

'===============================================================================
' Yorumlar
'===============================================================================
Public Function YorumSayisi(ByVal gorevID As Long) As Long
    Dim ws As Worksheet, i As Long, n As Long
    Set ws = Sayfa(SH_YORUMLAR)
    For i = 2 To SonSatir(ws)
        If CLng(Val(ws.Cells(i, Y_GOREV).Value)) = gorevID Then n = n + 1
    Next i
    YorumSayisi = n
End Function

Public Function YorumSatirlari(ByVal gorevID As Long) As Collection
    Dim ws As Worksheet, i As Long
    Dim c As New Collection
    Set ws = Sayfa(SH_YORUMLAR)
    For i = 2 To SonSatir(ws)
        If CLng(Val(ws.Cells(i, Y_GOREV).Value)) = gorevID Then c.Add i
    Next i
    Set YorumSatirlari = c
End Function

Public Sub YorumEkle(ByVal gorevID As Long, ByVal metin As String)
    Dim ws As Worksheet, r As Long
    Set ws = Sayfa(SH_YORUMLAR)
    r = SonSatir(ws) + 1
    ws.Cells(r, Y_ID).Value = YeniID(AYAR_SON_YORUM_ID)
    ws.Cells(r, Y_GOREV).Value = gorevID
    ws.Cells(r, Y_KULLANICI).Value = CLng(Val(Ayar(AYAR_OTURUM, 0)))
    ws.Cells(r, Y_TARIH).Value = Now
    ws.Cells(r, Y_METIN).Value = metin
End Sub

Private Sub YorumlariSil(ByVal gorevID As Long)
    Dim ws As Worksheet, i As Long
    Set ws = Sayfa(SH_YORUMLAR)
    For i = SonSatir(ws) To 2 Step -1
        If CLng(Val(ws.Cells(i, Y_GOREV).Value)) = gorevID Then ws.Rows(i).Delete
    Next i
End Sub

'===============================================================================
' Gecikme kayitlari
'===============================================================================
Public Function GecikmeSatiri(ByVal gorevID As Long) As Long
    Dim ws As Worksheet, i As Long
    Set ws = Sayfa(SH_GECIKMELER)
    For i = 2 To SonSatir(ws)
        If CLng(Val(ws.Cells(i, D_GOREV).Value)) = gorevID Then
            GecikmeSatiri = i
            Exit Function
        End If
    Next i
End Function

Public Sub GecikmeKaydet(ByVal gorevID As Long, ByVal neden As String, _
                         ByVal sorumluluk As String)
    Dim ws As Worksheet, r As Long
    Set ws = Sayfa(SH_GECIKMELER)
    r = GecikmeSatiri(gorevID)
    If r = 0 Then
        r = SonSatir(ws) + 1
        ws.Cells(r, D_ID).Value = YeniID(AYAR_SON_GECIKME)
        ws.Cells(r, D_GOREV).Value = gorevID
    End If
    ws.Cells(r, D_NEDEN).Value = neden
    ws.Cells(r, D_SORUMLULUK).Value = sorumluluk
    ws.Cells(r, D_KAYDEDEN).Value = CLng(Val(Ayar(AYAR_OTURUM, 0)))
    ws.Cells(r, D_TARIH).Value = Now
End Sub

'===============================================================================
' Off / tatil gunleri
'===============================================================================
Public Function OffGunMu(ByVal tarih As Date) As Boolean
    Dim ws As Worksheet, i As Long
    Set ws = Sayfa(SH_OFFGUNLER)
    For i = 2 To SonSatir(ws)
        If ws.Cells(i, O_TARIH).Value <> "" Then
            If CDate(ws.Cells(i, O_TARIH).Value) = tarih Then
                OffGunMu = True
                Exit Function
            End If
        End If
    Next i
End Function

Public Sub OffGunDegistir(ByVal tarih As Date, Optional ByVal aciklama As String = "Idari tatil")
    Dim ws As Worksheet, i As Long
    Set ws = Sayfa(SH_OFFGUNLER)
    For i = SonSatir(ws) To 2 Step -1
        If ws.Cells(i, O_TARIH).Value <> "" Then
            If CDate(ws.Cells(i, O_TARIH).Value) = tarih Then
                ws.Rows(i).Delete          ' zaten isaretliyse kaldir
                Exit Sub
            End If
        End If
    Next i
    i = SonSatir(ws) + 1
    ws.Cells(i, O_TARIH).Value = tarih
    ws.Cells(i, O_ACIKLAMA).Value = aciklama
End Sub

Public Function HaftaSonuMu(ByVal tarih As Date) As Boolean
    HaftaSonuMu = (Weekday(tarih, vbMonday) >= 6)
End Function
