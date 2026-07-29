Attribute VB_Name = "modYorum"
'===============================================================================
' modYorum - Gorev detay ve yorum paneli
'
' Cizelgede bir bara (veya yorum rozetine) tiklandiginda acilir. Panelde
' gorevin ozeti, mevcut yorumlar ve yeni yorum kutusu bulunur.
'===============================================================================
Option Explicit

' Alan numaralari - DetayAc icindeki Alan cagrilarinin sirasiyla birebir ayni.
' MT_BASLIK de bir alan sayilir, atlanirsa numaralar kayar.
Private Const A_GOREV = 1, A_PROJE = 2, A_TARIHLER = 3, A_SORUMLU = 4
Private Const A_BASLIK = 5, A_GECMIS = 6, A_YENI = 7

'===============================================================================
' Bara tiklandi
'===============================================================================
Public Sub BaraTiklandi()
    Dim gorevID As Long
    gorevID = TiklananGorevID()
    If gorevID = 0 Then Exit Sub
    DetayAc gorevID
End Sub

' OnAction ile cagrilan sekil, gorev ID'sini AlternativeText alaninda tasir.
Public Function TiklananGorevID() As Long
    Dim sk As Shape, ws As Worksheet
    On Error Resume Next
    Set ws = modVeri.Sayfa(SH_CIZELGE)
    Set sk = ws.Shapes(Application.Caller)
    On Error GoTo 0
    If sk Is Nothing Then Exit Function
    TiklananGorevID = CLng(Val(sk.AlternativeText))
End Function

Public Sub DetayAc(ByVal gorevID As Long)
    Dim g As TGorev, p As TProje, k As TKategori
    Dim gecikme As Long

    If Not modGiris.OturumGerekli() Then Exit Sub

    g = modVeri.GorevGetir(gorevID)
    If g.ID = 0 Then Exit Sub
    p = modVeri.ProjeGetir(g.ProjeID)
    k = modVeri.KategoriGetir(g.KategoriID)
    gecikme = modVeri.GecikmeGunu(g)

    modModal.Temizle
    modModal.Baglam = CStr(gorevID)

    modModal.Alan "Gorev", MT_SALTOKUR, _
        g.Ad & "   [" & modGorev.DurumEtiketi(g.Durum) & "]"
    modModal.Alan "Proje / Kategori / Ilerleme", MT_SALTOKUR, _
        p.Ad & "  " & ChrW(183) & "  " & k.Ad & "  " & ChrW(183) & "  %" & Format$(g.Yuzde, "0")
    modModal.Alan "Tarihler", MT_SALTOKUR, TarihOzeti(g, gecikme)
    modModal.Alan "Sorumlular", MT_SALTOKUR, modGorev.IDlerdenAdlar(g.Sorumlular)

    modModal.Alan "Yorumlar", MT_BASLIK
    modModal.Alan "Gecmis", MT_SALTOKUR, YorumGecmisi(gorevID)
    modModal.Alan AktifKullaniciEtiketi() & " olarak yorum yaz", MT_COKSATIR, ""

    modModal.Goster "yorum", "Gorev detayi", "Yorumu gonder"
End Sub

Private Function TarihOzeti(ByRef g As TGorev, ByVal gecikme As Long) As String
    Dim s As String
    s = Format$(g.PlanBaslangic, "dd.mm.yyyy") & "  " & ChrW(8594) & "  " & _
        Format$(g.PlanBitis, "dd.mm.yyyy") & " (planlanan)"
    If g.GerceklesenBitis > 0 Then
        s = s & "     Gerceklesen: " & Format$(CDate(g.GerceklesenBitis), "dd.mm.yyyy")
    End If
    If gecikme > 0 Then s = s & "     SAPMA: +" & gecikme & " gun"
    TarihOzeti = s
End Function

Private Function AktifKullaniciEtiketi() As String
    Dim u As TKullanici
    u = modVeri.AktifKullanici()
    If u.ID = 0 Then AktifKullaniciEtiketi = "Misafir" Else AktifKullaniciEtiketi = u.Ad
End Function

'===============================================================================
' Yorum listesi
'===============================================================================
Public Function YorumGecmisi(ByVal gorevID As Long) As String
    Dim satirlar As Collection, i As Long, ws As Worksheet
    Dim u As TKullanici, s As String, t As Variant

    Set ws = modVeri.Sayfa(SH_YORUMLAR)
    Set satirlar = modVeri.YorumSatirlari(gorevID)

    If satirlar.Count = 0 Then
        YorumGecmisi = "Henuz yorum yok."
        Exit Function
    End If

    For i = 1 To satirlar.Count
        u = modVeri.KullaniciGetir(CLng(Val(ws.Cells(satirlar(i), Y_KULLANICI).Value)))
        t = ws.Cells(satirlar(i), Y_TARIH).Value
        s = s & IIf(Len(s) > 0, vbLf, "") & _
            u.Ad & "  " & ChrW(183) & "  " & Format$(t, "dd.mm.yyyy hh:nn") & vbLf & _
            "   " & CStr(ws.Cells(satirlar(i), Y_METIN).Value)
    Next i
    YorumGecmisi = s
End Function

'===============================================================================
' Kaydetme
'===============================================================================
Public Function YorumKaydet() As Boolean
    Dim metin As String, gorevID As Long

    gorevID = CLng(Val(modModal.Baglam))
    metin = Trim$(modModal.Deger(A_YENI))

    If gorevID = 0 Then
        YorumKaydet = True
        Exit Function
    End If

    If Len(metin) = 0 Then
        ' Yorum yazilmadiysa sadece paneli kapat, hata verme
        YorumKaydet = True
        Exit Function
    End If

    If Not modVeri.YetkiVar(YETKI_YORUM) Then
        MsgBox "Yorum yapma yetkin yok.", vbExclamation, APP_ADI
        Exit Function
    End If

    modVeri.YorumEkle gorevID, metin
    modCizelge.Ciz
    YorumKaydet = True
End Function
