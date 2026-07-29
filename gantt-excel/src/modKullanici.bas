Attribute VB_Name = "modKullanici"
'===============================================================================
' modKullanici - Kullanici ve kategori yonetimi (yalnizca yetkili kullanicilar)
'===============================================================================
Option Explicit

' Kullanici paneli alan numaralari
Private Const U_AD = 1, U_KADI = 2, U_PAROLA = 3, U_ROL = 4, U_RENK = 5
Private Const U_BASLIK = 6, U_YGOREV = 7, U_YYORUM = 8, U_YGECIKME = 9
Private Const U_YKULLANICI = 10, U_SAYFA = 11, U_AKTIF = 12
Private Const U_BASLIK2 = 13, U_LISTE = 14

' Kategori paneli alan numaralari
Private Const T_ADI = 1, T_RENGI = 2, T_BASLIK = 3, T_LISTESI = 4

'===============================================================================
' Kullanici paneli
'===============================================================================
Public Sub KullaniciPaneliAc()
    If Not modGiris.OturumGerekli() Then Exit Sub
    If Not modVeri.YetkiVar(YETKI_KULLANICI) Then
        MsgBox "Kullanici yonetimi icin yonetici yetkisi gerekiyor.", vbExclamation, APP_ADI
        Exit Sub
    End If

    modModal.Temizle
    modModal.Baglam = ""

    modModal.Alan "Ad Soyad", MT_METIN, ""
    modModal.Alan "Kullanici adi", MT_METIN, ""
    modModal.Alan "Parola", MT_METIN, ""
    modModal.Alan "Rol", MT_LISTE, "Kullanici", "Kullanici,Yonetici"
    modModal.Alan "Avatar rengi (hex)", MT_METIN, RastgeleRenk()

    modModal.Alan "Yetkiler", MT_BASLIK
    modModal.Alan "Gorev ekleyebilir", MT_ONAY, "Evet"
    modModal.Alan "Yorum yapabilir", MT_ONAY, "Evet"
    modModal.Alan "Gecikme kaydi girebilir", MT_ONAY, "Hayir"
    modModal.Alan "Kullanici yonetebilir", MT_ONAY, "Hayir"
    modModal.Alan "Erisebilecegi sayfa", MT_LISTE, SH_CIZELGE, SH_CIZELGE
    modModal.Alan "Aktif", MT_ONAY, "Evet"

    modModal.Alan "Kayitli kullanicilar (" & modVeri.KullaniciSayisi() & ")", MT_BASLIK
    modModal.Alan "Liste", MT_SALTOKUR, KullaniciListesiMetni()

    modModal.Goster "kullanici_ekle", "Kullanici ekle", "Kullaniciyi ekle"
End Sub

Public Function KullaniciKaydet() As Boolean
    Dim ad As String, kadi As String, parola As String, rol As String, renk As String
    Dim yeni As Long, satir As Long, ws As Worksheet

    ad = Trim$(modModal.Deger(U_AD))
    kadi = LCase$(Trim$(modModal.Deger(U_KADI)))
    parola = modModal.Deger(U_PAROLA)
    rol = IIf(InStr(1, modModal.Deger(U_ROL), "Yonetici", vbTextCompare) > 0, _
              ROL_ADMIN, ROL_KULLANICI)
    renk = Replace(Trim$(modModal.Deger(U_RENK)), "#", "")

    If Len(ad) = 0 Then
        MsgBox "Ad Soyad bos birakilamaz.", vbExclamation, APP_ADI: Exit Function
    End If
    If Len(kadi) = 0 Then
        MsgBox "Kullanici adi bos birakilamaz.", vbExclamation, APP_ADI: Exit Function
    End If
    If InStr(kadi, " ") > 0 Then
        MsgBox "Kullanici adi bosluk iceremez.", vbExclamation, APP_ADI: Exit Function
    End If
    If Len(parola) < 4 Then
        MsgBox "Parola en az 4 karakter olmali.", vbExclamation, APP_ADI: Exit Function
    End If
    If modVeri.KullaniciAdiylaBul(kadi) > 0 Then
        MsgBox "'" & kadi & "' kullanici adi zaten kayitli.", vbExclamation, APP_ADI
        Exit Function
    End If
    If Len(renk) <> 6 Then renk = "64748B"

    yeni = modVeri.KullaniciEkle(ad, kadi, parola, rol, renk)

    ' Yetkileri panelden gelen degerlerle ez
    satir = modVeri.KullaniciSatiri(yeni)
    Set ws = modVeri.Sayfa(SH_KULLANICILAR)
    ws.Cells(satir, K_Y_GOREV).Value = modModal.DegerOnay(U_YGOREV)
    ws.Cells(satir, K_Y_YORUM).Value = modModal.DegerOnay(U_YYORUM)
    ws.Cells(satir, K_Y_GECIKME).Value = modModal.DegerOnay(U_YGECIKME)
    ws.Cells(satir, K_Y_KULLANICI).Value = modModal.DegerOnay(U_YKULLANICI)
    ws.Cells(satir, K_SAYFA).Value = modModal.Deger(U_SAYFA)
    ws.Cells(satir, K_AKTIF).Value = modModal.DegerOnay(U_AKTIF)

    MsgBox ad & " eklendi." & vbCrLf & _
           "Kullanici adi: " & kadi, vbInformation, APP_ADI

    KullaniciKaydet = True
End Function

Private Function KullaniciListesiMetni() As String
    Dim ws As Worksheet, i As Long, s As String, u As TKullanici
    Set ws = modVeri.Sayfa(SH_KULLANICILAR)
    For i = 2 To modVeri.SonSatir(ws)
        u = modVeri.KullaniciOku(i)
        If u.ID > 0 Then
            s = s & IIf(Len(s) > 0, vbLf, "") & _
                Format$(u.ID, "00") & "  " & u.Ad & _
                "   (" & u.KullaniciAdi & ")   " & _
                IIf(u.Rol = ROL_ADMIN, "YONETICI", "kullanici") & _
                IIf(u.Aktif, "", "   [PASIF]")
        End If
    Next i
    If Len(s) = 0 Then s = "Kayitli kullanici yok."
    KullaniciListesiMetni = s
End Function

Private Function RastgeleRenk() As String
    Dim p As Variant
    p = Array("4F46E5", "0891B2", "DB2777", "16A34A", "F59E0B", "7C3AED", "0EA5E9", "DC2626")
    Randomize
    RastgeleRenk = p(Int(Rnd() * (UBound(p) + 1)))
End Function

'===============================================================================
' Kategori paneli
'===============================================================================
Public Sub KategoriPaneliAc()
    If Not modGiris.OturumGerekli() Then Exit Sub
    If Not modVeri.YetkiVar(YETKI_KULLANICI) Then
        MsgBox "Kategori yonetimi icin yonetici yetkisi gerekiyor.", vbExclamation, APP_ADI
        Exit Sub
    End If

    modModal.Temizle
    modModal.Alan "Yeni kategori adi", MT_METIN, ""
    modModal.Alan "Rengi (hex)", MT_METIN, RastgeleRenk()
    modModal.Alan "Mevcut kategoriler", MT_BASLIK
    modModal.Alan "Liste", MT_SALTOKUR, KategoriListesiMetni()

    modModal.Goster "kategori_ekle", "Kategoriler", "Kategoriyi ekle"
End Sub

Public Function KategoriKaydet() As Boolean
    Dim ad As String, renk As String, ws As Worksheet, r As Long, yeniID As Long, i As Long

    ad = UCase$(Trim$(modModal.Deger(T_ADI)))
    renk = Replace(Trim$(modModal.Deger(T_RENGI)), "#", "")

    If Len(ad) = 0 Then
        ' Bos birakildiysa sadece paneli kapat
        KategoriKaydet = True
        Exit Function
    End If
    If modVeri.KategoriAdiylaBul(ad) > 0 Then
        MsgBox "'" & ad & "' kategorisi zaten var.", vbExclamation, APP_ADI
        Exit Function
    End If
    If Len(renk) <> 6 Then renk = "64748B"

    Set ws = modVeri.Sayfa(SH_KATEGORILER)
    r = modVeri.SonSatir(ws) + 1
    For i = 2 To r - 1
        If CLng(Val(ws.Cells(i, T_ID).Value)) > yeniID Then yeniID = CLng(Val(ws.Cells(i, T_ID).Value))
    Next i
    yeniID = yeniID + 1

    ws.Cells(r, T_ID).Value = yeniID
    ws.Cells(r, T_AD).Value = ad
    ws.Cells(r, T_RENK).Value = renk
    ws.Cells(r, T_SIRA).Value = yeniID
    ws.Cells(r, T_AKTIF).Value = True

    KategoriKaydet = True
End Function

Private Function KategoriListesiMetni() As String
    Dim ws As Worksheet, i As Long, s As String, n As Long
    Set ws = modVeri.Sayfa(SH_KATEGORILER)
    For i = 2 To modVeri.SonSatir(ws)
        If ws.Cells(i, T_AD).Value <> "" Then
            n = KategoriGorevSayisi(CLng(Val(ws.Cells(i, T_ID).Value)))
            s = s & IIf(Len(s) > 0, vbLf, "") & _
                ws.Cells(i, T_AD).Value & "   #" & ws.Cells(i, T_RENK).Value & _
                "   " & n & " gorev"
        End If
    Next i
    If Len(s) = 0 Then s = "Kayitli kategori yok."
    KategoriListesiMetni = s
End Function

Private Function KategoriGorevSayisi(ByVal katID As Long) As Long
    Dim ws As Worksheet, i As Long, n As Long
    Set ws = modVeri.Sayfa(SH_GOREVLER)
    For i = 2 To modVeri.SonSatir(ws)
        If CLng(Val(ws.Cells(i, G_KATEGORI).Value)) = katID Then n = n + 1
    Next i
    KategoriGorevSayisi = n
End Function
