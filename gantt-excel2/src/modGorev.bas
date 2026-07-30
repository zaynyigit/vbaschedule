Attribute VB_Name = "modGorev"
'===============================================================================
' modGorev - Gorev ekleme / duzenleme
'===============================================================================
Option Explicit

' Alan numaralari - Panel sirasi ile birebir ayni olmali
Private Const A_PROJE = 1, A_KATEGORI = 2, A_AD = 3, A_SORUMLU = 4, A_SIRA = 5
Private Const A_BAS = 6, A_PLANBIT = 7, A_GERCEK = 8, A_DURUM = 9
Private Const A_CUBUK = 10, A_YUZDE = 11, A_ONCEKI = 12, A_ACIKLAMA = 13

'===============================================================================
' Acma
'===============================================================================
Public Sub GorevEkleAc()
    GorevEkleAcTarihle Date
End Sub

Public Sub GorevEkleAcTarihle(ByVal baslangic As Date)
    If Not modGiris.OturumGerekli() Then Exit Sub
    If Not modVeri.YetkiVar(YETKI_GOREV_EKLE) Then
        MsgBox "Gorev ekleme yetkin yok. Yoneticine basvur.", vbExclamation, APP_ADI
        Exit Sub
    End If

    modModal.Temizle
    modModal.Baglam = ""

    modModal.Alan "Proje", MT_LISTE, IlkProjeAdi(), ProjeListesi()
    modModal.Alan "Kategori", MT_LISTE, IlkKategoriAdi(), KategoriListesi()
    modModal.Alan "Gorev adi", MT_METIN, ""
    modModal.Alan "Sorumlular (virgulle ayir)", MT_METIN, ""
    modModal.Alan "Sira no", MT_SAYI, CStr(SonrakiSira())
    modModal.Alan "Baslangic", MT_TARIH, Format$(baslangic, "dd.mm.yyyy")
    modModal.Alan "Planlanan bitis", MT_TARIH, Format$(baslangic, "dd.mm.yyyy")
    modModal.Alan "Gerceklesen bitis (bos = devam)", MT_TARIH, ""
    modModal.Alan "Durum", MT_LISTE, "Planlandi", "Planlandi,Devam ediyor,Tamamlandi"
    modModal.Alan "Cubuk yazisi", MT_METIN, ""
    modModal.Alan "Ilerleme %", MT_SAYI, "0"
    modModal.Alan "Onceki aydan kalan", MT_METIN, ""
    modModal.Alan "Aciklama", MT_COKSATIR, ""

    modModal.Goster "gorev_ekle", "Gorev ekle", "Gorevi ekle"
End Sub

Public Sub GorevDuzenleAc(ByVal gorevID As Long)
    Dim g As TGorev, p As TProje, k As TKategori
    If Not modGiris.OturumGerekli() Then Exit Sub

    g = modVeri.GorevGetir(gorevID)
    If g.ID = 0 Then Exit Sub
    p = modVeri.ProjeGetir(g.ProjeID)
    k = modVeri.KategoriGetir(g.KategoriID)

    modModal.Temizle
    modModal.Baglam = CStr(gorevID)

    modModal.Alan "Proje", MT_LISTE, p.Ad, ProjeListesi()
    modModal.Alan "Kategori", MT_LISTE, k.Ad, KategoriListesi()
    modModal.Alan "Gorev adi", MT_METIN, g.Ad
    modModal.Alan "Sorumlular (virgulle ayir)", MT_METIN, IDlerdenAdlar(g.Sorumlular)
    modModal.Alan "Sira no", MT_SAYI, CStr(g.Sira)
    modModal.Alan "Baslangic", MT_TARIH, Format$(g.PlanBaslangic, "dd.mm.yyyy")
    modModal.Alan "Planlanan bitis", MT_TARIH, Format$(g.PlanBitis, "dd.mm.yyyy")
    modModal.Alan "Gerceklesen bitis (bos = devam)", MT_TARIH, _
                  IIf(g.GerceklesenBitis > 0, Format$(CDate(g.GerceklesenBitis), "dd.mm.yyyy"), "")
    modModal.Alan "Durum", MT_LISTE, DurumEtiketi(g.Durum), "Planlandi,Devam ediyor,Tamamlandi"
    modModal.Alan "Cubuk yazisi", MT_METIN, g.CubukYazisi
    modModal.Alan "Ilerleme %", MT_SAYI, CStr(g.Yuzde)
    modModal.Alan "Onceki aydan kalan", MT_METIN, g.OncekiAyNot
    modModal.Alan "Aciklama", MT_COKSATIR, g.Aciklama

    modModal.Goster "gorev_duzenle", "Gorev duzenle - " & g.Ad, "Degisiklikleri kaydet"
End Sub

'===============================================================================
' Kaydetme
'===============================================================================
Public Function GorevKaydet(ByVal duzenleme As Boolean) As Boolean
    Dim g As TGorev
    Dim bas As Date, planBit As Date, gercek As Date
    Dim hata As String

    ' --- dogrulama ---
    If Len(Trim$(modModal.Deger(A_AD))) = 0 Then
        hata = "Gorev adi bos birakilamaz."
    End If

    bas = modModal.DegerTarih(A_BAS)
    planBit = modModal.DegerTarih(A_PLANBIT)

    If hata = "" And bas = 0 Then hata = "Baslangic tarihi gecerli degil."
    If hata = "" And planBit = 0 Then hata = "Planlanan bitis tarihi gecerli degil."
    If hata = "" And planBit < bas Then
        hata = "Planlanan bitis, baslangictan once olamaz."
    End If

    If Len(Trim$(modModal.Deger(A_GERCEK))) > 0 Then
        gercek = modModal.DegerTarih(A_GERCEK)
        If hata = "" And gercek > 0 And gercek < bas Then
            hata = "Gerceklesen bitis, baslangictan once olamaz."
        End If
    End If

    If hata = "" Then
        Dim yz As Double
        yz = modModal.DegerSayi(A_YUZDE)
        If yz < 0 Or yz > 100 Then hata = "Ilerleme yuzdesi 0 ile 100 arasinda olmali."
    End If

    If hata <> "" Then
        MsgBox hata, vbExclamation, APP_ADI
        Exit Function
    End If

    ' --- doldur ---
    Dim bulunamayan As String
    g.ProjeID = modVeri.ProjeAdiylaBul(modModal.Deger(A_PROJE))
    g.KategoriID = modVeri.KategoriAdiylaBul(modModal.Deger(A_KATEGORI))
    g.Ad = Trim$(modModal.Deger(A_AD))
    g.Sorumlular = AdlardanIDler(modModal.Deger(A_SORUMLU), bulunamayan)
    If Len(bulunamayan) > 0 Then
        MsgBox "Bu isimler kullanici listesinde yok, sorumlu olarak atanmadi:" & vbCrLf & _
               "  " & bulunamayan & vbCrLf & vbCrLf & _
               "Kullanicilar panelinden ekleyip gorevi tekrar duzenleyebilirsin.", _
               vbInformation, APP_ADI
    End If
    g.Sira = CLng(modModal.DegerSayi(A_SIRA))
    g.PlanBaslangic = bas
    g.PlanBitis = planBit
    g.GerceklesenBitis = IIf(gercek > 0, CDbl(gercek), 0)
    g.Durum = DurumKodu(modModal.Deger(A_DURUM))
    g.CubukYazisi = modModal.Deger(A_CUBUK)
    g.Yuzde = modModal.DegerSayi(A_YUZDE)
    g.OncekiAyNot = modModal.Deger(A_ONCEKI)
    g.Aciklama = modModal.Deger(A_ACIKLAMA)

    If g.ProjeID = 0 Then
        MsgBox "Secilen proje bulunamadi.", vbExclamation, APP_ADI
        Exit Function
    End If
    If g.KategoriID = 0 Then
        MsgBox "Secilen kategori bulunamadi.", vbExclamation, APP_ADI
        Exit Function
    End If

    ' Tamamlandi isaretlendi ama gerceklesen bitis girilmediyse bugunu yaz
    If g.Durum = DURUM_TAMAM And g.GerceklesenBitis = 0 Then
        g.GerceklesenBitis = CDbl(Date)
    End If
    If g.Durum = DURUM_TAMAM Then g.Yuzde = 100

    If duzenleme Then
        g.ID = CLng(Val(modModal.Baglam))
        modVeri.GorevGuncelle g
    Else
        modVeri.GorevEkle g
    End If

    ' Gecikme olustuysa nedenini hemen sor
    Dim gecikme As Long
    gecikme = modVeri.GecikmeGunu(g)

    modCizelge.Ciz
    GorevKaydet = True

    If gecikme > 0 And modVeri.GecikmeSatiri(g.ID) = 0 Then
        ' Gecikme paneli ayni panel sayfasini devralir. Burada True donersek
        ' modModal.TamamTiklandi paneli hemen kapatiyor ve panel bir anlik
        ' gorunup kayboluyordu; devir gerceklesirse False donuyoruz.
        If modGecikme.GecikmePaneliAc(g.ID) Then GorevKaydet = False
    End If
End Function

'===============================================================================
' Silme
'===============================================================================
Public Sub GorevSilSor(ByVal gorevID As Long)
    Dim g As TGorev
    g = modVeri.GorevGetir(gorevID)
    If g.ID = 0 Then Exit Sub
    If MsgBox("'" & g.Ad & "' gorevi ve tum yorumlari silinecek." & vbCrLf & _
              "Emin misin?", vbExclamation + vbYesNo + vbDefaultButton2, APP_ADI) = vbYes Then
        modVeri.GorevSil gorevID
        modCizelge.Ciz
    End If
End Sub

'===============================================================================
' Yardimcilar
'===============================================================================
Public Function DurumKodu(ByVal etiket As String) As String
    Select Case True
        Case InStr(1, etiket, "Tamam", vbTextCompare) > 0: DurumKodu = DURUM_TAMAM
        Case InStr(1, etiket, "Devam", vbTextCompare) > 0: DurumKodu = DURUM_DEVAM
        Case Else: DurumKodu = DURUM_PLAN
    End Select
End Function

Public Function DurumEtiketi(ByVal kod As String) As String
    Select Case kod
        Case DURUM_TAMAM: DurumEtiketi = "Tamamlandi"
        Case DURUM_DEVAM: DurumEtiketi = "Devam ediyor"
        Case Else: DurumEtiketi = "Planlandi"
    End Select
End Function

Public Function ProjeListesi() As String
    Dim ws As Worksheet, i As Long, s As String
    Set ws = modVeri.Sayfa(SH_PROJELER)
    For i = 2 To modVeri.SonSatir(ws)
        If ws.Cells(i, P_AD).Value <> "" Then
            s = s & IIf(Len(s) > 0, ",", "") & ws.Cells(i, P_AD).Value
        End If
    Next i
    ProjeListesi = s
End Function

Public Function KategoriListesi() As String
    Dim ws As Worksheet, i As Long, s As String
    Set ws = modVeri.Sayfa(SH_KATEGORILER)
    For i = 2 To modVeri.SonSatir(ws)
        If ws.Cells(i, T_AD).Value <> "" Then
            s = s & IIf(Len(s) > 0, ",", "") & ws.Cells(i, T_AD).Value
        End If
    Next i
    KategoriListesi = s
End Function

Private Function IlkProjeAdi() As String
    Dim p As Variant
    p = Split(ProjeListesi(), ",")
    If UBound(p) >= 0 Then IlkProjeAdi = p(0)
End Function

Private Function IlkKategoriAdi() As String
    Dim p As Variant
    p = Split(KategoriListesi(), ",")
    If UBound(p) >= 0 Then IlkKategoriAdi = p(0)
End Function

Private Function SonrakiSira() As Long
    Dim ws As Worksheet, i As Long, m As Long
    Set ws = modVeri.Sayfa(SH_GOREVLER)
    For i = 2 To modVeri.SonSatir(ws)
        If CLng(Val(ws.Cells(i, G_SIRA).Value)) > m Then m = CLng(Val(ws.Cells(i, G_SIRA).Value))
    Next i
    SonrakiSira = m + 1
End Function

' "Yigit Bayir, Adem Kaya" -> "1,2"
' Tanimadigi isimleri sessizce yutuyordu; artik bulunamayan'da geri veriyor ki
' cagiran taraf kullaniciyi uyarabilsin.
Public Function AdlardanIDler(ByVal adlar As String, _
                              Optional ByRef bulunamayan As String) As String
    Dim p As Variant, i As Long, ws As Worksheet, j As Long, s As String, ad As String
    Dim eslesti As Boolean

    bulunamayan = ""
    If Len(Trim$(adlar)) = 0 Then Exit Function
    Set ws = modVeri.Sayfa(SH_KULLANICILAR)
    p = Split(adlar, ",")
    For i = LBound(p) To UBound(p)
        ad = Trim$(p(i))
        If Len(ad) > 0 Then
            eslesti = False
            For j = 2 To modVeri.SonSatir(ws)
                If StrComp(CStr(ws.Cells(j, K_AD).Value), ad, vbTextCompare) = 0 _
                Or StrComp(CStr(ws.Cells(j, K_KADI).Value), ad, vbTextCompare) = 0 Then
                    s = s & IIf(Len(s) > 0, ",", "") & CStr(ws.Cells(j, K_ID).Value)
                    eslesti = True
                    Exit For
                End If
            Next j
            If Not eslesti Then
                bulunamayan = bulunamayan & IIf(Len(bulunamayan) > 0, ", ", "") & ad
            End If
        End If
    Next i
    AdlardanIDler = s
End Function

' "1,2" -> "Yigit Bayir, Adem Kaya"
Public Function IDlerdenAdlar(ByVal idler As String) As String
    Dim p As Variant, i As Long, u As TKullanici, s As String
    If Len(Trim$(idler)) = 0 Then Exit Function
    p = Split(idler, ",")
    For i = LBound(p) To UBound(p)
        u = modVeri.KullaniciGetir(CLng(Val(p(i))))
        If u.ID > 0 Then s = s & IIf(Len(s) > 0, ", ", "") & u.Ad
    Next i
    IDlerdenAdlar = s
End Function
