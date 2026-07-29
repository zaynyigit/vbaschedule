Attribute VB_Name = "modSabitler"
'===============================================================================
' modSabitler - Tum proje genelinde kullanilan sabitler ve yerlesim tanimlari
'
' Bu modulde HICBIR mantik yoktur. Sadece isim, konum ve renk tanimlari.
' Yerlesimi degistirmek istersen once buraya bak; kod geri kalanini buradan okur.
'===============================================================================
Option Explicit

'--- Surum ---------------------------------------------------------------------
Public Const APP_ADI            As String = "Gunluk Proje Cizelgesi"
Public Const APP_SURUM          As String = "1.0"
Public Const APP_TELIF          As String = "Copyright by Yigit Bayir"

'--- Sayfa adlari --------------------------------------------------------------
' Onunde "_" olanlar veri sayfasidir, kullaniciya gosterilmez (xlSheetVeryHidden).
Public Const SH_GIRIS           As String = "Giris"
Public Const SH_CIZELGE         As String = "Cizelge"
Public Const SH_KULLANICILAR    As String = "_Kullanicilar"
Public Const SH_PROJELER        As String = "_Projeler"
Public Const SH_KATEGORILER     As String = "_Kategoriler"
Public Const SH_GOREVLER        As String = "_Gorevler"
Public Const SH_YORUMLAR        As String = "_Yorumlar"
Public Const SH_GECIKMELER      As String = "_Gecikmeler"
Public Const SH_OFFGUNLER       As String = "_OffGunler"
Public Const SH_AYARLAR         As String = "_Ayarlar"
Public Const SH_FOTOLAR         As String = "_Fotolar"

'--- Cizelge sayfasi satir yerlesimi -------------------------------------------
Public Const R_BASLIK           As Long = 2     ' Uygulama basligi + arac cubugu
Public Const R_EFSANE           As Long = 4     ' Renk aciklama seridi
Public Const R_AYADI            As Long = 6     ' Buyuk AY / YIL basligi
Public Const R_GUNHARF          As Long = 7     ' Pzt/Sal/Car harfleri
Public Const R_GUNNO            As Long = 8     ' 1..31 gun numaralari
Public Const R_VERI_BAS         As Long = 9     ' Ilk veri satiri

'--- Cizelge sayfasi sutun yerlesimi -------------------------------------------
Public Const C_SERIT            As Long = 1     ' A  - proje renk seridi
Public Const C_KATEGORI         As Long = 2     ' B
Public Const C_SIRA             As Long = 3     ' C
Public Const C_SORUMLU          As Long = 4     ' D
Public Const C_GOREV            As Long = 5     ' E
Public Const C_ONCEKI           As Long = 6     ' F
Public Const C_GUN1             As Long = 7     ' G  - ayin 1'i
Public Const GUN_SAYISI_MAX     As Long = 31
' Gun sutunlari dinamiktir: C_GUN1 + (gun - 1)
' Sagdaki iki sutun da aya gore kayar; hesabi modCizelge.SutunIlerleme yapar.

'--- Olculer (nokta / point cinsinden; 1 pt = 1/72 inc) ------------------------
Public Const W_SERIT            As Double = 5#
Public Const W_KATEGORI         As Double = 62#
Public Const W_SIRA             As Double = 22#
Public Const W_SORUMLU          As Double = 66#
Public Const W_GOREV            As Double = 150#
Public Const W_ONCEKI           As Double = 118#
Public Const W_GUN              As Double = 17#
Public Const W_ILERLEME         As Double = 56#
Public Const W_ACIKLAMA         As Double = 150#

Public Const H_SATIR            As Double = 26#   ' gorev satiri yuksekligi
Public Const H_PROJE            As Double = 22#   ' proje grup satiri
Public Const H_AYADI            As Double = 40#
Public Const H_GUNHARF          As Double = 14#
Public Const H_GUNNO            As Double = 17#

'--- Renkler -------------------------------------------------------------------
' DIKKAT: VBA renkleri BGR sirali Long olarak saklar, hex ile ayni degildir.
' O yuzden hicbir yerde ciplak sayi kullanma; her zaman Hx("RRGGBB") cagir.
' Boylece tasarimdaki hex kodunu birebir yazabiliyoruz ve cevirme hatasi olmuyor.

Public Function Hx(ByVal h As String) As Long
    If Left$(h, 1) = "#" Then h = Mid$(h, 2)
    Hx = CLng("&H" & Mid$(h, 1, 2)) _
       + CLng("&H" & Mid$(h, 3, 2)) * 256& _
       + CLng("&H" & Mid$(h, 5, 2)) * 65536&
End Function

Public Function CLR_ZEMIN() As Long
    CLR_ZEMIN = Hx("FFFFFF")
End Function
Public Function CLR_ZEMIN_ALT() As Long
    CLR_ZEMIN_ALT = Hx("FCFDFF")
End Function
Public Function CLR_BASLIK_ZEMIN() As Long
    CLR_BASLIK_ZEMIN = Hx("F6F8FC")
End Function
Public Function CLR_CIZGI() As Long
    CLR_CIZGI = Hx("EAEEF4")
End Function
Public Function CLR_METIN() As Long
    CLR_METIN = Hx("1B2230")
End Function
Public Function CLR_METIN_SOLUK() As Long
    CLR_METIN_SOLUK = Hx("7C8598")
End Function
Public Function CLR_HAFTASONU() As Long
    CLR_HAFTASONU = Hx("E9EDF4")
End Function
Public Function CLR_TATIL() As Long
    CLR_TATIL = Hx("D9DFE9")
End Function
Public Function CLR_BUGUN_ZEMIN() As Long
    CLR_BUGUN_ZEMIN = Hx("FEF2F2")
End Function
Public Function CLR_BUGUN_CIZGI() As Long
    CLR_BUGUN_CIZGI = Hx("DC2626")
End Function
Public Function CLR_KOYU_BANT() As Long
    CLR_KOYU_BANT = Hx("111827")
End Function
Public Function CLR_BEYAZ() As Long
    CLR_BEYAZ = Hx("FFFFFF")
End Function
' Durum renkleri
Public Function CLR_TAMAM() As Long
    CLR_TAMAM = Hx("16A34A")
End Function
Public Function CLR_DEVAM() As Long
    CLR_DEVAM = Hx("F59E0B")
End Function
Public Function CLR_PLAN() As Long
    CLR_PLAN = Hx("94A3B8")
End Function
Public Function CLR_GECIKME() As Long
    CLR_GECIKME = Hx("DC2626")
End Function
' Durum koduna gore bar rengi
Public Function DurumRengi(ByVal durum As String) As Long
    Select Case LCase$(durum)
        Case DURUM_TAMAM: DurumRengi = CLR_TAMAM()
        Case DURUM_DEVAM: DurumRengi = CLR_DEVAM()
        Case Else:        DurumRengi = CLR_PLAN()
    End Select
End Function

'--- Durum kodlari -------------------------------------------------------------
Public Const DURUM_PLAN         As String = "plan"
Public Const DURUM_DEVAM        As String = "devam"
Public Const DURUM_TAMAM        As String = "tamam"

'--- Rol kodlari ---------------------------------------------------------------
Public Const ROL_ADMIN          As String = "admin"
Public Const ROL_KULLANICI      As String = "kullanici"

'--- Yetki anahtarlari (_Kullanicilar sayfasindaki sutunlara karsilik gelir) ---
Public Const YETKI_GOREV_EKLE   As String = "GorevEkle"
Public Const YETKI_YORUM        As String = "Yorum"
Public Const YETKI_GECIKME      As String = "Gecikme"
Public Const YETKI_KULLANICI    As String = "KullaniciYonet"

'--- Ayar anahtarlari (_Ayarlar sayfasi) ---------------------------------------
Public Const AYAR_AKTIF_YIL     As String = "AktifYil"
Public Const AYAR_AKTIF_AY      As String = "AktifAy"
Public Const AYAR_OTURUM        As String = "OturumKullaniciID"
Public Const AYAR_SON_GOREV_ID  As String = "SonGorevID"
Public Const AYAR_SON_YORUM_ID  As String = "SonYorumID"
Public Const AYAR_SON_GECIKME   As String = "SonGecikmeID"
Public Const AYAR_SON_KUL_ID    As String = "SonKullaniciID"

'--- Sekil adlandirma onekleri -------------------------------------------------
' Cizim yenilenirken bu oneklerle baslayan sekiller silinir. Kalici arayuz
' sekilleri (butonlar) UI_ onekini kullanir ve silinmez.
Public Const SP_BAR             As String = "gx_bar_"
Public Const SP_GECIKME         As String = "gx_slip_"
Public Const SP_HALKA           As String = "gx_ring_"
Public Const SP_BUGUN           As String = "gx_today"
Public Const SP_PROJE           As String = "gx_pspan_"
Public Const SP_AVATAR          As String = "gx_av_"
Public Const SP_ROZET           As String = "gx_badge_"
Public Const SP_UI              As String = "UI_"
Public Const SP_MODAL           As String = "MD_"

'--- Gun adlari ----------------------------------------------------------------
Public Function GunHarfi(ByVal haftaninGunu As Long) As String
    ' vbMonday=2 ... Excel Weekday(d, vbMonday) 1..7 dondurur
    Select Case haftaninGunu
        Case 1: GunHarfi = "P"   ' Pazartesi
        Case 2: GunHarfi = "S"   ' Sali
        Case 3: GunHarfi = "C"   ' Carsamba
        Case 4: GunHarfi = "P"   ' Persembe
        Case 5: GunHarfi = "C"   ' Cuma
        Case 6: GunHarfi = "C"   ' Cumartesi
        Case 7: GunHarfi = "P"   ' Pazar
        Case Else: GunHarfi = "?"
    End Select
End Function

Public Function AyAdi(ByVal ay As Long) As String
    Dim a As Variant
    a = Array("OCAK", "SUBAT", "MART", "NISAN", "MAYIS", "HAZIRAN", _
              "TEMMUZ", "AGUSTOS", "EYLUL", "EKIM", "KASIM", "ARALIK")
    If ay >= 1 And ay <= 12 Then AyAdi = a(ay - 1) Else AyAdi = "?"
End Function
