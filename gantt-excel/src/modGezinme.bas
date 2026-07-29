Attribute VB_Name = "modGezinme"
'===============================================================================
' modGezinme - Ay gecisleri ve sayfa olaylarinin karsiladigi eylemler
'===============================================================================
Option Explicit

Public Sub OncekiAy()
    AyKaydir -1
End Sub

Public Sub SonrakiAy()
    AyKaydir 1
End Sub

Public Sub BugunAyi()
    modVeri.AyarYaz AYAR_AKTIF_YIL, Year(Date)
    modVeri.AyarYaz AYAR_AKTIF_AY, Month(Date)
    modCizelge.Ciz
End Sub

Private Sub AyKaydir(ByVal adim As Long)
    Dim yil As Long, ay As Long, d As Date
    yil = CLng(Val(modVeri.Ayar(AYAR_AKTIF_YIL, Year(Date))))
    ay = CLng(Val(modVeri.Ayar(AYAR_AKTIF_AY, Month(Date))))
    d = DateAdd("m", adim, DateSerial(yil, ay, 1))
    modVeri.AyarYaz AYAR_AKTIF_YIL, Year(d)
    modVeri.AyarYaz AYAR_AKTIF_AY, Month(d)
    modCizelge.Ciz
End Sub

'===============================================================================
' Sayfa olaylarindan cagrilir (Cizelge sayfasinin kod modulu)
'===============================================================================

' Gun basligina cift tiklandi -> o gunu tatil yap / tatili kaldir
Public Function GunBasligiTiklandi(ByVal hedef As Range) As Boolean
    Dim yil As Long, ay As Long, gun As Long, t As Date

    If hedef.Row <> R_GUNNO Then Exit Function
    If hedef.Column < C_GUN1 Then Exit Function

    gun = hedef.Column - C_GUN1 + 1
    yil = CLng(Val(modVeri.Ayar(AYAR_AKTIF_YIL, Year(Date))))
    ay = CLng(Val(modVeri.Ayar(AYAR_AKTIF_AY, Month(Date))))
    If gun < 1 Or gun > Day(DateSerial(yil, ay + 1, 0)) Then Exit Function

    t = DateSerial(yil, ay, gun)

    Dim mesaj As String
    If modVeri.OffGunMu(t) Then
        mesaj = Format$(t, "dd.mm.yyyy") & " tatil isareti KALDIRILSIN mi?"
    Else
        mesaj = Format$(t, "dd.mm.yyyy") & " idari tatil olarak isaretlensin mi?"
    End If

    If MsgBox(mesaj, vbQuestion + vbYesNo, APP_ADI) = vbYes Then
        modVeri.OffGunDegistir t
        modCizelge.Ciz
    End If

    GunBasligiTiklandi = True
End Function

' Gun hucresine cift tiklandi -> o tarihte yeni gorev ekle
Public Function GunHucresiTiklandi(ByVal hedef As Range) As Boolean
    Dim yil As Long, ay As Long, gun As Long

    If hedef.Row < R_VERI_BAS Then Exit Function
    If hedef.Column < C_GUN1 Or hedef.Column > C_GUN1 + GUN_SAYISI_MAX - 1 Then Exit Function

    gun = hedef.Column - C_GUN1 + 1
    yil = CLng(Val(modVeri.Ayar(AYAR_AKTIF_YIL, Year(Date))))
    ay = CLng(Val(modVeri.Ayar(AYAR_AKTIF_AY, Month(Date))))
    If gun < 1 Or gun > Day(DateSerial(yil, ay + 1, 0)) Then Exit Function

    modGorev.GorevEkleAcTarihle DateSerial(yil, ay, gun)
    GunHucresiTiklandi = True
End Function
