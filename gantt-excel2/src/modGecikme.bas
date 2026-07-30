Attribute VB_Name = "modGecikme"
'===============================================================================
' modGecikme - Gecikme kaydi paneli
'
' Cizelgedeki kirmizi kesikli seride tiklandiginda acilir. Normal yorum
' panelinden bilerek ayri tutulmustur: kirmizi "Kaydet" butonu ve ustte
' planlanan / gerceklesen / sapma ucglusu.
'===============================================================================
Option Explicit

' Alan numaralari - GecikmePaneliAc sirasiyla birebir ayni
Private Const A_GOREV = 1, A_OZET = 2, A_NEDEN = 3, A_SORUMLULUK = 4
Private Const A_BASLIK = 5, A_GECMIS = 6, A_NOT = 7

Public Sub SerideTiklandi()
    Dim gorevID As Long
    gorevID = modYorum.TiklananGorevID()
    If gorevID = 0 Then Exit Sub
    GecikmePaneliAc gorevID
End Sub

' Panel gercekten acildiysa True doner. Gorev kaydetme akisi bu cevaba gore
' paneli devrediyor (bkz. modGorev.GorevKaydet).
Public Function GecikmePaneliAc(ByVal gorevID As Long) As Boolean
    Dim g As TGorev, gecikme As Long
    Dim ws As Worksheet, r As Long
    Dim neden As String, sorumluluk As String

    If Not modGiris.OturumGerekli() Then Exit Function

    g = modVeri.GorevGetir(gorevID)
    If g.ID = 0 Then Exit Function

    gecikme = modVeri.GecikmeGunu(g)
    If gecikme <= 0 Then
        MsgBox "Bu gorevde gecikme yok.", vbInformation, APP_ADI
        Exit Function
    End If

    ' Varsa onceki kaydi getir
    Set ws = modVeri.Sayfa(SH_GECIKMELER)
    r = modVeri.GecikmeSatiri(gorevID)
    If r > 0 Then
        neden = CStr(ws.Cells(r, D_NEDEN).Value)
        sorumluluk = CStr(ws.Cells(r, D_SORUMLULUK).Value)
    End If
    If Len(sorumluluk) = 0 Then sorumluluk = "Ic ekip"

    modModal.Temizle
    modModal.Baglam = CStr(gorevID)

    modModal.Alan "Gorev", MT_SALTOKUR, g.Ad
    modModal.Alan "Planlanan bitis  /  Gerceklesen  /  Sapma", MT_SALTOKUR, _
        Format$(g.PlanBitis, "dd.mm.yyyy") & "        " & _
        IIf(g.GerceklesenBitis > 0, Format$(CDate(g.GerceklesenBitis), "dd.mm.yyyy"), _
            "devam ediyor") & "        +" & gecikme & " gun"
    modModal.Alan "Gecikme nedeni", MT_COKSATIR, neden
    modModal.Alan "Sorumluluk", MT_LISTE, sorumluluk, _
        "Ic ekip,Dis yuklenici,Tedarikci,Musteri,Onay sureci,Diger"

    modModal.Alan "Konusma gecmisi", MT_BASLIK
    modModal.Alan "Gecmis", MT_SALTOKUR, modYorum.YorumGecmisi(gorevID)
    modModal.Alan "Not ekle (istege bagli)", MT_COKSATIR, ""

    modModal.Goster "gecikme", ChrW(9888) & " Gecikme kaydi", "Gecikmeyi kaydet", True
    GecikmePaneliAc = True
End Function

Public Function GecikmeKaydetPanel() As Boolean
    Dim gorevID As Long, neden As String, sorumluluk As String, ekNot As String

    gorevID = CLng(Val(modModal.Baglam))
    If gorevID = 0 Then GecikmeKaydetPanel = True: Exit Function

    If Not modVeri.YetkiVar(YETKI_GECIKME) Then
        MsgBox "Gecikme kaydi girme yetkin yok.", vbExclamation, APP_ADI
        Exit Function
    End If

    neden = Trim$(modModal.Deger(A_NEDEN))
    sorumluluk = modModal.Deger(A_SORUMLULUK)
    ekNot = Trim$(modModal.Deger(A_NOT))

    If Len(neden) = 0 Then
        MsgBox "Gecikme nedenini yazmadan kaydedemezsin." & vbCrLf & _
               "Bu alan bos kalirsa cizelge neden geciktigini gostermez.", _
               vbExclamation, APP_ADI
        Exit Function
    End If

    modVeri.GecikmeKaydet gorevID, neden, sorumluluk
    If Len(ekNot) > 0 Then modVeri.YorumEkle gorevID, ekNot

    modCizelge.Ciz
    GecikmeKaydetPanel = True
End Function
