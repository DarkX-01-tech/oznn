<%
' Gerekli degiskenler: secilen_yil, secilen_ay, menu_tipi
Dim rsDegisiklik, degisiklikVar, degisiklikHata
degisiklikVar = False
degisiklikHata = False
Dim degisiklikSayisi, degisenMenuGun, ogleDegisim, aksamDegisim
degisiklikSayisi = 0
degisenMenuGun = 0
ogleDegisim = 0
aksamDegisim = 0

On Error Resume Next
Set rsDegisiklik = ConnYemek.Execute("SELECT * FROM yemek_degisiklik_log WHERE yil = " & secilen_yil & " AND ay = " & secilen_ay & " AND menu_tipi = '" & menu_tipi & "' ORDER BY degisiklik_tarihi DESC")
If Err.Number <> 0 Then
    degisiklikHata = True
Else
    degisiklikVar = True
    If Not rsDegisiklik.EOF Then
        Dim rsDegOzet
        Set rsDegOzet = ConnYemek.Execute("SELECT COUNT(*) AS toplam, COUNT(DISTINCT menu_tarih) AS gun FROM yemek_degisiklik_log WHERE yil = " & secilen_yil & " AND ay = " & secilen_ay & " AND menu_tipi = '" & menu_tipi & "'")
        If Not rsDegOzet.EOF Then
            degisiklikSayisi = rsDegOzet("toplam")
            degisenMenuGun = rsDegOzet("gun")
        End If
        rsDegOzet.Close: Set rsDegOzet = Nothing

        Dim rsOgleAksam
        Set rsOgleAksam = ConnYemek.Execute("SELECT ogun_tipi FROM yemek_degisiklik_log WHERE yil = " & secilen_yil & " AND ay = " & secilen_ay & " AND menu_tipi = '" & menu_tipi & "'")
        Do While Not rsOgleAksam.EOF
            If OgunTipiOgleMi(rsOgleAksam("ogun_tipi")) Then
                ogleDegisim = ogleDegisim + 1
            Else
                aksamDegisim = aksamDegisim + 1
            End If
            rsOgleAksam.MoveNext
        Loop
        rsOgleAksam.Close: Set rsOgleAksam = Nothing
    End If
End If
Err.Clear

Dim guncellemeHata, guncellemeSayisi
guncellemeHata = False
guncellemeSayisi = 0
Dim guncellemeFiltre
If menu_tipi = "diyet" Then
    guncellemeFiltre = " AND (guncelleme_tipi LIKE '%Diyet%' OR aciklama LIKE '%Diyet%')"
Else
    guncellemeFiltre = " AND ((guncelleme_tipi NOT LIKE '%Diyet%' OR guncelleme_tipi IS NULL) AND (aciklama NOT LIKE '%Diyet%' OR aciklama IS NULL))"
End If

Dim rsGuncSay
Set rsGuncSay = ConnYemek.Execute("SELECT COUNT(*) AS adet FROM yemek_guncelleme_log WHERE yil = " & secilen_yil & " AND ay = " & secilen_ay & guncellemeFiltre)
If Err.Number <> 0 Then
    guncellemeHata = True
Else
    If Not rsGuncSay.EOF Then guncellemeSayisi = rsGuncSay("adet")
End If
rsGuncSay.Close: Set rsGuncSay = Nothing
Err.Clear
On Error GoTo 0
%>
