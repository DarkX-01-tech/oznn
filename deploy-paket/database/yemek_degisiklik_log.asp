<%
Function LogSafeStr(val)
    If IsNull(val) Or IsEmpty(val) Then
        LogSafeStr = ""
    Else
        LogSafeStr = Trim(val & "")
    End If
End Function

Function LogHtmlSafe(val)
    Dim s
    s = LogSafeStr(val)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    LogHtmlSafe = s
End Function

Function LogJsonStr(val)
    Dim s
    s = LogSafeStr(val)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(10), "")
    LogJsonStr = s
End Function

Function LogSqlStr(val)
    LogSqlStr = Replace(LogSafeStr(val), "'", "''")
End Function

Function LogMenuTarihStr(tarihVal)
    If IsNull(tarihVal) Or tarihVal = "" Then
        LogMenuTarihStr = ""
    Else
        LogMenuTarihStr = Month(tarihVal) & "/" & Day(tarihVal) & "/" & Year(tarihVal)
    End If
End Function

Sub LogYemekDegisiklik(menuTarih, ogunKodu, eskiDeger, yeniDeger, kullanici, menuTipi, yil, ay)
    Dim eski, yeni, menuTarihSql
    eski = LogSafeStr(eskiDeger)
    yeni = LogSafeStr(yeniDeger)
    If eski = yeni Then Exit Sub

    menuTarihSql = LogMenuTarihStr(menuTarih)
    If menuTarihSql = "" Then Exit Sub

    On Error Resume Next
    Dim sqlDeg
    sqlDeg = "INSERT INTO yemek_degisiklik_log (menu_tarih, degisiklik_tarihi, ogun_tipi, eski_deger, yeni_deger, kullanici, menu_tipi, yil, ay) VALUES (" & _
             "#" & menuTarihSql & "#, Now(), '" & LogSqlStr(ogunKodu) & "', '" & LogSqlStr(eski) & "', '" & LogSqlStr(yeni) & "', '" & LogSqlStr(kullanici) & "', '" & LogSqlStr(menuTipi) & "', " & CInt(yil) & ", " & CInt(ay) & ")"
    ConnYemek.Execute sqlDeg
    Err.Clear
    On Error GoTo 0
End Sub

Sub LogKayitOgunDegisiklikleri(menuTarih, eskiRs, yeniOgleCorba, yeniOgleAna, yeniOgleYan, yeniOgleTatli, yeniAksamCorba, yeniAksamAna, yeniAksamYan, yeniAksamTatli, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "ogle_corba", eskiRs("ogle_corba"), yeniOgleCorba, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "ogle_ana", eskiRs("ogle_ana_yemek"), yeniOgleAna, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "ogle_yan", eskiRs("ogle_yan_urun"), yeniOgleYan, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "ogle_tatli", eskiRs("ogle_tatli"), yeniOgleTatli, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "aksam_corba", eskiRs("aksam_corba"), yeniAksamCorba, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "aksam_ana", eskiRs("aksam_ana_yemek"), yeniAksamAna, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "aksam_yan", eskiRs("aksam_yan_urun"), yeniAksamYan, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "aksam_tatli", eskiRs("aksam_tatli"), yeniAksamTatli, kullanici, menuTipi, yil, ay)
End Sub

Sub LogGuncellemeOturumu(yil, ay, guncellemeTipi, kullanici, aciklama)
    On Error Resume Next
    Dim sqlGunc
    sqlGunc = "INSERT INTO yemek_guncelleme_log (yil, ay, guncelleme_tarihi, guncelleme_tipi, kullanici, aciklama) VALUES (" & _
              CInt(yil) & ", " & CInt(ay) & ", Now(), '" & LogSqlStr(guncellemeTipi) & "', '" & LogSqlStr(kullanici) & "', '" & LogSqlStr(aciklama) & "')"
    ConnYemek.Execute sqlGunc
    Err.Clear
    On Error GoTo 0
End Sub

Function OgunTipiOgleMi(ogunKodu)
    Dim k
    k = LCase(LogSafeStr(ogunKodu))
    If Left(k, 5) = "ogle_" Or Left(k, 4) = "ogle" Then
        OgunTipiOgleMi = True
    ElseIf InStr(k, "&#246;&#287;le") > 0 Or InStr(k, "&#214;&#287;le") > 0 Then
        OgunTipiOgleMi = True
    Else
        OgunTipiOgleMi = False
    End If
End Function

Function GetOgunTipiEtiket(ogunKodu)
    Select Case LCase(LogSafeStr(ogunKodu))
        Case "ogle_corba": GetOgunTipiEtiket = "&#214;&#287;le - &#199;orba"
        Case "ogle_ana": GetOgunTipiEtiket = "&#214;&#287;le - Ana Yemek"
        Case "ogle_yan": GetOgunTipiEtiket = "&#214;&#287;le - Yan &#220;r&#252;n"
        Case "ogle_tatli": GetOgunTipiEtiket = "&#214;&#287;le - Tatl&#305;"
        Case "aksam_corba": GetOgunTipiEtiket = "Ak&#351;am - &#199;orba"
        Case "aksam_ana": GetOgunTipiEtiket = "Ak&#351;am - Ana Yemek"
        Case "aksam_yan": GetOgunTipiEtiket = "Ak&#351;am - Yan &#220;r&#252;n"
        Case "aksam_tatli": GetOgunTipiEtiket = "Ak&#351;am - Tatl&#305;"
        Case Else: GetOgunTipiEtiket = ogunKodu
    End Select
End Function

Function EntityDecode(val)
    Dim s
    s = LogSafeStr(val)
    s = Replace(s, "&#351;", ChrW(351))
    s = Replace(s, "&#350;", ChrW(350))
    s = Replace(s, "&#287;", ChrW(287))
    s = Replace(s, "&#286;", ChrW(286))
    s = Replace(s, "&#252;", ChrW(252))
    s = Replace(s, "&#220;", ChrW(220))
    s = Replace(s, "&#246;", ChrW(246))
    s = Replace(s, "&#214;", ChrW(214))
    s = Replace(s, "&#231;", ChrW(231))
    s = Replace(s, "&#199;", ChrW(199))
    s = Replace(s, "&#305;", ChrW(305))
    s = Replace(s, "&#304;", ChrW(304))
    EntityDecode = s
End Function

Function GuncellemeMetinGoster(val)
    GuncellemeMetinGoster = LogHtmlSafe(EntityDecode(val))
End Function

Function GuncellemeTipEtiket(tip)
    Dim t
    t = LCase(EntityDecode(LogSafeStr(tip)))
    If InStr(t, "toplu") > 0 Or InStr(t, "ayl") > 0 Then
        If InStr(t, "diyet") > 0 Then
            GuncellemeTipEtiket = "Ayl&#305;k Liste G&#252;ncellenmesi (Diyet)"
        Else
            GuncellemeTipEtiket = "Ayl&#305;k Liste G&#252;ncellenmesi"
        End If
    ElseIf InStr(t, "tek") > 0 Or InStr(t, "tekil") > 0 Or InStr(t, "g&#252;nl") > 0 Then
        If InStr(t, "diyet") > 0 Then
            GuncellemeTipEtiket = "G&#252;nl&#252;k Liste G&#252;ncellenmesi (Diyet)"
        Else
            GuncellemeTipEtiket = "G&#252;nl&#252;k Liste G&#252;ncellenmesi"
        End If
    Else
        GuncellemeTipEtiket = GuncellemeMetinGoster(tip)
    End If
End Function

Function GuncellemeOzetMetni(tip, aciklama, yil, ayNo, ayAdiStr)
    Dim t, acik, ayEtiket
    t = LCase(EntityDecode(LogSafeStr(tip)))
    acik = EntityDecode(LogSafeStr(aciklama))
    If ayAdiStr <> "" Then
        ayEtiket = ayAdiStr
    Else
        ayEtiket = GetMonthName(CInt(ayNo))
    End If
    If InStr(t, "toplu") > 0 Or InStr(t, "ayl") > 0 Then
        GuncellemeOzetMetni = ayEtiket & " " & yil & " ay&#305;n&#305;n men&#252; listesi g&#252;ncellendi"
        If acik <> "" Then GuncellemeOzetMetni = GuncellemeOzetMetni & " (" & GuncellemeMetinGoster(aciklama) & ")"
    ElseIf acik <> "" Then
        GuncellemeOzetMetni = GuncellemeMetinGoster(aciklama)
    Else
        GuncellemeOzetMetni = GuncellemeTipEtiket(tip)
    End If
End Function

Function GuncellemeTipiGoster(rsG)
    GuncellemeTipiGoster = GuncellemeTipEtiket(rsG("guncelleme_tipi"))
End Function
%>
