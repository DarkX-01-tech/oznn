<%
Function LogSafeStr(val)
    If IsNull(val) Or IsEmpty(val) Then
        LogSafeStr = ""
    Else
        LogSafeStr = Trim(val & "")
    End If
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

Sub LogYemekDegisiklik(menuTarih, ogunTipi, eskiDeger, yeniDeger, kullanici, menuTipi, yil, ay)
    Dim eski, yeni, menuTarihSql
    eski = LogSafeStr(eskiDeger)
    yeni = LogSafeStr(yeniDeger)
    If eski = yeni Then Exit Sub

    menuTarihSql = LogMenuTarihStr(menuTarih)
    If menuTarihSql = "" Then Exit Sub

    On Error Resume Next
    Dim sqlDeg
    sqlDeg = "INSERT INTO yemek_degisiklik_log (menu_tarih, degisiklik_tarihi, ogun_tipi, eski_deger, yeni_deger, kullanici, menu_tipi, yil, ay) VALUES (" & _
             "#" & menuTarihSql & "#, Now(), '" & LogSqlStr(ogunTipi) & "', '" & LogSqlStr(eski) & "', '" & LogSqlStr(yeni) & "', '" & LogSqlStr(kullanici) & "', '" & LogSqlStr(menuTipi) & "', " & CInt(yil) & ", " & CInt(ay) & ")"
    ConnYemek.Execute sqlDeg
    Err.Clear
    On Error GoTo 0
End Sub

Sub LogKayitOgunDegisiklikleri(menuTarih, eskiRs, yeniOgleCorba, yeniOgleAna, yeniOgleYan, yeniOgleTatli, yeniAksamCorba, yeniAksamAna, yeniAksamYan, yeniAksamTatli, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "&#214;&#287;le - &#199;orba", eskiRs("ogle_corba"), yeniOgleCorba, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "&#214;&#287;le - Ana Yemek", eskiRs("ogle_ana_yemek"), yeniOgleAna, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "&#214;&#287;le - Yan &#220;r&#252;n", eskiRs("ogle_yan_urun"), yeniOgleYan, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "&#214;&#287;le - Tatl&#305;", eskiRs("ogle_tatli"), yeniOgleTatli, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "Ak&#351;am - &#199;orba", eskiRs("aksam_corba"), yeniAksamCorba, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "Ak&#351;am - Ana Yemek", eskiRs("aksam_ana_yemek"), yeniAksamAna, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "Ak&#351;am - Yan &#220;r&#252;n", eskiRs("aksam_yan_urun"), yeniAksamYan, kullanici, menuTipi, yil, ay)
    Call LogYemekDegisiklik(menuTarih, "Ak&#351;am - Tatl&#305;", eskiRs("aksam_tatli"), yeniAksamTatli, kullanici, menuTipi, yil, ay)
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

Function OgunTipiOgleMi(ogunTipi)
    If InStr(ogunTipi & "", "Ak&#351;am") > 0 Then
        OgunTipiOgleMi = False
    Else
        OgunTipiOgleMi = True
    End If
End Function
%>
