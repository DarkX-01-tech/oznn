<%
Sub EnsureAccessDatabase()
    Dim fso, dbPath, dbKlasor, cat

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    dbPath = AccessDbFizikselYol()
    dbKlasor = fso.GetParentFolderName(dbPath)

    If Not fso.FolderExists(dbKlasor) Then
        On Error Resume Next
        fso.CreateFolder dbKlasor
        Err.Clear
        On Error GoTo 0
    End If

    If Not fso.FileExists(dbPath) Then
        On Error Resume Next
        Set cat = Server.CreateObject("ADOX.Catalog")
        cat.Create AccessBaglantiMetni()

        If Err.Number <> 0 Then
            Err.Clear
            cat.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
        End If

        Set cat = Nothing
        On Error GoTo 0
    End If

    Set fso = Nothing
End Sub

Function AccessTabloVarMi(tabloAdi)
    On Error Resume Next
    Dim testConn, rs
    AccessTabloVarMi = False

    Set testConn = Server.CreateObject("ADODB.Connection")
    testConn.Open AccessBaglantiMetni()
    Set rs = testConn.Execute("SELECT COUNT(*) FROM [" & tabloAdi & "]")
    AccessTabloVarMi = (Err.Number = 0)

    If IsObject(rs) Then
        rs.Close
        Set rs = Nothing
    End If
    testConn.Close
    Set testConn = Nothing
    On Error GoTo 0
End Function

Sub EnsureAccessTables()
    Dim dbConn

    Set dbConn = Server.CreateObject("ADODB.Connection")
    dbConn.Open AccessBaglantiMetni()

    If Not AccessTabloVarMi("NobetListeDosyalar") Then
        dbConn.Execute "CREATE TABLE NobetListeDosyalar (" & _
            "id COUNTER PRIMARY KEY, " & _
            "yil INTEGER NOT NULL, " & _
            "bina TEXT(50) NOT NULL, " & _
            "ay_klasor TEXT(20) NOT NULL, " & _
            "dosya_adi TEXT(255) NOT NULL, " & _
            "baslik TEXT(255) NOT NULL, " & _
            "aktif YESNO NOT NULL, " & _
            "yukleyen TEXT(100), " & _
            "olusturma_tarihi DATETIME, " & _
            "guncelleme_tarihi DATETIME, " & _
            "guncelleme_sayisi INTEGER)"
    End If

    If Not AccessTabloVarMi("NobetAdminKullanicilar") Then
        dbConn.Execute "CREATE TABLE NobetAdminKullanicilar (" & _
            "id COUNTER PRIMARY KEY, " & _
            "kullanici_adi TEXT(50) NOT NULL, " & _
            "sifre TEXT(255) NOT NULL, " & _
            "ad_soyad TEXT(100), " & _
            "aktif YESNO NOT NULL, " & _
            "olusturma_tarihi DATETIME)"

        dbConn.Execute "INSERT INTO NobetAdminKullanicilar (kullanici_adi, sifre, ad_soyad, aktif, olusturma_tarihi) " & _
            "VALUES ('admin', 'admin123', 'Sistem Yoneticisi', True, Now())"
    End If

    dbConn.Close
    Set dbConn = Nothing

    EnsureAccessSchemaUpgrade
End Sub

Sub EnsureAccessSchemaUpgrade()
    On Error Resume Next
    Dim dbConn

    If Not AccessTabloVarMi("NobetListeDosyalar") Then Exit Sub

    Set dbConn = Server.CreateObject("ADODB.Connection")
    dbConn.Open AccessBaglantiMetni()
    dbConn.Execute "ALTER TABLE NobetListeDosyalar ADD COLUMN guncelleme_sayisi INTEGER"
    Err.Clear

    If Not AccessTabloVarMi("NobetListeIslemLog") Then
        dbConn.Execute "CREATE TABLE NobetListeIslemLog (" & _
            "id COUNTER PRIMARY KEY, " & _
            "yil INTEGER NOT NULL, " & _
            "bina TEXT(50) NOT NULL, " & _
            "ay_klasor TEXT(20) NOT NULL, " & _
            "dosya_adi TEXT(255) NOT NULL, " & _
            "baslik TEXT(255) NOT NULL, " & _
            "islem_tipi TEXT(20) NOT NULL, " & _
            "yukleyen TEXT(100), " & _
            "islem_tarihi DATETIME)"
    End If

    dbConn.Close
    Set dbConn = Nothing
    On Error GoTo 0
End Sub

Sub LogNobetIslem(yil, binaKodu, ayKlasor, dosyaAdi, baslik, islemTipi, yukleyen)
    On Error Resume Next
    If Not IsObject(conn) Then Exit Sub

    Dim sql
    sql = "INSERT INTO NobetListeIslemLog (yil, bina, ay_klasor, dosya_adi, baslik, islem_tipi, yukleyen, islem_tarihi) VALUES (" & _
          CInt(yil) & ", '" & SqlEscape(binaKodu) & "', '" & SqlEscape(ayKlasor) & "', '" & SqlEscape(dosyaAdi) & "', '" & _
          SqlEscape(baslik) & "', '" & SqlEscape(islemTipi) & "', '" & SqlEscape(yukleyen) & "', Now())"
    conn.Execute sql
    On Error GoTo 0
End Sub
%>
