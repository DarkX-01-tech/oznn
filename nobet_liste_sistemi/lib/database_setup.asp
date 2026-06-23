<%
Sub EnsureAccessDatabase()
    Dim fso, dbPath, dbKlasor, cat

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    dbPath = AccessDbFizikselYol()
    dbKlasor = fso.GetParentFolderName(dbPath)

    If Not fso.FolderExists(dbKlasor) Then
        fso.CreateFolder dbKlasor
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
            "guncelleme_tarihi DATETIME)"
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
            "VALUES ('admin', 'admin123', 'Sistem Yöneticisi', True, Now())"
    End If

    dbConn.Close
    Set dbConn = Nothing
End Sub
%>
