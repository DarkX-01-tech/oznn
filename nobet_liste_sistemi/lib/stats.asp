<!-- #include file="period.asp" -->
<%
Function IstatistikYillariDizisi()
    Dim yillar, rs, sql, yilDeger, sonuc(), i, sayac
    Set yillar = Server.CreateObject("Scripting.Dictionary")
    yillar.Add CStr(GuncelYil()), True

    On Error Resume Next
    If IsObject(conn) Then
        sql = "SELECT DISTINCT yil FROM NobetListeIslemLog ORDER BY yil DESC"
        Set rs = conn.Execute(sql)
        Do While Not rs.EOF
            yilDeger = CStr(rs("yil"))
            If Not yillar.Exists(yilDeger) Then yillar.Add yilDeger, True
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0

    ReDim sonuc(yillar.Count - 1)
    i = 0
    For Each yilDeger In yillar.Keys
        sonuc(i) = CInt(yilDeger)
        i = i + 1
    Next

    IstatistikYillariDizisi = sonuc
End Function

Function AyIstatistikteVarMi(yil, ayKlasor)
    Dim fso, yilKlasor, binaKlasor, ayKlasorObj, rs, sql
    AyIstatistikteVarMi = False

    On Error Resume Next
    If IsObject(conn) Then
        sql = "SELECT TOP 1 id FROM NobetListeIslemLog WHERE yil = " & CInt(yil) & " AND ay_klasor = '" & SqlEscape(ayKlasor) & "'"
        Set rs = conn.Execute(sql)
        AyIstatistikteVarMi = Not rs.EOF
        rs.Close
        Set rs = Nothing
        If AyIstatistikteVarMi Then Exit Function
    End If
    On Error GoTo 0

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    yilKlasor = ListelerKokYolu() & "\" & yil
    If fso.FolderExists(yilKlasor) Then
        For Each binaKlasor In fso.GetFolder(yilKlasor).SubFolders
            If fso.FolderExists(binaKlasor.Path & "\" & ayKlasor) Then
                AyIstatistikteVarMi = True
                Exit Function
            End If
        Next
    End If
End Function

Function DosyaKayitBilgisi(yil, binaKodu, ayKlasor, dosyaAdi)
    Dim rs, sql, bilgi(3)
    bilgi(0) = ""
    bilgi(1) = 0
    bilgi(2) = ""

    On Error Resume Next
    If Not IsObject(conn) Then
        DosyaKayitBilgisi = bilgi
        Exit Function
    End If

    sql = "SELECT TOP 1 olusturma_tarihi, guncelleme_sayisi, guncelleme_tarihi FROM NobetListeDosyalar " & _
          "WHERE yil = " & CInt(yil) & " AND bina = '" & SqlEscape(binaKodu) & "' AND ay_klasor = '" & SqlEscape(ayKlasor) & _
          "' AND dosya_adi = '" & SqlEscape(dosyaAdi) & "'"

    Set rs = conn.Execute(sql)
    If Not rs.EOF Then
        If Not IsNull(rs("olusturma_tarihi")) Then bilgi(0) = rs("olusturma_tarihi")
        If Not IsNull(rs("guncelleme_sayisi")) Then bilgi(1) = rs("guncelleme_sayisi")
        If Not IsNull(rs("guncelleme_tarihi")) Then bilgi(2) = rs("guncelleme_tarihi")
    End If
    rs.Close
    Set rs = Nothing
    On Error GoTo 0

    DosyaKayitBilgisi = bilgi
End Function

Function TarihGoster(tarihDeger)
    If IsDate(tarihDeger) Then
        TarihGoster = FormatDateTime(tarihDeger, 0)
    Else
        TarihGoster = "-"
    End If
End Function
%>
