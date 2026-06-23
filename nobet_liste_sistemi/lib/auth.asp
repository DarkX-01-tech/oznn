<%
Sub AdminGirisGerekli()
    If Session(SESSION_ADMIN_KEY) <> "1" Then
        Response.Redirect MODUL_WEB_YOLU & "admin/login.asp"
        Response.End
    End If
End Sub

Function AdminGirisYap(kullaniciAdi, sifre)
    On Error Resume Next
    AdminGirisYap = False

    If Not IsObject(conn) Then Exit Function

    Dim rs, sql
    sql = "SELECT TOP 1 id, kullanici_adi, ad_soyad FROM NobetAdminKullanicilar " & _
          "WHERE kullanici_adi = '" & SqlEscape(kullaniciAdi) & "' " & _
          "AND sifre = '" & SqlEscape(sifre) & "' AND aktif = 1"

    Set rs = conn.Execute(sql)
    If Not rs.EOF Then
        Session(SESSION_ADMIN_KEY) = "1"
        Session(SESSION_ADMIN_KEY & "_id") = rs("id")
        Session(SESSION_ADMIN_KEY & "_kullanici") = rs("kullanici_adi")
        If Not IsNull(rs("ad_soyad")) Then
            Session(SESSION_ADMIN_KEY & "_ad") = rs("ad_soyad")
        Else
            Session(SESSION_ADMIN_KEY & "_ad") = rs("kullanici_adi")
        End If
        AdminGirisYap = True
    End If

    If IsObject(rs) Then
        rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0
End Function

Sub AdminCikisYap()
    Session(SESSION_ADMIN_KEY) = ""
    Session(SESSION_ADMIN_KEY & "_id") = ""
    Session(SESSION_ADMIN_KEY & "_kullanici") = ""
    Session(SESSION_ADMIN_KEY & "_ad") = ""
End Sub
%>
