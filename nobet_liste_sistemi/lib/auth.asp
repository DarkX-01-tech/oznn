<!-- #include file="cookies.asp" -->
<%
Sub AdminGirisGerekli()
    If NobetCookieOku(SESSION_ADMIN_KEY) <> "1" Then
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
          "AND sifre = '" & SqlEscape(sifre) & "' AND aktif = True"

    Set rs = conn.Execute(sql)
    If Not rs.EOF Then
        NobetCookieYaz SESSION_ADMIN_KEY, "1"
        NobetCookieYaz SESSION_ADMIN_KEY & "_id", CStr(rs("id"))
        NobetCookieYaz SESSION_ADMIN_KEY & "_kullanici", rs("kullanici_adi")
        If Not IsNull(rs("ad_soyad")) Then
            NobetCookieYaz SESSION_ADMIN_KEY & "_ad", rs("ad_soyad")
        Else
            NobetCookieYaz SESSION_ADMIN_KEY & "_ad", rs("kullanici_adi")
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
    NobetCookieSil SESSION_ADMIN_KEY
    NobetCookieSil SESSION_ADMIN_KEY & "_id"
    NobetCookieSil SESSION_ADMIN_KEY & "_kullanici"
    NobetCookieSil SESSION_ADMIN_KEY & "_ad"
End Sub
%>
