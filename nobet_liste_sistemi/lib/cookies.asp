<%
Sub NobetCookieYaz(anahtar, deger)
    Response.Cookies(anahtar) = deger
    Response.Cookies(anahtar).Path = MODUL_WEB_YOLU
    Response.Cookies(anahtar).Expires = DateAdd("d", 1, Now())
End Sub

Sub NobetCookieSil(anahtar)
    Response.Cookies(anahtar) = ""
    Response.Cookies(anahtar).Path = MODUL_WEB_YOLU
    Response.Cookies(anahtar).Expires = DateAdd("d", -1, Now())
End Sub

Function NobetCookieOku(anahtar)
    If Request.Cookies(anahtar) <> "" Then
        NobetCookieOku = Request.Cookies(anahtar)
    Else
        NobetCookieOku = ""
    End If
End Function
%>
