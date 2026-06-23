<%
Sub NobetYanitKodSayfasiAyarla()
    Response.CodePage = NOBET_KOD_SAYFASI
    Response.CharSet = "utf-8"
    Response.ContentType = "text/html; charset=utf-8"
End Sub

Sub PortalOturumKodSayfasiSifirla()
    On Error Resume Next
    Session.CodePage = PORTAL_KOD_SAYFASI
    On Error GoTo 0
End Sub
%>
