<%
Sub NobetYanitKodSayfasiAyarla()
    Response.CodePage = NOBET_KOD_SAYFASI
    Response.CharSet = "utf-8"
    Response.ContentType = "text/html; charset=utf-8"
End Sub

Sub PortalOturumKodSayfasiSifirla()
    On Error Resume Next
    Session.CodePage = PORTAL_KOD_SAYFASI
    Session.LCID = 1055
    On Error GoTo 0
End Sub

Sub NobetIstekBaslat()
    Call PortalOturumKodSayfasiSifirla()
    Call NobetYanitKodSayfasiAyarla()
End Sub

Sub NobetIstekBitir()
    Call PortalOturumKodSayfasiSifirla()
End Sub

Sub NobetYonlendir(hedef)
    Call PortalOturumKodSayfasiSifirla()
    Response.Redirect hedef
    Response.End
End Sub
%>
