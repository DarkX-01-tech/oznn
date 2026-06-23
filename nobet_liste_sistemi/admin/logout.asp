<%@ Language=VBScript CodePage=1254 %>
<!-- #include file="../lib/constants.asp" -->
<!-- #include file="../lib/encoding.asp" -->
<!-- #include file="../lib/auth.asp" -->
<%
Call PortalOturumKodSayfasiSifirla()
AdminCikisYap
Response.Redirect "login.asp"
%>
