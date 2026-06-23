<%@ Language=VBScript CodePage=65001 %>
<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../lib/auth.asp" -->
<%
Session.CodePage = 1254
AdminCikisYap
Response.Redirect "login.asp"
%>
