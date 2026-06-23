<%@ Language=VBScript CodePage=65001 %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType = "text/html; charset=utf-8"
%>
<!-- #include file="../lib/constants.asp" -->
<!-- #include file="../lib/cookies.asp" -->
<!-- #include file="../lib/auth.asp" -->
<%
AdminCikisYap
Response.Redirect "login.asp"
%>
