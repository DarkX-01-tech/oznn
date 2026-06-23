<%@ Language=VBScript CodePage=1254 %>
<%
Response.CodePage = 1254
Response.CharSet = "windows-1254"
%>
<!-- #include file="../lib/constants.asp" -->
<!-- #include file="../lib/cookies.asp" -->
<!-- #include file="../lib/auth.asp" -->
<%
AdminCikisYap
Response.Redirect "login.asp"
%>
