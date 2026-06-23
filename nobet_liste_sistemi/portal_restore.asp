<%@ Language=VBScript CodePage=1254 %>
<%
On Error Resume Next
Session.CodePage = 1254
Session.LCID = 1055
Response.CodePage = 1254
Response.CharSet = "windows-1254"
On Error GoTo 0
Response.Redirect "http://10.201.65.10/"
Response.End
%>
