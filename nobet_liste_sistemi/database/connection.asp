<!-- #include file="../lib/database_setup.asp" -->
<%
Dim conn

Call EnsureAccessDatabase()
Call EnsureAccessTables()

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open AccessBaglantiMetni()
%>
