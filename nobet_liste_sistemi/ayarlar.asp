<!-- #include file="lib/constants.asp" -->
<!-- #include file="includes/charset.asp" -->
<%
Function ModulFizikselYol()
    ModulFizikselYol = Server.MapPath(MODUL_WEB_YOLU)
End Function

Function ListelerKokYolu()
    ListelerKokYolu = ModulFizikselYol() & "\" & LISTELER_KLASORU
End Function

Function AccessDbFizikselYol()
    AccessDbFizikselYol = ModulFizikselYol() & "\database\" & ACCESS_DB_DOSYA
End Function

Function AccessBaglantiMetni()
    AccessBaglantiMetni = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AccessDbFizikselYol() & ";"
End Function

Function SqlEscape(deger)
    SqlEscape = Replace(CStr(deger), "'", "''")
End Function
%>
