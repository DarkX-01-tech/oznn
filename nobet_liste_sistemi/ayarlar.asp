<%
' ============================================================
' Nöbet Liste Sistemi - Genel Ayarlar
' Sunucuya kurulumda MODUL_WEB_YOLU değerini kontrol edin.
' ============================================================

Const MODUL_WEB_YOLU = "/Admin/nobet_liste_sistemi/"
Const BINA_PENDIK = "pendik"
Const BINA_BASIBUYUK = "basibuyuk"
Const LISTELER_KLASORU = "listeler"
Const SESSION_ADMIN_KEY = "nobet_liste_admin"

Function ModulFizikselYol()
    ModulFizikselYol = Server.MapPath(MODUL_WEB_YOLU)
End Function

Function ListelerKokYolu()
    ListelerKokYolu = ModulFizikselYol() & "\" & LISTELER_KLASORU
End Function

Function SqlEscape(deger)
    SqlEscape = Replace(CStr(deger), "'", "''")
End Function
%>
