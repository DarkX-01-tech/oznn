<!-- #include file="includes/charset.asp" -->
<%
' ============================================================
' Nöbet Liste Sistemi - Genel Ayarlar
' Kurulumda MODUL_WEB_YOLU değerini kontrol edin.
' ============================================================

Const MODUL_WEB_YOLU = "/Admin/nobet_liste_sistemi/"
Const BINA_PENDIK = "pendik"
Const BINA_BASIBUYUK = "basibuyuk"
Const LISTELER_KLASORU = "listeler"
Const ACCESS_DB_DOSYA = "nobet_liste.mdb"
Const SESSION_ADMIN_KEY = "nobet_liste_admin"
Const SESSION_SECILI_YIL = "nobet_secili_yil"
Const SESSION_SECILI_AY = "nobet_secili_ay"
Const PORTAL_IMAGES_YOLU = "/Admin/images/"
Const PORTAL_BANNER_RESIM = "muhst_06.png"
Const PORTAL_ARKAPLAN_RESIM = "arka-plan.png"
Const PORTAL_ANA_SAYFA = "http://10.201.65.10/"

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
