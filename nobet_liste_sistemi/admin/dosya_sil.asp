<%@ Language=VBScript CodePage=65001 %>
<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../database/connection.asp" -->
<!-- #include file="../lib/config.asp" -->
<!-- #include file="../lib/functions.asp" -->
<!-- #include file="../lib/auth.asp" -->
<%
AdminGirisGerekli

Dim binaKodu, dosyaAdi, baslik, fso, hedefYol
binaKodu = Trim(Request("bina"))
dosyaAdi = Trim(Request("dosya"))

If binaKodu = "" Or dosyaAdi = "" Then
    Response.Redirect "panel.asp"
    Response.End
End If

If Not DosyaAdiGecerliMi(binaKodu, dosyaAdi) Then
    Response.Write "Geçersiz dosya tanımı."
    Response.End
End If

baslik = BaslikGetir(binaKodu, dosyaAdi)
hedefYol = NobetDosyaFizikselYolu(binaKodu, dosyaAdi)

Set fso = Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(hedefYol) Then
    fso.DeleteFile hedefYol, True
End If
Set fso = Nothing

NobetDosyaDbKaydet binaKodu, dosyaAdi, baslik, Session(SESSION_ADMIN_KEY & "_ad"), False
Response.Redirect "panel.asp?mesaj=silindi"
%>
