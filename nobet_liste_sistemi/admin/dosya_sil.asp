<%@ Language=VBScript CodePage=65001 %>
<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../database/connection.asp" -->
<!-- #include file="../lib/config.asp" -->
<!-- #include file="../lib/functions.asp" -->
<!-- #include file="../lib/auth.asp" -->
<%
AdminGirisGerekli

Dim binaKodu, dosyaAdi, baslik, fso, hedefYol, seciliYil, seciliAy, yilParam, ayParam
binaKodu = Trim(Request("bina"))
dosyaAdi = Trim(Request("dosya"))

yilParam = Trim(Request("yil"))
ayParam = LCase(Trim(Request("ay")))
If yilParam <> "" And IsNumeric(yilParam) Then
    seciliYil = CInt(yilParam)
Else
    seciliYil = GetSeciliYil()
End If
If ayParam <> "" And AyKlasorGecerliMi(ayParam) Then
    seciliAy = ayParam
Else
    seciliAy = GetSeciliAyKlasor()
End If

If binaKodu = "" Or dosyaAdi = "" Then
    Response.Redirect "listeler.asp"
    Response.End
End If

If Not DosyaAdiGecerliMi(binaKodu, dosyaAdi) Then
    Response.Write "Geçersiz dosya tanımı."
    Response.End
End If

baslik = BaslikGetir(binaKodu, dosyaAdi)
hedefYol = NobetDosyaFizikselYolu(binaKodu, dosyaAdi, seciliYil, seciliAy)

Set fso = Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(hedefYol) Then
    fso.DeleteFile hedefYol, True
End If
Set fso = Nothing

NobetDosyaDbKaydet binaKodu, dosyaAdi, baslik, Session(SESSION_ADMIN_KEY & "_ad"), False, seciliYil, seciliAy, "sil"
Response.Redirect "listeler.asp?mesaj=silindi"
%>
