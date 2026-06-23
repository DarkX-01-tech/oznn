<%@ Language=VBScript CodePage=1254 %>
<%
Response.CodePage = 1254
Response.CharSet = "windows-1254"
%>
<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../database/connection.asp" -->
<!-- #include file="../lib/config.asp" -->
<!-- #include file="../lib/functions.asp" -->
<!-- #include file="../lib/auth.asp" -->
<!-- #include file="../lib/upload.asp" -->
<%
AdminGirisGerekli

Dim binaKodu, dosyaAdi, baslik, hata, hedefYol, islemTipi, seciliYil, seciliAy, yilParam, ayParam
binaKodu = Trim(Request("bina"))
dosyaAdi = Trim(Request("dosya"))
hata = ""

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
EnsureAyKlasoru binaKodu, seciliYil, seciliAy
hedefYol = NobetDosyaFizikselYolu(binaKodu, dosyaAdi, seciliYil, seciliAy)
islemTipi = "yukle"

If NobetDosyaMevcut(binaKodu, dosyaAdi, seciliYil, seciliAy) Then
    islemTipi = "guncelle"
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim yuklemeHatasi, basarili, yukleyen
    basarili = DosyaYukleVeKaydet(hedefYol, yuklemeHatasi)

    If basarili Then
        yukleyen = NobetCookieOku(SESSION_ADMIN_KEY & "_ad")
        NobetDosyaDbKaydet binaKodu, dosyaAdi, baslik, yukleyen, True, seciliYil, seciliAy, islemTipi
        If islemTipi = "guncelle" Then
            Response.Redirect "listeler.asp?mesaj=guncellendi"
        Else
            Response.Redirect "listeler.asp?mesaj=yuklendi"
        End If
        Response.End
    Else
        hata = yuklemeHatasi
    End If
End If

Dim qs
qs = "bina=" & Server.URLEncode(binaKodu) & "&dosya=" & Server.URLEncode(dosyaAdi) & "&yil=" & seciliYil & "&ay=" & Server.URLEncode(seciliAy)
%>
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta http-equiv="Content-Language" content="tr">
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title><% If islemTipi = "guncelle" Then %>Dosya Güncelle<% Else %>Dosya Yükle<% End If %></title>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body class="admin-body">
  <div class="page-shell">
    <div class="admin-wrap upload-card">
      <div class="admin-topbar">
        <div>
          <p class="eyebrow">Dosya İşlemleri</p>
          <h1><% If islemTipi = "guncelle" Then %>Dosya Güncelle<% Else %>Dosya Yükle<% End If %></h1>
        </div>
        <div class="admin-nav">
          <a href="listeler.asp" class="btn btn-secondary btn-sm">Listelere Dön</a>
        </div>
      </div>

      <% If hata <> "" Then %>
        <div class="alert alert-error"><%= Server.HTMLEncode(hata) %></div>
      <% End If %>

      <div class="alert alert-success">Seçtiğiniz dosya, sistem adıyla kaydedilecektir: <strong><%= Server.HTMLEncode(dosyaAdi) %></strong></div>

      <div class="info-grid">
        <div class="info-item">
          <span class="info-label">Bina</span>
          <span class="info-value"><%= Server.HTMLEncode(BinaAdiGoster(binaKodu)) %></span>
        </div>
        <div class="info-item">
          <span class="info-label">Liste</span>
          <span class="info-value"><%= Server.HTMLEncode(baslik) %></span>
        </div>
        <div class="info-item">
          <span class="info-label">Yıl</span>
          <span class="info-value"><%= seciliYil %></span>
        </div>
        <div class="info-item">
          <span class="info-label">Ay</span>
          <span class="info-value"><%= AyBaslikFromKlasor(seciliAy) %></span>
        </div>
        <div class="info-item full">
          <span class="info-label">Kayıt Adı</span>
          <span class="info-value mono"><%= Server.HTMLEncode(dosyaAdi) %></span>
        </div>
      </div>

      <form class="upload-form" method="post" enctype="multipart/form-data" action="yukle.asp?<%= qs %>">
        <div class="form-group">
          <label for="dosya">PDF / Excel Dosyası (herhangi bir ad olabilir)</label>
          <input class="file-input" type="file" id="dosya" name="dosya" accept=".pdf,.xls,.xlsx" required>
        </div>
        <div class="form-actions">
          <button type="submit" class="btn btn-primary"><% If islemTipi = "guncelle" Then %>Güncelle<% Else %>Yükle<% End If %></button>
          <a href="listeler.asp" class="btn btn-secondary">İptal</a>
        </div>
      </form>
    </div>
  </div>
</body>
</html>
