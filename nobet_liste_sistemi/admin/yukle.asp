<%@ Language=VBScript CodePage=65001 %>
<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../database/connection.asp" -->
<!-- #include file="../lib/config.asp" -->
<!-- #include file="../lib/functions.asp" -->
<!-- #include file="../lib/auth.asp" -->
<!-- #include file="../lib/upload.asp" -->
<%
AdminGirisGerekli

Dim binaKodu, dosyaAdi, baslik, hata, hedefYol, islemTipi
binaKodu = Trim(Request("bina"))
dosyaAdi = Trim(Request("dosya"))
hata = ""

If binaKodu = "" Or dosyaAdi = "" Then
    Response.Redirect "panel.asp"
    Response.End
End If

If Not DosyaAdiGecerliMi(binaKodu, dosyaAdi) Then
    Response.Write "Geçersiz dosya tanımı."
    Response.End
End If

baslik = BaslikGetir(binaKodu, dosyaAdi)
EnsureAyKlasoru binaKodu
hedefYol = NobetDosyaFizikselYolu(binaKodu, dosyaAdi)
islemTipi = "yukle"

If NobetDosyaMevcut(binaKodu, dosyaAdi) Then
    islemTipi = "guncelle"
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim yuklemeHatasi, basarili
    basarili = DosyaYukleVeKaydet(hedefYol, yuklemeHatasi)

    If basarili Then
        NobetDosyaDbKaydet binaKodu, dosyaAdi, baslik, Session(SESSION_ADMIN_KEY & "_ad"), True
        If islemTipi = "guncelle" Then
            Response.Redirect "panel.asp?mesaj=guncellendi"
        Else
            Response.Redirect "panel.asp?mesaj=yuklendi"
        End If
        Response.End
    Else
        hata = yuklemeHatasi
    End If
End If
%>
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="utf-8">
  <meta http-equiv="Content-Language" content="tr">
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
          <a href="panel.asp" class="btn btn-secondary btn-sm">Panele Dön</a>
        </div>
      </div>

      <% If hata <> "" Then %>
        <div class="alert alert-error"><%= Server.HTMLEncode(hata) %></div>
      <% End If %>

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
          <span class="info-label">Dönem</span>
          <span class="info-value"><%= GetAyBaslikMetni() %></span>
        </div>
        <div class="info-item full">
          <span class="info-label">Hedef Klasör</span>
          <span class="info-value mono">listeler/<%= GetYil() %>/<%= Server.HTMLEncode(binaKodu) %>/<%= GetAyKlasorAdi() %>/<%= Server.HTMLEncode(dosyaAdi) %></span>
        </div>
      </div>

      <form class="upload-form" method="post" enctype="multipart/form-data" action="yukle.asp?bina=<%= Server.URLEncode(binaKodu) %>&dosya=<%= Server.URLEncode(dosyaAdi) %>">
        <div class="form-group">
          <label for="dosya">PDF / Excel Dosyası</label>
          <input class="file-input" type="file" id="dosya" name="dosya" accept=".pdf,.xls,.xlsx" required>
        </div>
        <div class="form-actions">
          <button type="submit" class="btn btn-primary"><% If islemTipi = "guncelle" Then %>Güncelle<% Else %>Yükle<% End If %></button>
          <a href="panel.asp" class="btn btn-secondary">İptal</a>
        </div>
      </form>
    </div>
  </div>
</body>
</html>
