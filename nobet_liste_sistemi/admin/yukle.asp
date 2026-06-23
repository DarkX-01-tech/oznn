<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../database/connection.asp" -->
<!-- #include file="../lib/config.asp" -->
<!-- #include file="../lib/functions.asp" -->
<!-- #include file="../lib/auth.asp" -->
<!-- #include file="../lib/upload.asp" -->
<%
AdminGirisGerekli

Dim binaKodu, dosyaAdi, baslik, mesaj, hata, hedefYol
binaKodu = Trim(Request("bina"))
dosyaAdi = Trim(Request("dosya"))
mesaj = ""
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

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim yuklemeHatasi, basarili
    basarili = DosyaYukleVeKaydet(hedefYol, yuklemeHatasi)

    If basarili Then
        NobetDosyaDbKaydet binaKodu, dosyaAdi, baslik, Session(SESSION_ADMIN_KEY & "_ad"), True
        mesaj = "Dosya başarıyla yüklendi. Liste linki aktif hale geldi."
    Else
        hata = yuklemeHatasi
    End If
End If
%>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <title>Dosya Yükle</title>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body>
  <div class="admin-wrap">
    <div class="admin-header">
      <h1>Dosya Yükle</h1>
      <div class="admin-nav"><a href="panel.asp">Panele Dön</a></div>
    </div>

    <% If mesaj <> "" Then %>
      <div class="alert alert-success"><%= Server.HTMLEncode(mesaj) %></div>
    <% End If %>
    <% If hata <> "" Then %>
      <div class="alert alert-error"><%= Server.HTMLEncode(hata) %></div>
    <% End If %>

    <p><strong>Bina:</strong> <%= Server.HTMLEncode(BinaAdiGoster(binaKodu)) %></p>
    <p><strong>Liste:</strong> <%= Server.HTMLEncode(baslik) %></p>
    <p><strong>Dönem:</strong> <%= GetAyBaslikMetni() %></p>
    <p><strong>Hedef:</strong> listeler/<%= GetYil() %>/<%= Server.HTMLEncode(binaKodu) %>/<%= GetAyKlasorAdi() %>/<%= Server.HTMLEncode(dosyaAdi) %></p>

    <form method="post" enctype="multipart/form-data" action="yukle.asp?bina=<%= Server.URLEncode(binaKodu) %>&dosya=<%= Server.URLEncode(dosyaAdi) %>">
      <div class="form-group">
        <label for="dosya">PDF / Excel Dosyası</label>
        <input type="file" id="dosya" name="dosya" accept=".pdf,.xls,.xlsx" required>
      </div>
      <button type="submit" class="btn btn-primary">Yükle</button>
      <a href="panel.asp" class="btn btn-secondary">İptal</a>
    </form>
  </div>
</body>
</html>
