<%@ Language=VBScript CodePage=65001 %>
<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../database/connection.asp" -->
<!-- #include file="../lib/config.asp" -->
<!-- #include file="../lib/functions.asp" -->
<!-- #include file="../lib/auth.asp" -->
<%
AdminGirisGerekli
EnsureTumAyKlasorleri

Sub RenderAdminTablo(binaKodu, listeDizisi)
    Dim i, satir, baslik, dosyaAdi, aktif, webYolu
    For i = 0 To UBound(listeDizisi)
        satir = listeDizisi(i)
        baslik = satir(0)
        dosyaAdi = satir(1)
        aktif = NobetDosyaAktif(binaKodu, dosyaAdi)
        NobetDosyaDbSenkronize binaKodu, dosyaAdi, baslik

        Response.Write "<tr>"
        Response.Write "<td>" & Server.HTMLEncode(baslik) & "</td>"
        Response.Write "<td><code>" & Server.HTMLEncode(dosyaAdi) & "</code></td>"

        If aktif Then
            webYolu = NobetDosyaWebYolu(binaKodu, dosyaAdi)
            Response.Write "<td class=""durum-aktif"">Aktif</td>"
            Response.Write "<td>"
            Response.Write "<a class=""btn btn-secondary"" href=""" & webYolu & """ target=""_blank"">Görüntüle</a> "
            Response.Write "<a class=""btn btn-primary"" href=""yukle.asp?bina=" & Server.URLEncode(binaKodu) & "&dosya=" & Server.URLEncode(dosyaAdi) & """>Güncelle</a> "
            Response.Write "<a class=""btn btn-danger"" href=""dosya_sil.asp?bina=" & Server.URLEncode(binaKodu) & "&dosya=" & Server.URLEncode(dosyaAdi) & """ onclick=""return confirm('Dosya silinsin mi?');"">Sil</a>"
            Response.Write "</td>"
        Else
            Response.Write "<td class=""durum-pasif"">Pasif</td>"
            Response.Write "<td><a class=""btn btn-primary"" href=""yukle.asp?bina=" & Server.URLEncode(binaKodu) & "&dosya=" & Server.URLEncode(dosyaAdi) & """>Yükle</a></td>"
        End If

        Response.Write "</tr>"
    Next
End Sub
%>
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Nöbet Liste Yönetimi</title>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body>
  <div class="admin-wrap">
    <div class="admin-header">
      <h1>Nöbet Liste Yönetimi</h1>
      <div class="admin-nav">
        <span><%= Server.HTMLEncode(Session(SESSION_ADMIN_KEY & "_ad")) %></span>
        <a href="../index.asp" target="_blank">Liste Sayfası</a>
        <a href="logout.asp">Çıkış</a>
      </div>
    </div>

    <p><strong>Dönem:</strong> <%= GetAyBaslikMetni() %></p>
    <p><strong>Klasör yapısı:</strong> listeler/<%= GetYil() %>/pendik|basibuyuk/<%= GetAyKlasorAdi() %>/</p>

    <table class="admin-table">
      <tr><th colspan="4" class="bina-baslik"><%= BinaAdiGoster(BINA_PENDIK) %></th></tr>
      <tr>
        <th>Liste Adı</th>
        <th>Dosya Adı</th>
        <th>Durum</th>
        <th>İşlem</th>
      </tr>
      <% RenderAdminTablo BINA_PENDIK, pendikNobetListeleri %>
    </table>

    <br>

    <table class="admin-table">
      <tr><th colspan="4" class="bina-baslik"><%= BinaAdiGoster(BINA_BASIBUYUK) %></th></tr>
      <tr>
        <th>Liste Adı</th>
        <th>Dosya Adı</th>
        <th>Durum</th>
        <th>İşlem</th>
      </tr>
      <% RenderAdminTablo BINA_BASIBUYUK, basibuyukNobetListeleri %>
    </table>
  </div>
</body>
</html>
