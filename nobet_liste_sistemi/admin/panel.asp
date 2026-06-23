<%@ Language=VBScript CodePage=65001 %>
<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../database/connection.asp" -->
<!-- #include file="../lib/config.asp" -->
<!-- #include file="../lib/functions.asp" -->
<!-- #include file="../lib/auth.asp" -->
<!-- #include file="../lib/ui.asp" -->
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
            Response.Write "<td><span class=""badge badge-aktif"">Aktif</span></td>"
            Response.Write "<td><div class=""action-group"">"
            Response.Write "<a class=""btn btn-secondary btn-sm"" href=""" & webYolu & """ target=""_blank"">Görüntüle</a>"
            Response.Write "<a class=""btn btn-primary btn-sm"" href=""yukle.asp?bina=" & Server.URLEncode(binaKodu) & "&dosya=" & Server.URLEncode(dosyaAdi) & """>Güncelle</a>"
            Response.Write "<a class=""btn btn-danger btn-sm"" href=""dosya_sil.asp?bina=" & Server.URLEncode(binaKodu) & "&dosya=" & Server.URLEncode(dosyaAdi) & """ onclick=""return confirm('Dosya silinsin mi?');"">Sil</a>"
            Response.Write "</div></td>"
        Else
            Response.Write "<td><span class=""badge badge-pasif"">Pasif</span></td>"
            Response.Write "<td><a class=""btn btn-primary btn-sm"" href=""yukle.asp?bina=" & Server.URLEncode(binaKodu) & "&dosya=" & Server.URLEncode(dosyaAdi) & """>Yükle</a></td>"
        End If

        Response.Write "</tr>"
    Next
End Sub
%>
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="utf-8">
  <meta http-equiv="Content-Language" content="tr">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Nöbet Liste Yönetimi</title>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body class="admin-body">
  <div class="page-shell">
    <div class="admin-wrap">
      <div class="admin-topbar">
        <div>
          <p class="eyebrow">Yönetim Paneli</p>
          <h1>Nöbet Liste Yönetimi</h1>
        </div>
        <div class="admin-nav">
          <span class="user-chip"><%= Server.HTMLEncode(Session(SESSION_ADMIN_KEY & "_ad")) %></span>
          <a href="../index.asp" target="_blank" class="btn btn-secondary btn-sm">Liste Sayfası</a>
          <a href="logout.asp" class="btn btn-danger btn-sm">Çıkış</a>
        </div>
      </div>

      <% FlashMesajGoster %>

      <div class="summary-grid">
        <div class="summary-card">
          <span class="label">Dönem</span>
          <span class="value"><%= GetAyBaslikMetni() %></span>
        </div>
        <div class="summary-card">
          <span class="label">Yıl Klasörü</span>
          <span class="value"><%= GetYil() %></span>
        </div>
        <div class="summary-card">
          <span class="label">Ay Klasörü</span>
          <span class="value path">listeler/<%= GetYil() %>/<%= GetAyKlasorAdi() %>/</span>
        </div>
      </div>

      <div class="admin-section">
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
      </div>

      <div class="admin-section">
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
    </div>
  </div>
</body>
</html>
