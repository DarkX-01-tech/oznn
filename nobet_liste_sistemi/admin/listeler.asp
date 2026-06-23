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
<!-- #include file="../lib/ui.asp" -->
<%
AdminGirisGerekli
EnsureSeciliDonemKlasorleri

Dim seciliYil, seciliAy

seciliYil = GetSeciliYil()
seciliAy = GetSeciliAyKlasor()

Sub RenderAdminTablo(binaKodu, listeDizisi, yil, ayKlasor)
    Dim i, satir, baslik, dosyaAdi, aktif, webYolu, qs
    qs = "yil=" & yil & "&ay=" & Server.URLEncode(ayKlasor)

    For i = 0 To UBound(listeDizisi)
        satir = listeDizisi(i)
        baslik = satir(0)
        dosyaAdi = satir(1)
        aktif = NobetDosyaAktif(binaKodu, dosyaAdi, yil, ayKlasor)
        NobetDosyaDbSenkronize binaKodu, dosyaAdi, baslik, yil, ayKlasor

        Response.Write "<tr>"
        Response.Write "<td>" & Server.HTMLEncode(baslik) & "</td>"
        Response.Write "<td><code>" & Server.HTMLEncode(dosyaAdi) & "</code></td>"

        If aktif Then
            webYolu = NobetDosyaWebYolu(binaKodu, dosyaAdi, yil, ayKlasor)
            Response.Write "<td><span class=""badge badge-aktif"">Aktif</span></td>"
            Response.Write "<td><div class=""action-group"">"
            Response.Write "<a class=""btn btn-secondary btn-sm"" href=""" & webYolu & """ target=""_blank"">Görüntüle</a>"
            Response.Write "<a class=""btn btn-primary btn-sm"" href=""yukle.asp?bina=" & Server.URLEncode(binaKodu) & "&dosya=" & Server.URLEncode(dosyaAdi) & "&" & qs & """>Güncelle</a>"
            Response.Write "<a class=""btn btn-danger btn-sm"" href=""dosya_sil.asp?bina=" & Server.URLEncode(binaKodu) & "&dosya=" & Server.URLEncode(dosyaAdi) & "&" & qs & """ onclick=""return confirm('Dosya silinsin mi?');"">Sil</a>"
            Response.Write "</div></td>"
        Else
            Response.Write "<td><span class=""badge badge-pasif"">Pasif</span></td>"
            Response.Write "<td><a class=""btn btn-primary btn-sm"" href=""yukle.asp?bina=" & Server.URLEncode(binaKodu) & "&dosya=" & Server.URLEncode(dosyaAdi) & "&" & qs & """>Yükle</a></td>"
        End If

        Response.Write "</tr>"
    Next
End Sub
%>
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta http-equiv="Content-Language" content="tr">
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Liste Yönetimi</title>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body class="admin-body">
  <div class="page-shell">
    <div class="admin-wrap">
      <div class="admin-topbar">
        <div>
          <p class="eyebrow">Liste Yönetimi</p>
          <h1>Nöbet Listeleri</h1>
        </div>
        <% AdminUstLinkler %>
      </div>

      <p class="breadcrumb"><a href="donem_sec.asp">← Dönem Seçimine Dön</a></p>

      <% FlashMesajGoster %>
      <% DonemOzetKartlariGoster seciliYil, seciliAy %>

      <div class="admin-section">
        <table class="admin-table">
          <tr><th colspan="4" class="bina-baslik"><%= BinaAdiGoster(BINA_PENDIK) %></th></tr>
          <tr><th>Liste Adı</th><th>Dosya Adı</th><th>Durum</th><th>İşlem</th></tr>
          <% RenderAdminTablo BINA_PENDIK, pendikNobetListeleri, seciliYil, seciliAy %>
        </table>
      </div>

      <div class="admin-section">
        <table class="admin-table">
          <tr><th colspan="4" class="bina-baslik"><%= BinaAdiGoster(BINA_BASIBUYUK) %></th></tr>
          <tr><th>Liste Adı</th><th>Dosya Adı</th><th>Durum</th><th>İşlem</th></tr>
          <% RenderAdminTablo BINA_BASIBUYUK, basibuyukNobetListeleri, seciliYil, seciliAy %>
        </table>
      </div>
    </div>
  </div>
</body>
</html>
