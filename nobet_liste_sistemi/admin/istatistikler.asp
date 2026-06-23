<%@ Language=VBScript CodePage=65001 %>
<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../database/connection.asp" -->
<!-- #include file="../lib/config.asp" -->
<!-- #include file="../lib/functions.asp" -->
<!-- #include file="../lib/stats.asp" -->
<!-- #include file="../lib/auth.asp" -->
<!-- #include file="../lib/ui.asp" -->
<%
AdminGirisGerekli

Dim seciliYil, seciliAy, yilParam, ayParam, i, satir, baslik, dosyaAdi, bilgi
yilParam = Trim(Request("yil"))
ayParam = LCase(Trim(Request("ay")))

If yilParam <> "" And IsNumeric(yilParam) Then
    seciliYil = CInt(yilParam)
Else
    seciliYil = GetSeciliYil()
End If

If ayParam <> "" And AyKlasorGecerliMi(ayParam) Then
    seciliAy = ayParam
    DonemKaydet seciliYil, seciliAy
Else
    seciliAy = ""
End If

Sub RenderIstatistikSatirlari(binaKodu, listeDizisi, yil, ayKlasor)
  For i = 0 To UBound(listeDizisi)
    satir = listeDizisi(i)
    baslik = satir(0)
    dosyaAdi = satir(1)
    bilgi = DosyaKayitBilgisi(yil, binaKodu, ayKlasor, dosyaAdi)

    Response.Write "<tr>"
    Response.Write "<td>" & Server.HTMLEncode(BinaAdiGoster(binaKodu)) & "</td>"
    Response.Write "<td>" & Server.HTMLEncode(baslik) & "</td>"
    Response.Write "<td><code>" & Server.HTMLEncode(dosyaAdi) & "</code></td>"
    If NobetDosyaAktif(binaKodu, dosyaAdi, yil, ayKlasor) Then
      Response.Write "<td><span class=""badge badge-aktif"">Aktif</span></td>"
    Else
      Response.Write "<td><span class=""badge badge-pasif"">Pasif</span></td>"
    End If
    Response.Write "<td>" & TarihGoster(bilgi(0)) & "</td>"
    Response.Write "<td>" & bilgi(1) & "</td>"
    Response.Write "<td>" & TarihGoster(bilgi(2)) & "</td>"
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
  <title>İstatistikler</title>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body class="admin-body">
  <div class="page-shell">
    <div class="admin-wrap">
      <div class="admin-topbar">
        <div>
          <p class="eyebrow">Raporlama</p>
          <h1>İstatistikler</h1>
        </div>
        <% AdminNavGoster "istatistikler" %>
      </div>

      <% If yilParam = "" Then %>
        <div class="dashboard-card">
          <h2>Yıl Seçin</h2>
          <div class="link-grid">
            <% Dim yillar: yillar = IstatistikYillariDizisi()
            For i = 0 To UBound(yillar) %>
              <a class="link-card" href="istatistikler.asp?yil=<%= yillar(i) %>"><%= yillar(i) %></a>
            <% Next %>
          </div>
        </div>
      <% ElseIf ayParam = "" Then %>
        <div class="dashboard-card">
          <h2><%= seciliYil %> - Ay Seçin</h2>
          <p><a href="istatistikler.asp">← Yıllara Dön</a></p>
          <div class="link-grid">
            <% For i = 1 To 12 %>
              <a class="link-card" href="istatistikler.asp?yil=<%= seciliYil %>&ay=<%= Server.URLEncode(AyKlasorFromNumara(i)) %>"><%= TurkceAyAdi(i) %></a>
            <% Next %>
          </div>
        </div>
      <% Else %>
        <div class="dashboard-card">
          <p><a href="istatistikler.asp?yil=<%= seciliYil %>">← <%= seciliYil %> Aylarına Dön</a></p>
          <h2><%= AyBaslikFromKlasor(seciliAy) %> <%= seciliYil %> İstatistikleri</h2>
        </div>

        <table class="admin-table">
          <tr>
            <th>Bina</th>
            <th>Liste Adı</th>
            <th>Dosya Adı</th>
            <th>Durum</th>
            <th>İlk Yükleme</th>
            <th>Güncelleme Sayısı</th>
            <th>Son Güncelleme</th>
          </tr>
          <% RenderIstatistikSatirlari BINA_PENDIK, pendikNobetListeleri, seciliYil, seciliAy %>
          <% RenderIstatistikSatirlari BINA_BASIBUYUK, basibuyukNobetListeleri, seciliYil, seciliAy %>
        </table>
      <% End If %>
    </div>
  </div>
</body>
</html>
