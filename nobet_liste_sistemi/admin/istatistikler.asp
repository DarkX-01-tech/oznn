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

Dim seciliYil, seciliAy, yilParam, ayParam, i, yillar, aylar
yilParam = Trim(Request("yil"))
ayParam = LCase(Trim(Request("ay")))

If yilParam <> "" And IsNumeric(yilParam) Then
    seciliYil = CInt(yilParam)
Else
    seciliYil = 0
End If

If ayParam <> "" And AyKlasorGecerliMi(ayParam) Then
    seciliAy = ayParam
Else
    seciliAy = ""
End If
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
        <% AdminUstLinkler %>
      </div>

      <% If yilParam = "" Then %>
        <div class="dashboard-card">
          <p><a href="panel.asp">← Yönetim Paneline Dön</a></p>
          <h2>Yıl Seçin</h2>
          <p>Yalnızca listesi yüklenmiş yıllar görüntülenir.</p>
          <div class="link-grid">
            <% yillar = IstatistikYillariDizisi()
            For i = 0 To UBound(yillar) %>
              <a class="link-card" href="istatistikler.asp?yil=<%= yillar(i) %>"><%= yillar(i) %></a>
            <% Next %>
          </div>
        </div>
      <% ElseIf ayParam = "" Then %>
        <div class="dashboard-card">
          <p><a href="istatistikler.asp">← Yıllara Dön</a></p>
          <h2><%= seciliYil %> — Ay Seçin</h2>
          <p>Yalnızca dosya yüklenmiş aylar listelenir.</p>
          <div class="link-grid">
            <% aylar = YukluAylarForYil(seciliYil)
            If UBound(aylar) < 0 Then %>
              <p class="empty-note">Bu yıl için yüklenmiş liste bulunamadı.</p>
            <% Else
              For i = 0 To UBound(aylar) %>
                <a class="link-card" href="istatistikler.asp?yil=<%= seciliYil %>&ay=<%= Server.URLEncode(aylar(i)) %>"><%= AyBaslikFromKlasor(aylar(i)) %></a>
              <% Next
            End If %>
          </div>
        </div>
      <% Else %>
        <div class="dashboard-card">
          <p><a href="istatistikler.asp?yil=<%= seciliYil %>">← <%= seciliYil %> Aylarına Dön</a></p>
          <h2><%= AyBaslikFromKlasor(seciliAy) %> <%= seciliYil %> İstatistikleri</h2>
        </div>

        <div class="content-panel istatistik-panel">
          <% RenderIstatistikTamListe seciliYil, seciliAy %>
        </div>
      <% End If %>
    </div>
  </div>
</body>
</html>
