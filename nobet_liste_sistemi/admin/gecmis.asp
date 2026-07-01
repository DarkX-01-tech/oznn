<%@ Language=VBScript CodePage=65001 %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType = "text/html; charset=utf-8"
%>
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
  <meta http-equiv="Content-Language" content="tr">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Geçmiş Nöbet Listeleri</title>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body class="admin-body">
  <div class="page-shell">
    <div class="admin-wrap">
      <div class="admin-topbar">
        <div>
          <p class="eyebrow">Arşiv</p>
          <h1>Geçmiş Nöbet Listeleri</h1>
        </div>
        <% AdminUstLinkler %>
      </div>

      <% If yilParam = "" Then %>
        <div class="dashboard-card">
          <p><a href="panel.asp">← Yönetim Paneline Dön</a></p>
          <h2>Yıl Seçin</h2>
          <p>Geçmiş döneme ait yüklenmiş listeler. Bulunduğunuz ay burada görünmez.</p>
          <div class="link-grid">
            <% yillar = GecmisYillariDizisi()
            If Not DiziDoluMu(yillar) Then %>
              <p class="empty-note">Henüz geçmiş dönem listesi bulunmuyor.</p>
            <% Else
              For i = 0 To UBound(yillar) %>
                <a class="link-card" href="gecmis.asp?yil=<%= yillar(i) %>"><%= yillar(i) %></a>
              <% Next
            End If %>
          </div>
        </div>
      <% ElseIf ayParam = "" Then %>
        <div class="dashboard-card">
          <p><a href="gecmis.asp">← Yıllara Dön</a></p>
          <h2><%= seciliYil %> — Ay Seçin</h2>
          <p>Yalnızca geçmiş dönem ve dosya yüklenmiş aylar listelenir.</p>
          <div class="link-grid">
            <% aylar = GecmisAylarForYil(seciliYil)
            If Not DiziDoluMu(aylar) Then %>
              <p class="empty-note">Bu yıl için geçmiş dönem listesi bulunamadı.</p>
            <% Else
              For i = 0 To UBound(aylar) %>
                <a class="link-card" href="gecmis.asp?yil=<%= seciliYil %>&ay=<%= Server.URLEncode(aylar(i)) %>"><%= AyBaslikFromKlasor(aylar(i)) %></a>
              <% Next
            End If %>
          </div>
        </div>
      <% Else %>
        <div class="dashboard-card">
          <p><a href="gecmis.asp?yil=<%= seciliYil %>">← <%= seciliYil %> Aylarına Dön</a></p>
          <h2><%= AyBaslikFromKlasor(seciliAy) %> <%= seciliYil %></h2>
        </div>

        <div class="content-panel gecmis-panel">
          <% RenderTamListeGorunumu seciliYil, seciliAy %>
        </div>
      <% End If %>
    </div>
  </div>
</body>
</html>
