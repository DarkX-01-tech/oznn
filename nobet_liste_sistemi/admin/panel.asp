<%@ Language=VBScript CodePage=65001 %>
<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../database/connection.asp" -->
<!-- #include file="../lib/config.asp" -->
<!-- #include file="../lib/functions.asp" -->
<!-- #include file="../lib/auth.asp" -->
<!-- #include file="../lib/ui.asp" -->
<%
AdminGirisGerekli
%>
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="utf-8">
  <meta http-equiv="Content-Language" content="tr">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Yönetim Paneli</title>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body class="admin-body">
  <div class="page-shell">
    <div class="admin-wrap">
      <div class="admin-topbar">
        <div>
          <p class="eyebrow">Ana Yönetim</p>
          <h1>Nöbet Liste Yönetimi</h1>
        </div>
        <% AdminUstLinkler %>
      </div>

      <% FlashMesajGoster %>

      <div class="panel-cards">
        <div class="panel-card panel-card-primary">
          <div class="panel-card-icon">📤</div>
          <h2>Nöbet Listesi Yükleme / Güncelleme</h2>
          <p>Yıl ve ay seçerek ilgili dönemin listelerini yükleyin veya güncelleyin.</p>
          <a href="donem_sec.asp" class="btn btn-primary">Devam Et</a>
        </div>

        <div class="panel-card">
          <div class="panel-card-icon">📊</div>
          <h2>İstatistikler</h2>
          <p>Yüklenen listelerin yükleme ve güncelleme istatistiklerini görüntüleyin.</p>
          <a href="istatistikler.asp" class="btn btn-primary">Devam Et</a>
        </div>

        <div class="panel-card">
          <div class="panel-card-icon">📁</div>
          <h2>Geçmiş Nöbet Listeleri</h2>
          <p>Geçmiş dönemlere ait yüklenmiş nöbet listelerini inceleyin.</p>
          <a href="gecmis.asp" class="btn btn-primary">Devam Et</a>
        </div>
      </div>
    </div>
  </div>
  <!-- #include file="../includes/page_footer.asp" -->
</body>
</html>
