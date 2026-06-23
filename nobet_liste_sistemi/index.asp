<%@ Language=VBScript CodePage=65001 %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType = "text/html; charset=utf-8"
%>
<!-- #include file="ayarlar.asp" -->
<!-- #include file="database/connection.asp" -->
<!-- #include file="lib/config.asp" -->
<!-- #include file="lib/functions.asp" -->
<%
EnsureTumAyKlasorleri
%>
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta http-equiv="Content-Language" content="tr">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Nöbet Listeleri - MÜ Pendik E.A.H.</title>
  <link rel="icon" href="<%= PORTAL_IMAGES_YOLU %>hastane_portal_logo.png">
  <link rel="stylesheet" href="assets/style.css">
  <script>
    document.addEventListener('DOMContentLoaded', function() {
      var btn = document.getElementById('scrollTopBtn');
      if (!btn) return;
      window.addEventListener('scroll', function() {
        if (window.scrollY > 300) {
          btn.style.visibility = 'visible';
          btn.style.opacity = '1';
        } else {
          btn.style.opacity = '0';
          btn.style.visibility = 'hidden';
        }
      });
    });
    function scrollToTop() {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  </script>
</head>
<body class="portal-body" style="background-image: url('<%= PORTAL_IMAGES_YOLU & PORTAL_ARKAPLAN_RESIM %>');">
  <div id="scrollTopBtn" class="ghost-btn scroll-top-btn" onclick="scrollToTop()" title="Yukarı Çık">^</div>

  <div align="center" class="golgeliKutu portal-kutu">
    <table class="portal-table" width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td class="banner-cell">
          <img src="<%= PORTAL_IMAGES_YOLU & PORTAL_BANNER_RESIM %>" width="100%" height="165" alt="Marmara Üniversitesi Pendik Eğitim ve Araştırma Hastanesi">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF" valign="top" align="center" class="portal-icerik">

          <div class="baslik">
            <div class="left-side">
              <button type="button" class="back-button ghost-btn" onclick="window.location.href='<%= PORTAL_ANA_SAYFA %>'" title="Ana Sayfa">
                <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
                  <path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/>
                </svg>
              </button>
            </div>
            <div class="center-title">
              PENDİK EĞİTİM &amp; ARAŞTIRMA HASTANESİ<br/>
              NÖBET LİSTELERİ
            </div>
            <div class="right-side"></div>
          </div>

          <div class="donem-etiket"><%= GetAyBaslikMetni() %></div>

          <div class="liste-wrapper">
            <div class="liste-container">
              <div class="liste-bolum-baslik">Pendik E.A.H. Nöbet Listeleri</div>
              <table class="nobet-liste-tablo" border="0" cellpadding="0" cellspacing="0">
                <% RenderNobetListeTablosu BINA_PENDIK, pendikNobetListeleri, GuncelYil(), GuncelAyKlasor() %>
              </table>
            </div>
          </div>

          <div class="baslik bolum-baslik">
            <div class="left-side"></div>
            <div class="center-title center-title-sm">
              PROF. DR. ASAF ATASEVEN EK HİZMET BİNASI<br/>
              NÖBET LİSTELERİ
            </div>
            <div class="right-side"></div>
          </div>

          <div class="donem-etiket"><%= GetAyBaslikMetni() %></div>

          <div class="liste-wrapper">
            <div class="liste-container">
              <div class="liste-bolum-baslik">Asaf Ataseven Ek Hizmet Binası</div>
              <table class="nobet-liste-tablo" border="0" cellpadding="0" cellspacing="0">
                <% RenderNobetListeTablosu BINA_BASIBUYUK, basibuyukNobetListeleri, GuncelYil(), GuncelAyKlasor() %>
                <tr class="nobet-liste-row">
                  <td class="nobet-liste-cell nobet-liste-cell-phone">
                    <span class="liste-icon liste-icon-phone" aria-hidden="true"></span>
                    <span class="duz-metn">Pacs Destek (0531 682 44 36)</span>
                  </td>
                </tr>
              </table>
            </div>
          </div>

        </td>
      </tr>
    </table>
  </div>
</body>
</html>
