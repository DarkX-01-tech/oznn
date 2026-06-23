<%@ Language=VBScript CodePage=1254 %>
<%
Response.CodePage = 1254
Response.CharSet = "windows-1254"
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
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
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
      window.scrollTo(0, 0);
    }
  </script>
</head>
<body class="portal-body" style="background-image: url('<%= PORTAL_IMAGES_YOLU & PORTAL_ARKAPLAN_RESIM %>');">
  <div id="scrollTopBtn" onclick="scrollToTop()">^</div>

  <div align="center" class="golgeliKutu portal-kutu">
    <table class="portal-table" width="900" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="3">
          <img src="<%= PORTAL_IMAGES_YOLU & PORTAL_BANNER_RESIM %>" width="900" height="165" alt="">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF" width="900" valign="top" align="center" class="portal-icerik">
          <div class="baslik-row">
            <a href="<%= PORTAL_ANA_SAYFA %>" class="home-icon-btn" title="Ana Sayfaya Dön" aria-label="Ana Sayfaya Dön">
              <svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true">
                <path fill="currentColor" d="M12 3l9 8h-3v9h-5v-6H11v6H6v-9H3l9-8z"/>
              </svg>
            </a>
            <div class="baslik">
              PENDİK EĞİTİM &amp; ARAŞTIRMA HASTANESİ<br/>
              NÖBET LİSTELERİ
            </div>
          </div>

          <div class="ay-baslik"><%= GetAyBaslikMetni() %></div>
          <table border="0" width="100%" class="liste-tablo" style="border-collapse: collapse">
            <% RenderNobetListeTablosu BINA_PENDIK, pendikNobetListeleri, GuncelYil(), GuncelAyKlasor() %>
          </table>

          <div class="baslik">
            PROF. DR. ASAF ATASEVEN EK HİZMET BİNASI<br/>
            NÖBET LİSTELERİ
          </div>

          <div class="ay-baslik"><%= GetAyBaslikMetni() %></div>
          <table border="0" width="100%" class="liste-tablo" style="border-collapse: collapse">
            <% RenderNobetListeTablosu BINA_BASIBUYUK, basibuyukNobetListeleri, GuncelYil(), GuncelAyKlasor() %>
            <tr>
              <td class="yazi-stil no-icon">
                <span class="duz-metn">Pacs Destek (0531 682 44 36)</span>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </div>
</body>
</html>
