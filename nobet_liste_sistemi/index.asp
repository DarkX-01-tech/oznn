<%@ Language=VBScript CodePage=65001 %>
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
  <meta charset="utf-8">
  <meta http-equiv="Content-Language" content="tr">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Nöbet Listeleri - MÜ Pendik E.A.H.</title>
  <link rel="stylesheet" href="assets/style.css">
  <script>
    document.addEventListener('DOMContentLoaded', function() {
      var btn = document.getElementById('scrollTopBtn');
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
<body class="public-body">
  <div id="scrollTopBtn" onclick="scrollToTop()">^</div>

  <div class="page-shell">
    <div class="golgeliKutu">
      <div class="hero-band">
        <h1>Pendik Eğitim &amp; Araştırma Hastanesi</h1>
        <p>Nöbet Listeleri &middot; <%= GetAyBaslikMetni() %></p>
      </div>

      <div class="content-panel">
        <div class="section-title">Pendik E.A.H. Nöbet Listeleri</div>
        <div class="ay-baslik"><%= GetAyBaslikMetni() %></div>
        <table class="liste-tablo">
            <% RenderNobetListeTablosu BINA_PENDIK, pendikNobetListeleri, GuncelYil(), GuncelAyKlasor() %>
        </table>

        <div class="section-title">Prof. Dr. Asaf Ataseven Ek Hizmet Binası</div>
        <div class="ay-baslik"><%= GetAyBaslikMetni() %></div>
        <table class="liste-tablo">
            <% RenderNobetListeTablosu BINA_BASIBUYUK, basibuyukNobetListeleri, GuncelYil(), GuncelAyKlasor() %>
          <tr>
            <td class="yazi-stil no-icon">
              <span class="duz-metn">Pacs Destek (0531 682 44 36)</span>
            </td>
          </tr>
        </table>
      </div>
    </div>
  </div>
</body>
</html>
