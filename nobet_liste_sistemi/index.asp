<!-- #include file="ayarlar.asp" -->
<!-- #include file="database/connection.asp" -->
<!-- #include file="lib/config.asp" -->
<!-- #include file="lib/functions.asp" -->
<%
EnsureTumAyKlasorleri
%>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Language" content="tr">
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
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
<body>
  <div id="scrollTopBtn" onclick="scrollToTop()">^</div>

  <div align="center" class="golgeliKutu">
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center" style="padding: 20px 0 0;">
          <div class="baslik">
            PENDİK Eğitim &amp; Araştırma Hastanesi<br/>
            NÖBET LİSTELERİ
          </div>
          <div class="ay-baslik"><%= GetAyBaslikMetni() %></div>
        </td>
      </tr>
      <tr>
        <td style="padding: 0 30px 30px;">
          <table border="0" width="100%" style="border-collapse: collapse">
            <% RenderNobetListeTablosu BINA_PENDIK, pendikNobetListeleri %>
          </table>

          <div class="baslik">
            PROF. DR. ASAF ATASEVEN EK HİZMET BİNASI<br/>
            NÖBET LİSTELERİ
          </div>
          <div class="ay-baslik"><%= GetAyBaslikMetni() %></div>

          <table border="0" width="100%" style="border-collapse: collapse">
            <% RenderNobetListeTablosu BINA_BASIBUYUK, basibuyukNobetListeleri %>
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
