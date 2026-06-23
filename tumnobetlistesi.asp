<!-- #include file="admin/database/Connection.asp" -->
<!-- #include file="ayarlar.asp" -->
<!-- #include file="lib/nobet_liste_lib.asp" -->
<!-- #include file="lib/nobet_liste_config.asp" -->
<%
EnsureAyKlasoru
%>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Language" content="tr">
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <title>MÜ Pendik E.A.H. Portal</title>
  <link rel="icon" href="images/hastane_portal_logo.png"/>

  <style type="text/css">
    @font-face {
      font-family: 'Open Sans';
      src: url('/webfonts/OpenSans-Regular.ttf') format('truetype');
      font-weight: 400;
      font-style: normal;
    }
    @font-face {
      font-family: 'Open Sans';
      src: url('/webfonts/OpenSans-SemiBold.ttf') format('truetype');
      font-weight: 600;
      font-style: normal;
    }
    @font-face {
      font-family: 'Open Sans';
      src: url('/webfonts/OpenSans-ExtraBold.ttf') format('truetype');
      font-weight: 800;
      font-style: normal;
    }

    body {
      background: url('images/arka-plan.png') no-repeat center center fixed;
      background-size: cover;
      font-family: 'Open Sans', sans-serif;
      margin: 0;
      padding: 0;
      overflow-x: hidden;
    }

    .golgeliKutu {
      box-shadow: 0 8px 16px rgba(0,0,0,0.2);
      width: 900px;
      margin: 40px auto;
      background: #fff;
      border-radius: 8px;
      overflow: hidden;
      transition: background-color 0.3s, box-shadow 0.3s;
    }
    .golgeliKutu:hover {
      background-color: #f9f9f9;
      box-shadow: 0 12px 24px rgba(0,0,0,0.3);
    }

    .baslik {
      font: 600 24px/1 'Open Sans';
      color: #850303;
      text-align: center;
      margin: 25px 0 8px;
      text-transform: uppercase;
      letter-spacing: 1px;
      padding-bottom: 12px;
      text-shadow: 1px 1px 4px rgba(0,0,0,0.3);
      transition: color 0.3s, border-color 0.3s;
    }
    .baslik:hover {
      color: #343a40;
      border-color: #c70039;
    }
    .baslik::after {
      content: "";
      display: block;
      width: 725px;
      height: 2px;
      background: #343a40;
      margin: 15px auto 0;
    }

    .ay-baslik {
      font-family: 'Open Sans', sans-serif;
      font-size: 18px;
      font-weight: 600;
      color: #25abb9;
      text-align: center;
      margin: 0 0 20px;
      letter-spacing: 0.5px;
    }

    .duyuru-baslik {
      font-family: 'Open Sans', sans-serif;
      font-size: 16px;
      color: #333;
      margin-bottom: 5px;
      font-weight: 600;
      text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
      border-bottom: 2px solid #850303;
      padding-bottom: 4px;
      display: inline-block;
      transition: color 0.3s, border-color 0.3s;
    }
    .duyuru-baslik:hover {
      color: #850303;
      border-color: #343a40;
    }

    .duyuru-icerik {
      font-family: 'Open Sans', sans-serif;
      font-size: 12px;
      color: #555;
      font-weight: 600;
      line-height: 1.6;
      margin-bottom: 10px;
      text-align: justify;
      text-shadow: 1px 1px 2px rgba(0,0,0,0.05);
      overflow-wrap: break-word;
      word-break: break-word;
      white-space: normal;
    }
    .duyuru-icerik b,
    .duyuru-icerik strong {
      font-weight: 800;
    }

    a {
      text-decoration: none;
      color: #45b8c3;
      transition: color 0.3s;
      font-weight: normal;
    }
    a:hover {
      color: #1e8c99;
    }

    #scrollTopBtn {
      position: fixed;
      bottom: 30px;
      right: 30px;
      width: 50px;
      height: 50px;
      background: #25abb9;
      color: #fff;
      font-size: 24px;
      text-align: center;
      line-height: 50px;
      border-radius: 50%;
      box-shadow: 0 4px 8px rgba(0,0,0,0.3);
      cursor: pointer;
      transition: all 0.3s ease, opacity 0.5s ease;
      opacity: 0;
      visibility: hidden;
      z-index: 999;
    }
    #scrollTopBtn:hover {
      background: #1e8c99;
      box-shadow: 0 6px 12px rgba(0,0,0,0.5);
      transform: scale(1.1);
    }

    .yazi-stil {
      font-family: 'Open Sans', sans-serif;
      font-size: 13px;
      color: #333;
      line-height: 1.6;
      margin-bottom: 15px;
      position: relative;
      padding-left: 40px;
      font-weight: 600;
      transition: color 0.3s;
      vertical-align: middle;
    }
    .yazi-stil::before {
      content: '\1F4CB';
      color: #d9534f;
      position: absolute;
      left: 15px;
      top: 0;
      font-size: 14px;
    }

    .duyuru-link {
      color:       #25abb9       !important;
      font-family: 'Open Sans', sans-serif !important;
      font-weight: 600            !important;
      font-size:   12px           !important;
      line-height: 1.6            !important;
      text-decoration: none       !important;
      vertical-align: middle;
    }
    .duyuru-link:hover {
      color: #1e8c99 !important;
    }

    .duyuru-link-pasif {
      color:       #b0b0b0       !important;
      font-family: 'Open Sans', sans-serif !important;
      font-weight: 600            !important;
      font-size:   12px           !important;
      line-height: 1.6            !important;
      text-decoration: none       !important;
      vertical-align: middle;
      cursor: not-allowed;
    }

    .no-icon::before {
      content: '\1F4DE'       !important;
      color:   #000           !important;
      position: absolute;
      left: 15px;
      top: 0;
      font-size: 14px;
      vertical-align: middle;
    }
    .no-icon .duz-metn {
      padding-left: 0;
      color:       #000      !important;
      font-family: 'Open Sans', sans-serif !important;
      font-weight: 600       !important;
      font-size:   12px      !important;
      line-height: 1.6       !important;
      display: inline;
      vertical-align: middle;
    }

    .duz-metn::before {
      content: none !important;
    }
  </style>

  <script>
    document.addEventListener('DOMContentLoaded', function() {
      setTimeout(function() {
        window.scrollTo({ top: 0, behavior: 'smooth' });
      }, 15000);
    });

    document.addEventListener('DOMContentLoaded', function() {
      var scrollTopBtn = document.getElementById('scrollTopBtn');
      window.addEventListener('scroll', function() {
        if (window.scrollY > 300) {
          scrollTopBtn.style.visibility = "visible";
          scrollTopBtn.style.opacity = "1";
        } else {
          scrollTopBtn.style.opacity = "0";
          scrollTopBtn.style.visibility = "hidden";
        }
      });
    });

    function scrollToTop() {
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  </script>
</head>

<body>
  <div id="scrollTopBtn" onclick="scrollToTop()">^</div>

  <div align="center" class="golgeliKutu">
    <table id="Table_01" width="900" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="3">
          <img src="images/muhst_06.png" width="900" height="165" alt="">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF" width="166" height="604" valign="top">
          <!--#include file="sol.asp"-->
        </td>

        <td bgcolor="#FFFFFF" width="734" height="604" valign="top" align="center">
          <div class="baslik">
            PENDİK Eğitim &amp; Araştırma Hastanesi <br/>
            NÖBET LİSTELERİ
          </div>
          <div class="ay-baslik"><%= GetAyBaslikMetni() %></div>

          <table border="0" width="95%" id="table2" style="border-collapse: collapse">
          </table>

          <table border="0" width="100%" style="border-collapse: collapse">
            <% RenderNobetListeTablosu pendikNobetListeleri %>
          </table>

          <div class="baslik">
            PROF. DR. ASAF ATASEVEN EK HİZMET BİNASI <br/>
            NÖBET LİSTELERİ
          </div>
          <div class="ay-baslik"><%= GetAyBaslikMetni() %></div>

          <table border="0" width="100%" style="border-collapse: collapse">
            <% RenderNobetListeTablosu asafAtasevenNobetListeleri %>
            <tr>
              <td class="yazi-stil no-icon">
                <span class="duz-metn">
                  Pacs Destek (0531 682 44 36)
                </span>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </div>
</body>
</html>
