<%
'====================================
' Cache Kontrolü
'====================================
Response.CacheControl = "no-cache, no-store, must-revalidate"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="database/Connection.asp" -->
<%
'====================================
' Parametreler
'====================================
Dim secilen_yil, secilen_ay, menu_tipi, hedef_tablo, tip_param

If Request.QueryString("yil") <> "" And Request.QueryString("ay") <> "" Then
    secilen_yil = CInt(Request.QueryString("yil"))
    secilen_ay = CInt(Request.QueryString("ay"))
Else
    Response.Redirect "panel.asp"
End If

If Request.QueryString("tip") = "diyet" Then
    menu_tipi = "diyet"
    hedef_tablo = "diyet_yemek_listesi"
    tip_param = "&tip=diyet"
Else
    menu_tipi = "normal"
    hedef_tablo = "yemek_listesi"
    tip_param = ""
End If

'====================================
' Veritabanından Verileri Çek
'====================================
Dim rsYemek, sqlYemek

sqlYemek = "SELECT * FROM " & hedef_tablo & " " & _
           "WHERE YEAR(tarih) = " & secilen_yil & " " & _
           "AND MONTH(tarih) = " & secilen_ay & " " & _
           "ORDER BY tarih ASC"
Set rsYemek = ConnYemek.Execute(sqlYemek)

Function GetMonthName(monthNum)
    Select Case monthNum
        Case 1: GetMonthName = "Ocak"
        Case 2: GetMonthName = "&#350;ubat"
        Case 3: GetMonthName = "Mart"
        Case 4: GetMonthName = "Nisan"
        Case 5: GetMonthName = "May&#305;s"
        Case 6: GetMonthName = "Haziran"
        Case 7: GetMonthName = "Temmuz"
        Case 8: GetMonthName = "A&#287;ustos"
        Case 9: GetMonthName = "Eyl&#252;l"
        Case 10: GetMonthName = "Ekim"
        Case 11: GetMonthName = "Kas&#305;m"
        Case 12: GetMonthName = "Aral&#305;k"
    End Select
End Function

Dim ay_adi
ay_adi = GetMonthName(secilen_ay)
%>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Language" content="tr">
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
  <meta http-equiv="Pragma" content="no-cache">
  <meta http-equiv="Expires" content="0">
  <title><%= ay_adi %> | <%= secilen_yil %> - <% If menu_tipi = "diyet" Then %>Diyet <% End If %>Yemek Listesi</title>
  <link rel="icon" href="../../images/hastane_portal_logo.png"/>

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

    :root {
      --primary-color: <% If menu_tipi = "diyet" Then %>#4caf50<% Else %>#25abb9<% End If %>;
      --primary-dark: <% If menu_tipi = "diyet" Then %>#388e3c<% Else %>#1e8c99<% End If %>;
      --secondary-color: #dc3545;
      --white: #fff;
    }

    * { margin: 0; padding: 0; box-sizing: border-box; }

    body {
      background: #f5f5f5;
      font-family: 'Open Sans', sans-serif;
      margin: 0; padding: 0;
      overflow-x: hidden;
    }

    .container { max-width: 100%; margin: 0 auto; background: #fff; }

    .header-banner { width: 100%; height: auto; display: block; }

    .baslik { text-align: center; margin: 15px 0 5px 0; }

    .baslik .center-title {
      font-size: 28px; font-weight: 600;
      text-transform: uppercase; letter-spacing: 1px;
      background: linear-gradient(45deg, var(--secondary-color), var(--primary-color));
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
      margin: 0; padding: 0;
    }

    .type-badge {
      display: inline-block;
      background: <% If menu_tipi = "diyet" Then %>rgba(76,175,80,0.1)<% Else %>rgba(37,171,185,0.1)<% End If %>;
      color: var(--primary-dark);
      padding: 4px 14px; border-radius: 12px;
      font-size: 12px; font-weight: 600;
      margin-top: 8px;
    }

    .takvim-wrapper { padding: 8px 20px 20px 20px; }

    .tablo-container-wrapper {
      border: 2px solid #ddd; border-radius: 8px;
      overflow: hidden; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
    }

    .tablo-header {
      background-color: var(--primary-color);
      display: table; width: 100%; table-layout: fixed;
    }

    .tablo-header-row { display: table-row; }

    .tablo-header-cell {
      display: table-cell; padding: 14px 10px;
      text-align: center; font-weight: 700; font-size: 13px;
      color: #ffffff; border: 1px solid var(--primary-dark);
      text-transform: uppercase; letter-spacing: 0.5px;
      font-family: 'Open Sans', sans-serif;
    }

    .tablo-header-cell.tarih { width: 15%; }
    .tablo-header-cell.ogle { width: 42.5%; }
    .tablo-header-cell.aksam { width: 42.5%; }

    .tablo-body-container {
      max-height: 490px; overflow-y: auto; overflow-x: hidden;
    }

    .tablo-body-container::-webkit-scrollbar { width: 8px; }
    .tablo-body-container::-webkit-scrollbar-track { background: #f1f1f1; }
    .tablo-body-container::-webkit-scrollbar-thumb { background: var(--primary-color); border-radius: 5px; }
    .tablo-body-container::-webkit-scrollbar-thumb:hover { background: var(--primary-dark); }

    .yemek-tablo {
      width: 100%; border-collapse: collapse;
      background-color: #fff; font-family: 'Open Sans', sans-serif;
      table-layout: fixed;
    }

    .yemek-tablo td {
      padding: 12px 12px; border: 1px solid #ddd;
      font-size: 12px; color: #333; vertical-align: middle;
    }

    .yemek-tablo .tarih-col {
      width: 15%; text-align: center; font-weight: 600;
      background-color: #f5f5f5; vertical-align: middle;
    }

    .tarih-date {
      display: block; font-size: 16px; font-weight: 700;
      color: var(--primary-color); margin-bottom: 6px;
    }

    .tarih-day {
      display: inline-block; font-size: 11px; font-weight: 600;
      color: #555; background: #fff; padding: 4px 10px;
      border-radius: 5px; text-transform: uppercase;
      letter-spacing: 0.5px; border: 1px solid #ddd;
    }

    .yemek-tablo .ogle-col,
    .yemek-tablo .aksam-col { width: 42.5%; }

    .yemek-tablo td ul {
      margin: 0; padding-left: 18px;
      list-style-type: disc; list-style-position: outside;
    }

    .yemek-tablo td ul li { margin: 4px 0; line-height: 1.5; color: #333; }
    .yemek-tablo td ul li::marker { color: var(--primary-color); }

    .yemek-tablo tbody tr:nth-child(even) { background-color: #fafafa; }
    .yemek-tablo tbody tr:hover { background-color: #f0f9fa; transition: background-color 0.3s; }

    .bos-mesaj {
      text-align: center; padding: 60px 20px;
      color: #666; font-size: 15px; margin: 20px;
    }
    .bos-mesaj strong { display: block; font-size: 18px; color: #333; margin-bottom: 10px; }

    #scrollTopBtn {
      position: fixed; bottom: 30px; right: 30px;
      width: 50px; height: 50px;
      background-color: var(--primary-color);
      color: #fff; font-size: 24px; text-align: center;
      line-height: 50px; border-radius: 50%;
      box-shadow: 0 4px 8px rgba(0, 0, 0, 0.3);
      cursor: pointer; transition: all 0.3s ease, opacity 0.5s ease;
      opacity: 0; visibility: hidden; z-index: 999;
      font-family: 'Open Sans', sans-serif;
    }
    #scrollTopBtn:hover {
      background-color: var(--primary-dark);
      box-shadow: 0 6px 12px rgba(0, 0, 0, 0.5);
      transform: scale(1.1);
    }
  </style>

  <script>
    document.addEventListener("DOMContentLoaded", function() {
      var btn = document.getElementById("scrollTopBtn");
      var container = document.querySelector('.tablo-body-container');
      if (container) {
        container.addEventListener("scroll", function() {
          if (container.scrollTop > 300) {
            btn.style.visibility = 'visible';
            btn.style.opacity = '1';
          } else {
            btn.style.opacity = '0';
            btn.style.visibility = 'hidden';
          }
        });
      }
    });

    function scrollToTop() {
      var container = document.querySelector('.tablo-body-container');
      if (container) {
        container.scrollTo({ top: 0, behavior: 'smooth' });
      }
    }
  </script>
</head>
<body>
  <div id="scrollTopBtn" onclick="scrollToTop()">^</div>

  <div class="container">
    <table id="Table_01" width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="3">
          <img src="../../images/muhst_06.png" width="100%" height="165" alt="">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF" width="100%" valign="top" align="center">

          <div class="baslik">
            <div class="center-title">
              AYLIK <% If menu_tipi = "diyet" Then %>D&#304;YET <% End If %>YEMEK MEN&#220;S&#220; (<%= ay_adi %>)
            </div>
            <% If menu_tipi = "diyet" Then %>
            <span class="type-badge">Diyet Men&#252;s&#252;</span>
            <% End If %>
          </div>

          <div class="takvim-wrapper">
            <% If Not rsYemek.EOF Then %>
            <div class="tablo-container-wrapper">

              <div class="tablo-header">
                <div class="tablo-header-row">
                  <div class="tablo-header-cell tarih">TAR&#304;H</div>
                  <div class="tablo-header-cell ogle">&#214;&#286;LE YEME&#286;&#304;</div>
                  <div class="tablo-header-cell aksam">AK&#350;AM YEME&#286;&#304;</div>
                </div>
              </div>

              <div class="tablo-body-container">
                <table class="yemek-tablo">
                  <tbody>
                    <% Do While Not rsYemek.EOF %>
                    <tr>
                      <td class="tarih-col">
                        <span class="tarih-date"><%= Right("0" & Day(rsYemek("tarih")), 2) %>.<%= Right("0" & Month(rsYemek("tarih")), 2) %></span>
                        <span class="tarih-day"><%= rsYemek("gun_adi") %></span>
                      </td>
                      <td class="ogle-col">
                        <ul>
                          <% If rsYemek("ogle_corba") <> "" Then %>
                          <li><%= rsYemek("ogle_corba") %></li>
                          <% End If %>
                          <% If rsYemek("ogle_ana_yemek") <> "" Then %>
                          <li><%= rsYemek("ogle_ana_yemek") %></li>
                          <% End If %>
                          <% If rsYemek("ogle_yan_urun") <> "" Then %>
                          <li><%= rsYemek("ogle_yan_urun") %></li>
                          <% End If %>
                          <% If rsYemek("ogle_tatli") <> "" Then %>
                          <li><%= rsYemek("ogle_tatli") %></li>
                          <% End If %>
                        </ul>
                      </td>
                      <td class="aksam-col">
                        <ul>
                          <% If rsYemek("aksam_corba") <> "" Then %>
                          <li><%= rsYemek("aksam_corba") %></li>
                          <% End If %>
                          <% If rsYemek("aksam_ana_yemek") <> "" Then %>
                          <li><%= rsYemek("aksam_ana_yemek") %></li>
                          <% End If %>
                          <% If rsYemek("aksam_yan_urun") <> "" Then %>
                          <li><%= rsYemek("aksam_yan_urun") %></li>
                          <% End If %>
                          <% If rsYemek("aksam_tatli") <> "" Then %>
                          <li><%= rsYemek("aksam_tatli") %></li>
                          <% End If %>
                        </ul>
                      </td>
                    </tr>
                    <%
                    rsYemek.MoveNext
                    Loop
                    %>
                  </tbody>
                </table>
              </div>
            </div>
            <% Else %>
            <div class="bos-mesaj">
              <p>Bu Ay &#304;&#231;in Hen&#252;z <% If menu_tipi = "diyet" Then %>Diyet <% End If %>Yemek Men&#252;s&#252; Eklenmemi&#351;.</p>
            </div>
            <% End If %>
          </div>

        </td>
      </tr>
    </table>
  </div>
</body>
</html>
<%
rsYemek.Close
Set rsYemek = Nothing
%>
