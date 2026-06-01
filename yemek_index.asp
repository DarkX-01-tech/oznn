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
' Izleme Modu Kontrolu
'====================================
Dim izlemeModuAktif, izlemeParam, izlemeJS
If Request.QueryString("izleme") = "1" Then
    izlemeModuAktif = True
    izlemeParam = "&izleme=1"
    izlemeJS = "true"
Else
    izlemeModuAktif = False
    izlemeParam = ""
    izlemeJS = "false"
End If

'====================================
' Ay Seçimi
'====================================
Dim secilen_yil, secilen_ay
Dim baslangic_yil, baslangic_ay
baslangic_yil = 2026
baslangic_ay = 3

Dim bugun
bugun = Now()
Dim bugun_yil, bugun_ay, bugun_gun
bugun_yil = Year(bugun)
bugun_ay = Month(bugun)
bugun_gun = Day(bugun)

If Request.QueryString("yil") <> "" And Request.QueryString("ay") <> "" Then
    secilen_yil = CInt(Request.QueryString("yil"))
    secilen_ay = CInt(Request.QueryString("ay"))
    If secilen_yil < baslangic_yil Or (secilen_yil = baslangic_yil And secilen_ay < baslangic_ay) Then
        secilen_yil = baslangic_yil
        secilen_ay = baslangic_ay
    End If
Else
    If bugun >= DateSerial(baslangic_yil, baslangic_ay, 1) Then
        secilen_yil = bugun_yil
        secilen_ay = bugun_ay
    Else
        secilen_yil = baslangic_yil
        secilen_ay = baslangic_ay
    End If
End If

'====================================
' Veritabanından Verileri Çek
'====================================
Dim rsYemek, sqlYemek

sqlYemek = "SELECT * FROM yemek_listesi " & _
           "WHERE aktif = True " & _
           "AND YEAR(tarih) = " & secilen_yil & " " & _
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
  <title>M&#220; Pendik E.A.H. Portal - Yemek Listesi</title>
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
      --primary-color: #25abb9;
      --primary-dark: #1e8c99;
      --secondary-color: #dc3545;
      --secondary-dark: #c82333;
      --white: #fff;
      --shadow-medium: rgba(37,171,185,0.3);
      --shadow-heavy: rgba(30,140,153,0.5);
      --transition-speed: 0.25s;
    }

    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
      -webkit-user-select: none;
      -moz-user-select: none;
      -ms-user-select: none;
      user-select: none;
    }

    body {
      background: url('../../images/arka-plan.png') no-repeat center center fixed;
      background-size: cover;
      font-family: 'Open Sans', sans-serif;
      margin: 0;
      padding: 0;
      overflow-x: hidden;
    }

    .golgeliKutu {
      box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2);
      width: 900px;
      margin: 40px auto;
      background-color: #fff;
      border-radius: 8px;
      overflow: hidden;
      transition: box-shadow 0.3s, background-color 0.3s;
    }

    .baslik {
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 15px 0 5px 0;
    }

    .left-side,
    .right-side {
      width: 100px;
      display: flex;
      justify-content: center;
      align-items: center;
    }

    .center-title {
      flex: 1;
      text-align: center;
      font-size: 28px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 1px;
      background: linear-gradient(45deg, var(--secondary-color), var(--primary-color));
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
    }

    .takvim-wrapper {
      padding: 8px 20px 20px 20px;
      box-sizing: border-box;
      display: flex;
      flex-direction: column;
      align-items: center;
    }

    .takvim-header {
      width: 100%;
      margin-bottom: 15px;
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 10px;
      justify-content: flex-start;
    }

    .nav-left {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .back-button {
      background-color: var(--secondary-color);
      border: 2px solid var(--white);
      cursor: pointer;
      box-shadow: 0 4px 8px rgba(220,53,69,0.3);
      width: 40px;
      height: 40px;
      border-radius: 8px;
      display: flex;
      align-items: center;
      justify-content: center;
      transition: all var(--transition-speed) ease;
      position: relative;
      overflow: hidden;
      padding: 0;
    }

    .back-button:hover {
      transform: scale(1.05);
      box-shadow: 0 6px 12px rgba(200,35,51,0.5);
    }

    .back-button svg {
      width: 20px;
      height: 20px;
      fill: var(--white);
    }

    #scrollTopBtn {
      position: fixed;
      bottom: 30px;
      right: 30px;
      width: 50px;
      height: 50px;
      background-color: #25abb9;
      color: #fff;
      font-size: 24px;
      text-align: center;
      line-height: 50px;
      border-radius: 50%;
      box-shadow: 0 4px 8px rgba(0, 0, 0, 0.3);
      cursor: pointer;
      transition: all 0.3s ease, opacity 0.5s ease;
      opacity: 0;
      visibility: hidden;
      z-index: 999;
      font-family: 'Open Sans', sans-serif;
    }
    #scrollTopBtn:hover {
      background-color: #1e8c99;
      box-shadow: 0 6px 12px rgba(0, 0, 0, 0.5);
      transform: scale(1.1);
    }

    .tablo-container-wrapper {
      width: 98%;
      margin: 0 auto;
      border: 2px solid #ddd;
      border-radius: 8px;
      overflow: hidden;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
    }

    .tablo-header {
      background-color: #25abb9;
      display: table;
      width: 100%;
      table-layout: fixed;
    }

    .tablo-header-row {
      display: table-row;
    }

    .tablo-header-cell {
      display: table-cell;
      padding: 14px 10px;
      text-align: center;
      font-weight: 700;
      font-size: 13px;
      color: #ffffff;
      border: 1px solid #1e8c99;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      font-family: 'Open Sans', sans-serif;
    }

    .tablo-header-cell.tarih {
      width: 15%;
    }

    .tablo-header-cell.ogle {
      width: 42.5%;
    }

    .tablo-header-cell.aksam {
      width: 42.5%;
    }

    .tablo-body-container {
      max-height: 490px;
      overflow-y: auto;
      overflow-x: hidden;
    }

    .tablo-body-container::-webkit-scrollbar {
      width: 8px;
    }
    .tablo-body-container::-webkit-scrollbar-track {
      background: #f1f1f1;
    }
    .tablo-body-container::-webkit-scrollbar-thumb {
      background: #25abb9;
      border-radius: 5px;
    }
    .tablo-body-container::-webkit-scrollbar-thumb:hover {
      background: #1e8c99;
    }

    .yemek-tablo {
      width: 100%;
      border-collapse: collapse;
      background-color: #fff;
      font-family: 'Open Sans', sans-serif;
      table-layout: fixed;
    }

    .yemek-tablo td {
      padding: 12px 12px;
      border: 1px solid #ddd;
      font-size: 12px;
      color: #333;
      vertical-align: middle;
    }

    .yemek-tablo .tarih-col {
      width: 15%;
      text-align: center;
      font-weight: 600;
      background-color: #f5f5f5;
      vertical-align: middle;
    }

    .tarih-date {
      display: block;
      font-size: 15px;
      font-weight: 700;
      color: #25abb9;
      margin-bottom: 6px;
    }

    .tarih-day {
      display: inline-block;
      font-size: 11px;
      font-weight: 600;
      color: #555;
      background: #fff;
      padding: 4px 10px;
      border-radius: 5px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      border: 1px solid #ddd;
    }

    .yemek-tablo .ogle-col,
    .yemek-tablo .aksam-col {
      width: 42.5%;
    }

    .yemek-tablo td ul {
      margin: 0;
      padding-left: 18px;
      list-style-type: disc;
      list-style-position: outside;
    }

    .yemek-tablo td ul li {
      margin: 4px 0;
      line-height: 1.5;
      color: #333;
    }

    .yemek-tablo td ul li::marker {
      color: #25abb9;
    }

    .yemek-tablo tbody tr:nth-child(even) {
      background-color: #fafafa;
    }

    .yemek-tablo tbody tr:hover {
      background-color: #f0f9fa;
      transition: background-color 0.3s;
    }

    .yemek-tablo tbody tr.bugun-satir {
      background-color: #e0f7fa !important;
      border-left: 4px solid #25abb9;
    }
    .yemek-tablo tbody tr.bugun-satir .tarih-col {
      background-color: #d4f1f5 !important;
    }
    .yemek-tablo tbody tr.bugun-satir .tarih-date {
      color: #1e8c99;
      font-size: 16px;
    }
    .yemek-tablo tbody tr.bugun-satir .tarih-day {
      background: #25abb9;
      color: #fff;
      border-color: #1e8c99;
    }

    .bos-mesaj {
      text-align: center;
      padding: 60px 20px;
      color: #666;
      font-size: 15px;
      margin: 20px;
    }

    .bos-mesaj strong {
      display: block;
      font-size: 18px;
      color: #333;
      margin-bottom: 10px;
    }
  </style>

  <script>
    document.addEventListener("DOMContentLoaded", function() {
      var todayRow = document.querySelector('tr[data-today="1"]');
      var container = document.querySelector('.tablo-body-container');
      if (todayRow && container) {
        setTimeout(function() {
          var rowTop = todayRow.offsetTop - container.offsetTop;
          container.scrollTo({ top: rowTop - 50, behavior: 'smooth' });
        }, 300);
      }

      setTimeout(function() {
        var c = document.querySelector('.tablo-body-container');
        if (c) {
          c.scrollTo({ top: 0, behavior: 'smooth' });
        }
      }, 15000);

      var btn = document.getElementById("scrollTopBtn");
      var cont = document.querySelector('.tablo-body-container');
      if (cont) {
        cont.addEventListener("scroll", function() {
          if (cont.scrollTop > 300) {
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

  <div align="center" class="golgeliKutu">
    <table id="Table_01" width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="3">
          <img src="../../images/muhst_06.png" width="100%" height="165" alt="">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF" width="100%" valign="top" align="center">

          <div class="baslik">
            <div class="left-side"></div>
            <div class="center-title">
              AYLIK YEMEK MEN&#220;S&#220; (<%= ay_adi %>)
            </div>
            <div class="right-side"></div>
          </div>

          <div class="takvim-wrapper">

            <% If Not izlemeModuAktif Then %>
            <div class="takvim-header">
              <div class="nav-left">
                <button class="back-button" onclick="window.history.back()" title="Geri">
                  <svg width="24" height="24" viewBox="0 0 24 24">
                    <path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/>
                  </svg>
                </button>
              </div>
            </div>
            <% End If %>

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
                    <%
                    Dim satirSayac
                    satirSayac = 0
                    Do While Not rsYemek.EOF
                      satirSayac = satirSayac + 1
                      Dim kayitGun, kayitAy, kayitYil, bugunMu
                      kayitGun = Day(rsYemek("tarih"))
                      kayitAy = Month(rsYemek("tarih"))
                      kayitYil = Year(rsYemek("tarih"))
                      bugunMu = False
                      If kayitGun = bugun_gun And kayitAy = bugun_ay And kayitYil = bugun_yil Then
                        bugunMu = True
                      End If
                    %>
                    <tr id="gun_<%= kayitGun %>"<% If bugunMu Then %> data-today="1" class="bugun-satir"<% End If %>>
                      <td class="tarih-col">
                        <span class="tarih-date"><%= Right("0" & kayitGun, 2) %>.<%= Right("0" & kayitAy, 2) %>.<%= kayitYil %></span>
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
              <p>Bu Ay &#304;&#231;in Hen&#252;z Yemek Men&#252;s&#252; Eklenmemi&#351;.</p>
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
