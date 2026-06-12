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
      --primary-light: #5ec8d3;
      --primary-dark: #1e8c99;
      --primary-fade: rgba(37, 171, 185, 0.07);
      --secondary-color: #dc3545;
      --secondary-dark: #c82333;
      --white: #fff;
      --surface: #ffffff;
      --surface-alt: #f7fafb;
      --text-main: #2c3e45;
      --text-soft: #6b7c84;
      --line-color: #e4eaec;
      --line-accent: rgba(37, 171, 185, 0.22);
      --btn-border: #e0e0e0;
      --shadow-card: 0 10px 40px rgba(30, 60, 70, 0.1), 0 2px 8px rgba(30, 60, 70, 0.04);
      --shadow-medium: rgba(37, 171, 185, 0.3);
      --shadow-heavy: rgba(30, 140, 153, 0.5);
      --transition-speed: 0.25s;
      --ease-out: cubic-bezier(0.22, 1, 0.36, 1);
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
      color: var(--text-main);
      -webkit-font-smoothing: antialiased;
      -moz-osx-font-smoothing: grayscale;
    }

    body::before {
      content: '';
      position: fixed;
      inset: 0;
      background: rgba(255, 255, 255, 0.45);
      pointer-events: none;
      z-index: 0;
    }

    .golgeliKutu {
      position: relative;
      z-index: 1;
      width: 900px;
      margin: 40px auto;
      background-color: var(--surface);
      border-radius: 8px;
      overflow: hidden;
      box-shadow: var(--shadow-card);
      transition: box-shadow 0.3s ease, background-color 0.3s ease;
      animation: softReveal 0.55s var(--ease-out) both;
    }

    .golgeliKutu::before {
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      height: 3px;
      background: linear-gradient(90deg, var(--secondary-color), var(--primary-color), var(--primary-dark));
      z-index: 2;
    }

    @keyframes softReveal {
      from { opacity: 0; transform: translateY(10px); }
      to { opacity: 1; transform: translateY(0); }
    }

    .banner-wrap {
      position: relative;
      overflow: hidden;
      line-height: 0;
    }

    .banner-wrap img {
      display: block;
      width: 100%;
      height: 165px;
      object-fit: cover;
    }

    .banner-wrap::after {
      content: '';
      position: absolute;
      left: 0;
      right: 0;
      bottom: 0;
      height: 32px;
      background: linear-gradient(to bottom, transparent, var(--surface));
      pointer-events: none;
    }

    .content-area {
      background: var(--surface);
    }

    .baslik {
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 15px 0 5px 0;
      position: relative;
    }

    .baslik::after {
      content: '';
      position: absolute;
      left: 50%;
      bottom: -6px;
      transform: translateX(-50%);
      width: 64px;
      height: 2px;
      border-radius: 2px;
      background: linear-gradient(90deg, var(--secondary-color), var(--primary-color));
      opacity: 0.55;
    }

    .left-side {
      width: 100px;
      display: flex;
      justify-content: flex-start;
      align-items: center;
      padding-left: 10px;
    }

    .right-side {
      width: 100px;
      display: flex;
      justify-content: flex-end;
      align-items: center;
      padding-right: 10px;
    }

    .center-title {
      flex: 1;
      text-align: center;
      font-size: 28px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 1px;
      line-height: 1.25;
    }

    .title-gradient {
      background: linear-gradient(45deg, var(--secondary-color), var(--primary-color));
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
    }

    .title-accent {
      font-weight: 600;
      color: var(--primary-dark);
      letter-spacing: 0.5px;
    }

    .ghost-btn {
      background: transparent;
      border: 2px solid var(--btn-border);
      color: #888;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      transition: all var(--transition-speed) var(--ease-out);
      font-family: 'Open Sans', sans-serif;
      text-decoration: none;
      box-shadow: none;
    }

    .ghost-btn:hover {
      transform: translateY(-2px);
      background: transparent;
    }

    .ghost-btn:active {
      transform: translateY(0);
    }

    .back-button {
      width: 36px;
      height: 36px;
      border-radius: 50%;
      padding: 0;
    }

    .back-button:hover {
      border-color: var(--secondary-color);
      color: var(--secondary-color);
      box-shadow: 0 4px 12px rgba(220, 53, 69, 0.15);
    }

    .back-button svg {
      width: 18px;
      height: 18px;
      fill: currentColor;
      transition: fill var(--transition-speed) ease;
    }

    .takvim-wrapper {
      padding: 8px 20px 20px 20px;
      box-sizing: border-box;
      display: flex;
      flex-direction: column;
      align-items: center;
    }

    #scrollTopBtn {
      position: fixed;
      bottom: 30px;
      right: 30px;
      width: 50px;
      height: 50px;
      background: transparent;
      border: 2px solid var(--btn-border);
      color: #888;
      display: flex;
      align-items: center;
      justify-content: center;
      border-radius: 50%;
      box-shadow: none;
      cursor: pointer;
      transition: all 0.3s ease, opacity 0.5s ease, visibility 0.5s ease;
      opacity: 0;
      visibility: hidden;
      z-index: 999;
      font-family: 'Open Sans', sans-serif;
    }

    #scrollTopBtn svg {
      width: 22px;
      height: 22px;
      fill: currentColor;
      transition: transform 0.25s var(--ease-out);
    }

    #scrollTopBtn:hover {
      background: transparent;
      border-color: var(--primary-color);
      color: var(--primary-color);
      box-shadow: 0 4px 12px var(--shadow-medium);
      transform: translateY(-2px);
    }

    #scrollTopBtn:hover svg {
      transform: translateY(-2px);
    }

    .tablo-container-wrapper {
      width: 98%;
      margin: 0 auto;
      border: 2px solid var(--line-color);
      border-radius: 8px;
      overflow: hidden;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
      background: var(--surface);
    }

    .tablo-header {
      background: linear-gradient(180deg, #2bb8c6 0%, var(--primary-color) 45%, var(--primary-dark) 100%);
      display: table;
      width: 100%;
      table-layout: fixed;
      position: relative;
    }

    .tablo-header::after {
      content: '';
      position: absolute;
      inset: 0;
      background: linear-gradient(180deg, rgba(255,255,255,0.14) 0%, rgba(255,255,255,0) 42%);
      pointer-events: none;
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
      border: 1px solid rgba(255, 255, 255, 0.12);
      text-transform: uppercase;
      letter-spacing: 0.5px;
      font-family: 'Open Sans', sans-serif;
      position: relative;
      z-index: 1;
      text-shadow: 0 1px 1px rgba(0, 0, 0, 0.12);
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
      scroll-behavior: smooth;
      background: var(--surface);
    }

    .tablo-body-container::-webkit-scrollbar {
      width: 8px;
    }
    .tablo-body-container::-webkit-scrollbar-track {
      background: #f1f1f1;
    }
    .tablo-body-container::-webkit-scrollbar-thumb {
      background: linear-gradient(180deg, var(--primary-light), var(--primary-color));
      border-radius: 5px;
    }
    .tablo-body-container::-webkit-scrollbar-thumb:hover {
      background: var(--primary-dark);
    }

    .yemek-tablo {
      width: 100%;
      border-collapse: collapse;
      background-color: var(--surface);
      font-family: 'Open Sans', sans-serif;
      table-layout: fixed;
    }

    .yemek-tablo td {
      padding: 12px 12px;
      border: 1px solid var(--line-color);
      font-size: 12px;
      color: #333;
      vertical-align: middle;
      transition: background-color 0.3s ease;
    }

    .yemek-tablo .tarih-col {
      width: 15%;
      text-align: center;
      font-weight: 600;
      background: linear-gradient(180deg, #f8fafa 0%, #f1f5f6 100%);
      vertical-align: middle;
    }

    .tarih-date {
      display: block;
      font-size: 15px;
      font-weight: 700;
      color: var(--primary-color);
      margin-bottom: 6px;
    }

    .tarih-day {
      display: inline-block;
      font-size: 11px;
      font-weight: 600;
      color: #555;
      background: var(--white);
      padding: 4px 10px;
      border-radius: 5px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      border: 1px solid var(--line-color);
    }

    .yemek-tablo .ogle-col,
    .yemek-tablo .aksam-col {
      width: 42.5%;
    }

    .yemek-tablo td ul {
      margin: 0;
      padding: 0;
      list-style: none;
    }

    .yemek-tablo td ul li {
      position: relative;
      margin: 4px 0;
      padding-left: 14px;
      line-height: 1.5;
      color: #333;
    }

    .yemek-tablo td ul li::before {
      content: '';
      position: absolute;
      left: 0;
      top: 0.58em;
      width: 4px;
      height: 4px;
      border-radius: 1px;
      transform: rotate(45deg);
      background: var(--primary-color);
      opacity: 0.75;
    }

    .yemek-tablo tbody tr:nth-child(even) {
      background-color: #fafafa;
    }

    .yemek-tablo tbody tr:hover {
      background-color: #f0f9fa;
      transition: background-color 0.3s ease;
    }

    .yemek-tablo tbody tr.bugun-satir {
      background-color: #e0f7fa !important;
      border-left: 4px solid var(--primary-color);
    }

    .yemek-tablo tbody tr.bugun-satir .tarih-col {
      background-color: #d4f1f5 !important;
    }

    .yemek-tablo tbody tr.bugun-satir .tarih-date {
      color: var(--primary-dark);
      font-size: 16px;
    }

    .yemek-tablo tbody tr.bugun-satir .tarih-day {
      background: var(--primary-color);
      color: var(--white);
      border-color: var(--primary-dark);
    }

    .bos-mesaj {
      text-align: center;
      padding: 60px 20px;
      color: var(--text-soft);
      font-size: 15px;
      margin: 20px;
      border: 1px solid var(--line-color);
      border-radius: 8px;
      background: linear-gradient(180deg, var(--surface-alt), var(--surface));
    }

    .bos-mesaj-icon {
      display: flex;
      align-items: center;
      justify-content: center;
      width: 48px;
      height: 48px;
      margin: 0 auto 14px;
      border-radius: 50%;
      border: 1px solid var(--line-accent);
      background: var(--primary-fade);
    }

    .bos-mesaj-icon svg {
      width: 22px;
      height: 22px;
      fill: var(--primary-color);
      opacity: 0.85;
    }

    .bos-mesaj p {
      margin: 0;
      line-height: 1.55;
      color: #666;
    }

    .bos-mesaj strong {
      display: block;
      font-size: 18px;
      color: #333;
      margin-bottom: 10px;
      font-weight: 700;
    }
  </style>

  <script>
    document.addEventListener("DOMContentLoaded", function() {
      var todayRow = document.querySelector('tr[data-today="1"]');
      var container = document.querySelector('.tablo-body-container');
      if (todayRow && container) {
        setTimeout(function() {
          var scrollPos = todayRow.offsetTop - Math.max(0, (container.clientHeight - todayRow.offsetHeight) / 2);
          if (scrollPos < 0) scrollPos = 0;
          container.scrollTo({ top: scrollPos, behavior: 'smooth' });
        }, 300);
      }

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
  <div id="scrollTopBtn" onclick="scrollToTop()" title="Yukar&#305; Kayd&#305;r">
    <svg viewBox="0 0 24 24"><path d="M7.41 15.41L12 10.83l4.59 4.58L18 14l-6-6-6 6z"/></svg>
  </div>

  <div align="center" class="golgeliKutu">
    <table id="Table_01" width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="3">
          <div class="banner-wrap">
            <img src="../../images/muhst_06.png" width="100%" height="165" alt="">
          </div>
        </td>
      </tr>
      <tr>
        <td class="content-area" width="100%" valign="top" align="center">

          <div class="baslik">
            <div class="left-side">
              <% If Not izlemeModuAktif Then %>
              <button class="back-button ghost-btn" onclick="window.location.href='http://10.201.65.10'" title="Ana Sayfa">
                <svg width="24" height="24" viewBox="0 0 24 24">
                  <path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/>
                </svg>
              </button>
              <% End If %>
            </div>
            <div class="center-title">
              <span class="title-gradient">AYLIK YEMEK MEN&#220;S&#220;</span> <span class="title-accent">(<%= ay_adi %>)</span>
            </div>
            <div class="right-side"></div>
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
              <div class="bos-mesaj-icon">
                <svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>
              </div>
              <p><strong>Men&#252; Bulunamad&#305;</strong>Bu ay i&#231;in hen&#252;z yemek men&#252;s&#252; eklenmemi&#351;.</p>
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
