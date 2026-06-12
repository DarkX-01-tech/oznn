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
      --primary-light: #4ec4d0;
      --primary-dark: #1e8c99;
      --primary-soft: rgba(37, 171, 185, 0.08);
      --primary-glow: rgba(37, 171, 185, 0.18);
      --secondary-color: #dc3545;
      --secondary-dark: #c82333;
      --white: #fff;
      --surface: #ffffff;
      --surface-muted: #f8fbfc;
      --text-primary: #1a2b33;
      --text-secondary: #5a6b73;
      --text-muted: #8a9aa3;
      --border-light: rgba(37, 171, 185, 0.12);
      --border-subtle: #e8eef0;
      --btn-border: #dde5e8;
      --shadow-soft: 0 4px 24px rgba(26, 43, 51, 0.06);
      --shadow-medium: rgba(37, 171, 185, 0.28);
      --shadow-heavy: rgba(30, 140, 153, 0.45);
      --radius-sm: 8px;
      --radius-md: 14px;
      --radius-lg: 20px;
      --transition-speed: 0.28s;
      --transition-smooth: cubic-bezier(0.4, 0, 0.2, 1);
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
      color: var(--text-primary);
      -webkit-font-smoothing: antialiased;
      -moz-osx-font-smoothing: grayscale;
    }

    body::before {
      content: '';
      position: fixed;
      inset: 0;
      background: linear-gradient(160deg, rgba(255,255,255,0.72) 0%, rgba(240,248,250,0.55) 50%, rgba(230,245,248,0.65) 100%);
      pointer-events: none;
      z-index: 0;
    }

    .golgeliKutu {
      position: relative;
      z-index: 1;
      width: 900px;
      margin: 36px auto 48px;
      background: var(--surface);
      border-radius: var(--radius-lg);
      overflow: hidden;
      border: 1px solid rgba(255, 255, 255, 0.85);
      box-shadow:
        0 1px 2px rgba(26, 43, 51, 0.04),
        0 8px 24px rgba(26, 43, 51, 0.08),
        0 24px 48px rgba(37, 171, 185, 0.1);
      transition: box-shadow 0.4s var(--transition-smooth), transform 0.4s var(--transition-smooth);
      animation: pageFadeIn 0.6s var(--transition-smooth) both;
    }

    @keyframes pageFadeIn {
      from { opacity: 0; transform: translateY(16px); }
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
      height: 48px;
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
      margin: 0;
      padding: 22px 20px 18px;
      position: relative;
    }

    .baslik::after {
      content: '';
      position: absolute;
      left: 24px;
      right: 24px;
      bottom: 0;
      height: 1px;
      background: linear-gradient(90deg, transparent, var(--border-light) 20%, var(--border-light) 80%, transparent);
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
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 8px;
    }

    .title-main {
      font-size: 26px;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 1.8px;
      line-height: 1.2;
      background: linear-gradient(135deg, var(--secondary-color) 0%, var(--primary-color) 55%, var(--primary-dark) 100%);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
    }

    .title-month {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 5px 16px;
      font-size: 12px;
      font-weight: 600;
      letter-spacing: 0.8px;
      text-transform: uppercase;
      color: var(--primary-dark);
      background: linear-gradient(135deg, var(--primary-soft), rgba(255,255,255,0.9));
      border: 1px solid var(--border-light);
      border-radius: 999px;
      box-shadow: 0 2px 8px var(--primary-glow);
    }

    .title-month::before {
      content: '';
      width: 6px;
      height: 6px;
      border-radius: 50%;
      background: var(--primary-color);
      box-shadow: 0 0 6px var(--primary-color);
    }

    .ghost-btn {
      background: rgba(255, 255, 255, 0.7);
      backdrop-filter: blur(8px);
      -webkit-backdrop-filter: blur(8px);
      border: 1.5px solid var(--btn-border);
      color: var(--text-muted);
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      transition: all var(--transition-speed) var(--transition-smooth);
      font-family: 'Open Sans', sans-serif;
      text-decoration: none;
      box-shadow: var(--shadow-soft);
    }

    .ghost-btn:hover {
      transform: translateY(-2px);
      background: rgba(255, 255, 255, 0.95);
      box-shadow: 0 6px 20px rgba(26, 43, 51, 0.1);
    }

    .ghost-btn:active {
      transform: translateY(0);
    }

    .back-button {
      width: 38px;
      height: 38px;
      border-radius: 50%;
      padding: 0;
    }

    .back-button:hover {
      border-color: var(--secondary-color);
      color: var(--secondary-color);
      box-shadow: 0 6px 18px rgba(220, 53, 69, 0.18);
    }

    .back-button svg {
      width: 18px;
      height: 18px;
      fill: currentColor;
      transition: fill var(--transition-speed) ease;
    }

    .takvim-wrapper {
      padding: 20px 24px 28px;
      box-sizing: border-box;
      display: flex;
      flex-direction: column;
      align-items: center;
    }

    #scrollTopBtn {
      position: fixed;
      bottom: 32px;
      right: 32px;
      width: 48px;
      height: 48px;
      background: rgba(255, 255, 255, 0.88);
      backdrop-filter: blur(10px);
      -webkit-backdrop-filter: blur(10px);
      border: 1.5px solid var(--border-light);
      color: var(--primary-color);
      display: flex;
      align-items: center;
      justify-content: center;
      border-radius: 50%;
      box-shadow: 0 4px 20px rgba(37, 171, 185, 0.15);
      cursor: pointer;
      transition: all 0.35s var(--transition-smooth), opacity 0.5s ease, visibility 0.5s ease;
      opacity: 0;
      visibility: hidden;
      z-index: 999;
      font-family: 'Open Sans', sans-serif;
    }

    #scrollTopBtn svg {
      width: 20px;
      height: 20px;
      fill: currentColor;
      transition: transform 0.3s var(--transition-smooth);
    }

    #scrollTopBtn:hover {
      background: var(--primary-color);
      border-color: var(--primary-color);
      color: var(--white);
      box-shadow: 0 8px 28px var(--shadow-medium);
      transform: translateY(-3px);
    }

    #scrollTopBtn:hover svg {
      transform: translateY(-1px);
    }

    .tablo-container-wrapper {
      width: 100%;
      margin: 0 auto;
      border: 1px solid var(--border-light);
      border-radius: var(--radius-md);
      overflow: hidden;
      box-shadow: var(--shadow-soft);
      background: var(--surface);
    }

    .tablo-header {
      background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-dark) 100%);
      display: table;
      width: 100%;
      table-layout: fixed;
      position: relative;
    }

    .tablo-header::after {
      content: '';
      position: absolute;
      inset: 0;
      background: linear-gradient(180deg, rgba(255,255,255,0.12) 0%, transparent 50%);
      pointer-events: none;
    }

    .tablo-header-row {
      display: table-row;
    }

    .tablo-header-cell {
      display: table-cell;
      padding: 16px 12px;
      text-align: center;
      font-weight: 700;
      font-size: 12px;
      color: #ffffff;
      border: none;
      border-right: 1px solid rgba(255, 255, 255, 0.15);
      text-transform: uppercase;
      letter-spacing: 1px;
      font-family: 'Open Sans', sans-serif;
      position: relative;
      z-index: 1;
      text-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);
    }

    .tablo-header-cell:last-child {
      border-right: none;
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
      max-height: 500px;
      overflow-y: auto;
      overflow-x: hidden;
      background: var(--surface-muted);
      scroll-behavior: smooth;
    }

    .tablo-body-container::-webkit-scrollbar {
      width: 6px;
    }
    .tablo-body-container::-webkit-scrollbar-track {
      background: transparent;
    }
    .tablo-body-container::-webkit-scrollbar-thumb {
      background: linear-gradient(180deg, var(--primary-light), var(--primary-color));
      border-radius: 999px;
    }
    .tablo-body-container::-webkit-scrollbar-thumb:hover {
      background: var(--primary-dark);
    }

    .yemek-tablo {
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      background-color: var(--surface);
      font-family: 'Open Sans', sans-serif;
      table-layout: fixed;
    }

    .yemek-tablo td {
      padding: 14px 16px;
      border: none;
      border-bottom: 1px solid var(--border-subtle);
      font-size: 12.5px;
      color: var(--text-primary);
      vertical-align: middle;
      line-height: 1.55;
      transition: background-color 0.25s var(--transition-smooth);
    }

    .yemek-tablo tbody tr:last-child td {
      border-bottom: none;
    }

    .yemek-tablo .tarih-col {
      width: 15%;
      text-align: center;
      font-weight: 600;
      background: linear-gradient(180deg, #f4f9fa 0%, #eef5f7 100%);
      vertical-align: middle;
      border-right: 1px solid var(--border-subtle);
    }

    .tarih-date {
      display: block;
      font-size: 16px;
      font-weight: 800;
      color: var(--primary-color);
      margin-bottom: 8px;
      letter-spacing: 0.3px;
    }

    .tarih-day {
      display: inline-block;
      font-size: 10px;
      font-weight: 700;
      color: var(--text-secondary);
      background: var(--white);
      padding: 5px 12px;
      border-radius: 999px;
      text-transform: uppercase;
      letter-spacing: 0.6px;
      border: 1px solid var(--border-subtle);
      box-shadow: 0 1px 3px rgba(26, 43, 51, 0.04);
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
      margin: 6px 0;
      padding-left: 16px;
      line-height: 1.55;
      color: var(--text-primary);
    }

    .yemek-tablo td ul li::before {
      content: '';
      position: absolute;
      left: 0;
      top: 0.62em;
      width: 5px;
      height: 5px;
      border-radius: 50%;
      background: linear-gradient(135deg, var(--primary-light), var(--primary-color));
      box-shadow: 0 0 4px var(--primary-glow);
    }

    .yemek-tablo tbody tr {
      transition: background-color 0.25s var(--transition-smooth), box-shadow 0.25s var(--transition-smooth);
    }

    .yemek-tablo tbody tr:nth-child(even) td:not(.tarih-col) {
      background-color: rgba(248, 251, 252, 0.6);
    }

    .yemek-tablo tbody tr:hover td {
      background-color: var(--primary-soft);
    }

    .yemek-tablo tbody tr:hover .tarih-col {
      background: linear-gradient(180deg, #e8f6f8 0%, #dff0f3 100%);
    }

    .yemek-tablo tbody tr.bugun-satir td {
      background-color: rgba(224, 247, 250, 0.85) !important;
    }

    .yemek-tablo tbody tr.bugun-satir {
      box-shadow: inset 4px 0 0 var(--primary-color);
    }

    .yemek-tablo tbody tr.bugun-satir .tarih-col {
      background: linear-gradient(180deg, #d4f1f5 0%, #c5ecf2 100%) !important;
    }

    .yemek-tablo tbody tr.bugun-satir .tarih-date {
      color: var(--primary-dark);
      font-size: 17px;
    }

    .yemek-tablo tbody tr.bugun-satir .tarih-day {
      background: linear-gradient(135deg, var(--primary-color), var(--primary-dark));
      color: var(--white);
      border-color: transparent;
      box-shadow: 0 2px 8px var(--shadow-medium);
    }

    .bos-mesaj {
      text-align: center;
      padding: 56px 32px;
      margin: 8px 0 0;
      width: 100%;
      color: var(--text-secondary);
      font-size: 15px;
      background: linear-gradient(135deg, var(--surface-muted), var(--surface));
      border: 1px dashed var(--border-light);
      border-radius: var(--radius-md);
    }

    .bos-mesaj-icon {
      display: flex;
      align-items: center;
      justify-content: center;
      width: 56px;
      height: 56px;
      margin: 0 auto 16px;
      border-radius: 50%;
      background: var(--primary-soft);
      color: var(--primary-color);
      font-size: 24px;
      line-height: 1;
    }

    .bos-mesaj p {
      max-width: 360px;
      margin: 0 auto;
      line-height: 1.6;
    }

    .bos-mesaj strong {
      display: block;
      font-size: 17px;
      font-weight: 700;
      color: var(--text-primary);
      margin-bottom: 8px;
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
              <span class="title-main">AYLIK YEMEK MEN&#220;S&#220;</span>
              <span class="title-month"><%= ay_adi %> <%= secilen_yil %></span>
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
              <div class="bos-mesaj-icon">&#127860;</div>
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
