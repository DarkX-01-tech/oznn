<!-- #include file="database/Connection.asp" -->
<%
If Session("yemek_izleme_giris") <> "OK" Then
    Response.Redirect "izleme_giris.asp"
End If
Session.Timeout = 120

'====================================
' Aktif Sekme Kontrolü
'====================================
Dim aktif_sekme
If Request.QueryString("sekme") = "diyet" Then
    aktif_sekme = "diyet"
Else
    aktif_sekme = "normal"
End If

Dim rsAylar, sqlAylar

If aktif_sekme = "diyet" Then
    sqlAylar = "SELECT YEAR(tarih) AS yil, MONTH(tarih) AS ay, " & _
               "COUNT(*) AS kayit_sayisi, " & _
               "MIN(tarih) AS ilk_tarih, " & _
               "MAX(aktif) AS aktif_var " & _
               "FROM diyet_yemek_listesi " & _
               "GROUP BY YEAR(tarih), MONTH(tarih) " & _
               "ORDER BY YEAR(tarih) DESC, MONTH(tarih) DESC"
Else
    sqlAylar = "SELECT YEAR(tarih) AS yil, MONTH(tarih) AS ay, " & _
               "COUNT(*) AS kayit_sayisi, " & _
               "MIN(tarih) AS ilk_tarih, " & _
               "MAX(aktif) AS aktif_var " & _
               "FROM yemek_listesi " & _
               "GROUP BY YEAR(tarih), MONTH(tarih) " & _
               "ORDER BY YEAR(tarih) DESC, MONTH(tarih) DESC"
End If

Set rsAylar = ConnYemek.Execute(sqlAylar)

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
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Yemek Listeleri - &#304;zleme Paneli</title>
  <link rel="icon" href="../../images/hastane_portal_logo.png" type="image/png">
  <style>
    :root { --main: <% If aktif_sekme = "diyet" Then %>#4caf50<% Else %>#45b8c3<% End If %>; --dark: <% If aktif_sekme = "diyet" Then %>#388e3c<% Else %>#2e8b91<% End If %>; --txt: #ffffff; --diyet: #4caf50; --diyet-dark: #388e3c; }
    * { margin: 0; padding: 0; box-sizing: border-box; -webkit-user-select: none; -moz-user-select: none; -ms-user-select: none; user-select: none; }
    html, body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; height: 100vh; overflow: hidden; background: linear-gradient(135deg, #f0f9fa, #e6f4f1); }
    input { -webkit-user-select: text; -moz-user-select: text; -ms-user-select: text; user-select: text; }

    #header { background: linear-gradient(90deg, var(--main), var(--dark)); padding: 8px 15px; color: var(--txt); display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 10px rgba(0,0,0,0.15); }
    .header-left { display: flex; align-items: center; gap: 8px; }
    .nav-button { background: transparent; color: var(--txt); width: 36px; height: 36px; border: 2px solid var(--txt); border-radius: 50%; text-decoration: none; transition: all 0.3s; display: flex; align-items: center; justify-content: center; cursor: pointer; }
    .nav-button:hover { background: var(--dark); transform: translateY(-2px); }
    .nav-button svg { width: 20px; height: 20px; fill: currentColor; }
    #header h1 { font-size: 18px; margin: 0; text-shadow: 1px 1px 2px rgba(0,0,0,0.2); display: flex; align-items: center; gap: 8px; }
    #header h1 svg { width: 20px; height: 20px; opacity: 0.85; }
    .search-wrapper { position: relative; }
    .search-wrapper svg { position: absolute; top: 50%; left: 10px; transform: translateY(-50%); width: 16px; height: 16px; fill: #999; pointer-events: none; }
    #search-bar { padding: 6px 30px; border-radius: 20px; border: 1px solid rgba(255,255,255,0.3); width: 200px; font-size: 13px; background: rgba(255,255,255,0.9); transition: all 0.3s; }
    #search-bar:focus { outline: none; width: 250px; background: #fff; box-shadow: 0 0 8px rgba(255,255,255,0.5); }

    .content { height: calc(100vh - 52px); overflow-y: auto; padding: 20px; }
    .content::-webkit-scrollbar { width: 6px; }
    .content::-webkit-scrollbar-track { background: #f1f1f1; }
    .content::-webkit-scrollbar-thumb { background: var(--main); border-radius: 10px; }

    /* SEKME */
    .sekme-wrapper {
      display: flex; gap: 4px; margin-bottom: 20px;
      background: #fff; border-radius: 12px; padding: 5px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.06); max-width: fit-content;
    }
    .sekme-btn {
      padding: 10px 22px; border-radius: 10px;
      font-size: 14px; font-weight: 600; cursor: pointer;
      display: flex; align-items: center; gap: 8px;
      text-decoration: none; transition: all 0.3s;
      border: 2px solid transparent; color: #777; background: transparent;
    }
    .sekme-btn:hover { color: #333; background: #f5f5f5; }
    .sekme-btn svg { width: 18px; height: 18px; fill: currentColor; }
    .sekme-btn.aktif-normal {
      background: linear-gradient(135deg, #45b8c3, #2e8b91);
      color: #fff; border-color: transparent;
      box-shadow: 0 4px 12px rgba(69,184,195,0.3);
    }
    .sekme-btn.aktif-normal svg { fill: #fff; }
    .sekme-btn.aktif-diyet {
      background: linear-gradient(135deg, var(--diyet), var(--diyet-dark));
      color: #fff; border-color: transparent;
      box-shadow: 0 4px 12px rgba(76,175,80,0.3);
    }
    .sekme-btn.aktif-diyet svg { fill: #fff; }

    .liste-container { background: #fff; border-radius: 12px; box-shadow: 0 6px 20px rgba(0,0,0,0.06); overflow: hidden; }
    .liste-header { background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; padding: 18px 22px; display: flex; justify-content: space-between; align-items: center; }
    .liste-header h2 { font-size: 18px; margin: 0; display: flex; align-items: center; gap: 10px; }
    .liste-header h2 svg { width: 24px; height: 24px; }

    .ay-kartlari { display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 20px; padding: 22px; }
    .ay-kart { background: #fff; border: 2px solid #e8e8e8; border-radius: 14px; padding: 0; transition: all 0.35s cubic-bezier(0.22,1,0.36,1); overflow: hidden; }
    .ay-kart:hover { border-color: var(--main); box-shadow: 0 12px 30px <% If aktif_sekme = "diyet" Then %>rgba(76,175,80,0.18)<% Else %>rgba(69,184,195,0.18)<% End If %>; transform: translateY(-5px); }
    .ay-kart-top { background: linear-gradient(135deg, <% If aktif_sekme = "diyet" Then %>rgba(76,175,80,0.06), rgba(56,142,60,0.03)<% Else %>rgba(69,184,195,0.06), rgba(46,139,145,0.03)<% End If %>); padding: 20px 20px 16px 20px; border-bottom: 1px solid #f0f0f0; }
    .ay-kart-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px; }
    .ay-kart-title { display: flex; align-items: center; gap: 12px; }
    .ay-kart-icon { width: 46px; height: 46px; border-radius: 12px; background: linear-gradient(135deg, var(--main), var(--dark)); display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
    .ay-kart-icon svg { width: 24px; height: 24px; fill: #fff; }
    .ay-kart-title h3 { font-size: 20px; color: #2c3e50; margin: 0; line-height: 1.2; }
    .ay-kart-title h3 span { font-size: 15px; color: #999; font-weight: 400; }
    .ay-kart-badge { background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; padding: 6px 14px; border-radius: 20px; font-size: 12px; font-weight: 700; }
    .ay-kart-info { display: flex; flex-direction: column; gap: 6px; }
    .info-item { display: flex; align-items: center; gap: 8px; font-size: 13px; color: #777; }
    .info-item svg { width: 15px; height: 15px; fill: var(--main); flex-shrink: 0; }

    .ay-kart-footer { padding: 14px 20px; display: flex; gap: 8px; }
    .btn-detay, .btn-istatistik { flex: 1; padding: 11px 10px; border-radius: 10px; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.3s; display: flex; align-items: center; justify-content: center; gap: 6px; text-decoration: none; background: #fff; color: var(--dark); border: 2px solid var(--main); }
    .btn-detay svg, .btn-istatistik svg { width: 16px; height: 16px; flex-shrink: 0; fill: var(--dark); }
    .btn-detay:hover, .btn-istatistik:hover { background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; border-color: var(--main); transform: translateY(-2px); box-shadow: 0 6px 16px <% If aktif_sekme = "diyet" Then %>rgba(76,175,80,0.4)<% Else %>rgba(69,184,195,0.4)<% End If %>; }
    .btn-detay:hover svg, .btn-istatistik:hover svg { fill: #fff; }

    .bos-durum { padding: 80px 30px; text-align: center; color: #999; }
    .bos-durum svg { width: 80px; height: 80px; fill: #ddd; margin-bottom: 20px; }
    .bos-durum h3 { font-size: 20px; color: #666; margin: 0 0 10px 0; }
    .bos-durum p { font-size: 14px; margin: 0; }

    @media (max-width: 768px) {
      #header { flex-wrap: wrap; padding: 10px; }
      #header h1 { font-size: 15px; }
      #search-bar { width: 150px; }
      #search-bar:focus { width: 180px; }
      .ay-kartlari { grid-template-columns: 1fr; }
      .ay-kart-footer { flex-direction: column; }
      .sekme-wrapper { width: 100%; }
      .sekme-btn { flex: 1; justify-content: center; padding: 10px 14px; font-size: 13px; }
    }
  </style>
  <script>
    function searchTable() {
      var input = document.getElementById("search-bar");
      var filter = input.value.toLowerCase();
      var cards = document.querySelectorAll(".ay-kart");
      for (var i = 0; i < cards.length; i++) {
        var txt = cards[i].textContent || cards[i].innerText;
        cards[i].style.display = (txt.toLowerCase().indexOf(filter) > -1) ? "" : "none";
      }
    }
  </script>
</head>
<body>
  <div id="header">
    <div class="header-left">
      <a href="izleme_panel.asp" class="nav-button" title="Panel"><svg viewBox="0 0 24 24"><path d="M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z"/></svg></a>
      <a href="izleme_panel.asp" class="nav-button" title="Panele D&#246;n"><svg viewBox="0 0 24 24"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg></a>
    </div>
    <h1><svg viewBox="0 0 24 24" fill="currentColor"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>Yemek Listeleri (Ayl&#305;k)</h1>
    <div class="search-wrapper">
      <svg viewBox="0 0 24 24"><path d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>
      <input type="text" id="search-bar" placeholder="Ay Ara..." onkeyup="searchTable()">
    </div>
  </div>

  <div class="content">

    <!-- SEKME -->
    <div class="sekme-wrapper">
      <a href="izleme_liste.asp?sekme=normal" class="sekme-btn<% If aktif_sekme = "normal" Then %> aktif-normal<% End If %>">
        <svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>
        Yemek Listesi
      </a>
      <a href="izleme_liste.asp?sekme=diyet" class="sekme-btn<% If aktif_sekme = "diyet" Then %> aktif-diyet<% End If %>">
        <svg viewBox="0 0 24 24"><path d="M17.21 9l-4.38-6.56c-.19-.28-.51-.42-.83-.42-.32 0-.64.14-.83.43L6.79 9C6.3 9.71 6 10.57 6 11.5 6 14.53 8.47 17 11.5 17h1c3.03 0 5.5-2.47 5.5-5.5 0-.93-.3-1.79-.79-2.5z"/></svg>
        Diyet Yemek Listesi
      </a>
    </div>

    <div class="liste-container">
      <div class="liste-header">
        <h2>
          <% If aktif_sekme = "diyet" Then %>
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M17.21 9l-4.38-6.56c-.19-.28-.51-.42-.83-.42-.32 0-.64.14-.83.43L6.79 9C6.3 9.71 6 10.57 6 11.5 6 14.53 8.47 17 11.5 17h1c3.03 0 5.5-2.47 5.5-5.5 0-.93-.3-1.79-.79-2.5z"/></svg>
          Ayl&#305;k Diyet Yemek Men&#252;leri
          <% Else %>
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10z"/></svg>
          Ayl&#305;k Yemek Men&#252;leri
          <% End If %>
        </h2>
      </div>
      <% If Not rsAylar.EOF Then %>
      <div class="ay-kartlari">
        <% Do While Not rsAylar.EOF
        Dim ay_adi
        ay_adi = GetMonthName(rsAylar("ay")) %>
        <div class="ay-kart">
          <div class="ay-kart-top">
            <div class="ay-kart-header">
              <div class="ay-kart-title">
                <div class="ay-kart-icon">
                  <% If aktif_sekme = "diyet" Then %>
                  <svg viewBox="0 0 24 24"><path d="M17.21 9l-4.38-6.56c-.19-.28-.51-.42-.83-.42-.32 0-.64.14-.83.43L6.79 9C6.3 9.71 6 10.57 6 11.5 6 14.53 8.47 17 11.5 17h1c3.03 0 5.5-2.47 5.5-5.5 0-.93-.3-1.79-.79-2.5z"/></svg>
                  <% Else %>
                  <svg viewBox="0 0 24 24"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10z"/></svg>
                  <% End If %>
                </div>
                <div><h3><%= ay_adi %> <span>/ <%= rsAylar("yil") %></span></h3></div>
              </div>
              <div class="ay-kart-badge"><%= rsAylar("kayit_sayisi") %> G&#252;n</div>
            </div>
            <div class="ay-kart-info">
              <div class="info-item"><svg viewBox="0 0 24 24" fill="currentColor"><path d="M11.99 2C6.47 2 2 6.48 2 12s4.47 10 9.99 10C17.52 22 22 17.52 22 12S17.52 2 11.99 2zM12 20c-4.42 0-8-3.58-8-8s3.58-8 8-8 8 3.58 8 8-3.58 8-8 8z"/><path d="M12.5 7H11v6l5.25 3.15.75-1.23-4.5-2.67z"/></svg><span>Toplam <%= rsAylar("kayit_sayisi") %> Kay&#305;t</span></div>
              <% If rsAylar("aktif_var") = True Then %>
              <div class="info-item"><svg viewBox="0 0 24 24" fill="currentColor"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><span>Aktif Kay&#305;tlar Mevcut</span></div>
              <% End If %>
            </div>
          </div>
          <div class="ay-kart-footer">
            <a href="javascript:void(0);" onclick="openMenu(<%= rsAylar("yil") %>, <%= rsAylar("ay") %>, '<%= aktif_sekme %>')" class="btn-detay">
              <svg viewBox="0 0 24 24" fill="currentColor"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>Men&#252;y&#252; G&#246;r
            </a>
            <a href="izleme_istatistik_sayfa.asp?yil=<%= rsAylar("yil") %>&ay=<%= rsAylar("ay") %><% If aktif_sekme = "diyet" Then %>&tip=diyet<% End If %>" class="btn-istatistik">
              <svg viewBox="0 0 24 24" fill="currentColor"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg>&#304;statistik
            </a>
          </div>
        </div>
        <% rsAylar.MoveNext
        Loop %>
      </div>
      <% Else %>
      <div class="bos-durum">
        <% If aktif_sekme = "diyet" Then %>
        <svg viewBox="0 0 24 24"><path d="M17.21 9l-4.38-6.56c-.19-.28-.51-.42-.83-.42-.32 0-.64.14-.83.43L6.79 9C6.3 9.71 6 10.57 6 11.5 6 14.53 8.47 17 11.5 17h1c3.03 0 5.5-2.47 5.5-5.5 0-.93-.3-1.79-.79-2.5z"/></svg>
        <h3>Hen&#252;z Diyet Men&#252;s&#252; Kayd&#305; Bulunmuyor</h3>
        <p>Diyet Yemek Men&#252;s&#252; Hen&#252;z Eklenmemi&#351;</p>
        <% Else %>
        <svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/></svg>
        <h3>Hen&#252;z Kay&#305;t Bulunmuyor</h3>
        <p>Yemek Men&#252;s&#252; Hen&#252;z Eklenmemi&#351;</p>
        <% End If %>
      </div>
      <% End If %>
    </div>
  </div>

  <!-- MENU MODAL -->
  <div id="menuModal" class="modal-overlay" onclick="if(event.target===this) closeMenu();">
    <div class="modal-box">
      <div class="modal-header" id="menuModalHeader">
        <h2><svg viewBox="0 0 24 24" fill="currentColor"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg><span id="menuModalTitle">Yemek Men&#252;s&#252;</span></h2>
        <button class="modal-close" onclick="closeMenu();">&times;</button>
      </div>
      <div class="modal-body"><iframe id="menuFrame" src=""></iframe></div>
    </div>
  </div>

  <style>
    .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); backdrop-filter: blur(3px); -webkit-backdrop-filter: blur(3px); z-index: 1000; justify-content: center; align-items: center; }
    .modal-overlay.show { display: flex; }
    .modal-box { background: #fff; border-radius: 14px; width: 900px; max-width: 95%; height: 88vh; box-shadow: 0 20px 60px rgba(0,0,0,0.3); overflow: hidden; animation: modalIn 0.3s ease; }
    @keyframes modalIn { from { opacity: 0; transform: translateY(-30px) scale(0.95); } to { opacity: 1; transform: translateY(0) scale(1); } }
    .modal-header { background: linear-gradient(90deg, var(--main), var(--dark)); color: #fff; padding: 14px 20px; display: flex; align-items: center; justify-content: space-between; }
    .modal-header.diyet-modal-header { background: linear-gradient(90deg, #4caf50, #388e3c); }
    .modal-header h2 { font-size: 16px; margin: 0; display: flex; align-items: center; gap: 10px; font-weight: 600; }
    .modal-header h2 svg { width: 20px; height: 20px; }
    .modal-close { background: rgba(255,255,255,0.2); border: 2px solid rgba(255,255,255,0.4); color: #fff; width: 32px; height: 32px; border-radius: 50%; font-size: 18px; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: all 0.3s; }
    .modal-close:hover { background: rgba(255,255,255,0.35); transform: rotate(90deg); }
    .modal-body { height: calc(88vh - 58px); overflow: hidden; }
    .modal-body iframe { width: 100%; height: 100%; border: none; }
  </style>

  <script>
    function openMenu(yil, ay, tip) {
      var modal = document.getElementById('menuModal');
      var header = document.getElementById('menuModalHeader');
      var ayAdlari = ['', 'Ocak', '\u015eubat', 'Mart', 'Nisan', 'May\u0131s', 'Haziran', 'Temmuz', 'A\u011fustos', 'Eyl\u00fcl', 'Ekim', 'Kas\u0131m', 'Aral\u0131k'];
      if (tip === 'diyet') {
        document.getElementById('menuModalTitle').textContent = ayAdlari[ay] + ' ' + yil + ' - Diyet Yemek Men\u00fcs\u00fc';
        header.className = 'modal-header diyet-modal-header';
      } else {
        document.getElementById('menuModalTitle').textContent = ayAdlari[ay] + ' ' + yil + ' - Yemek Men\u00fcs\u00fc';
        header.className = 'modal-header';
      }
      modal.classList.add('show');
      document.body.style.overflow = 'hidden';
      var url = 'yemek_gecmis.asp?yil=' + yil + '&ay=' + ay;
      if (tip === 'diyet') url += '&tip=diyet';
      document.getElementById('menuFrame').src = url;
    }
    function closeMenu() {
      var modal = document.getElementById('menuModal');
      modal.classList.remove('show');
      document.getElementById('menuFrame').src = '';
      document.body.style.overflow = 'auto';
    }
    document.addEventListener('keydown', function(e) { if (e.key === 'Escape') closeMenu(); });
  </script>
</body>
</html>
<% rsAylar.Close
Set rsAylar = Nothing %>
