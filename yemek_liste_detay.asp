<!-- #include file="database/Connection.asp" -->
<%
If Session("yemek_admin_giris") <> "OK" Then
    Response.Redirect "giris.asp"
End If
Dim admin_kullanici
admin_kullanici = Session("yemek_admin_kullanici")
Session.Timeout = 30
If Request.QueryString("yil") = "" Or Request.QueryString("ay") = "" Then
    Response.Redirect "yemek_liste.asp"
End If
Dim secilen_yil, secilen_ay, menu_tipi
secilen_yil = Request.QueryString("yil")
secilen_ay = Request.QueryString("ay")
If Request.QueryString("tip") = "diyet" Then
    menu_tipi = "diyet"
Else
    menu_tipi = "normal"
End If
Dim hedef_tablo
If menu_tipi = "diyet" Then
    hedef_tablo = "diyet_yemek_listesi"
Else
    hedef_tablo = "yemek_listesi"
End If
Dim tip_param
If menu_tipi = "diyet" Then
    tip_param = "&tip=diyet"
Else
    tip_param = ""
End If
If Request.QueryString("islem") = "toplu_sil" Then
    On Error Resume Next
    Dim sqlTopluSil
    sqlTopluSil = "DELETE FROM " & hedef_tablo & " WHERE YEAR(tarih) = " & secilen_yil & " AND MONTH(tarih) = " & secilen_ay
    ConnYemek.Execute sqlTopluSil
    If Err.Number = 0 Then
        If menu_tipi = "diyet" Then
            Response.Redirect "yemek_liste.asp?sekme=diyet&durum=toplu_silindi"
        Else
            Response.Redirect "yemek_liste.asp?durum=toplu_silindi"
        End If
    Else
        Response.Redirect "yemek_liste_detay.asp?yil=" & secilen_yil & "&ay=" & secilen_ay & tip_param & "&durum=hata"
    End If
    On Error GoTo 0
End If
Dim rsYemek, sqlYemek
sqlYemek = "SELECT COUNT(*) AS toplam_gun FROM " & hedef_tablo & " WHERE YEAR(tarih) = " & secilen_yil & " AND MONTH(tarih) = " & secilen_ay
Set rsYemek = ConnYemek.Execute(sqlYemek)
Dim toplam_gun: toplam_gun = 0
If Not rsYemek.EOF Then toplam_gun = rsYemek("toplam_gun")
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
Dim ay_adi: ay_adi = GetMonthName(CInt(secilen_ay))
Dim basarili_mesaj, hata_mesaj
If Request.QueryString("durum") = "silindi" Then
    basarili_mesaj = "Kay&#305;t ba&#351;ar&#305;yla silindi!"
ElseIf Request.QueryString("durum") = "guncellendi" Then
    basarili_mesaj = "T&#252;m kay&#305;tlar ba&#351;ar&#305;yla g&#252;ncellendi!"
ElseIf Request.QueryString("durum") = "eklendi" Then
    basarili_mesaj = ay_adi & " " & secilen_yil & " i&#231;in t&#252;m g&#252;nler ba&#351;ar&#305;yla eklendi!"
ElseIf Request.QueryString("durum") = "hata" Then
    hata_mesaj = "&#304;&#351;lem s&#305;ras&#305;nda bir hata olu&#351;tu!"
End If
Dim rsLog, sqlLog, logHata: logHata = False
On Error Resume Next
sqlLog = "SELECT * FROM yemek_guncelleme_log WHERE yil = " & secilen_yil & " AND ay = " & secilen_ay & " ORDER BY guncelleme_tarihi DESC"
Set rsLog = ConnYemek.Execute(sqlLog)
If Err.Number <> 0 Then logHata = True
Err.Clear
On Error GoTo 0
Dim logSayisi: logSayisi = 0
If Not logHata Then
    If Not rsLog.EOF Then
        Dim rsLogCount
        On Error Resume Next
        Set rsLogCount = ConnYemek.Execute("SELECT COUNT(*) AS adet FROM yemek_guncelleme_log WHERE yil = " & secilen_yil & " AND ay = " & secilen_ay)
        If Err.Number = 0 And Not rsLogCount.EOF Then logSayisi = rsLogCount("adet")
        If Not rsLogCount Is Nothing Then rsLogCount.Close: Set rsLogCount = Nothing
        Err.Clear
        On Error GoTo 0
    End If
End If
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>(<%= secilen_yil %>) <%= ay_adi %> Ay&#305; <% If menu_tipi = "diyet" Then %>Diyet <% End If %>Yemek Men&#252;s&#252;</title>
  <link rel="icon" href="../../images/hastane_portal_logo.png" type="image/png">
  <style>
    :root { --main: <% If menu_tipi = "diyet" Then %>#4caf50<% Else %>#45b8c3<% End If %>; --dark: <% If menu_tipi = "diyet" Then %>#388e3c<% Else %>#2e8b91<% End If %>; --txt: #fff; }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    html, body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; height: 100vh; overflow: hidden; background: linear-gradient(135deg, #f0f9fa, #e6f4f1); }
    #header { background: linear-gradient(90deg, var(--main), var(--dark)); padding: 8px 15px; color: var(--txt); display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 10px rgba(0,0,0,0.15); }
    #nav-buttons { display: flex; gap: 8px; }
    .nav-button { background: transparent; color: var(--txt); width: 36px; height: 36px; border: 2px solid var(--txt); border-radius: 50%; text-decoration: none; transition: all 0.3s; display: flex; align-items: center; justify-content: center; cursor: pointer; }
    .nav-button:hover { background: var(--dark); transform: translateY(-2px); }
    .nav-button svg { width: 20px; height: 20px; fill: currentColor; }
    #header h1 { font-size: 17px; margin: 0; display: flex; align-items: center; gap: 6px; }
    #header h1 svg { width: 18px; height: 18px; opacity: 0.8; }
    .header-badge { background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 12px; font-size: 11px; font-weight: 600; }
    .page-wrapper { height: calc(100vh - 52px); display: flex; align-items: center; justify-content: center; }
    .alert { padding: 10px 15px; border-radius: 6px; font-size: 13px; display: flex; align-items: center; gap: 8px; animation: slideDown 0.3s; position: fixed; top: 60px; left: 50%; transform: translateX(-50%); z-index: 200; box-shadow: 0 4px 12px rgba(0,0,0,0.15); white-space: nowrap; }
    .alert svg { width: 18px; height: 18px; flex-shrink: 0; }
    .alert-success { background: #e8f5e9; color: #2e7d32; border-left: 4px solid #4caf50; }
    .alert-error { background: #ffebee; color: #c62828; border-left: 4px solid #f44336; }
    @keyframes slideDown { from { opacity: 0; transform: translateX(-50%) translateY(-10px); } to { opacity: 1; transform: translateX(-50%) translateY(0); } }
    .ay-card { background: #fff; border-radius: 16px; box-shadow: 0 10px 40px rgba(0,0,0,0.08); overflow: hidden; width: 540px; animation: cardIn 0.5s cubic-bezier(0.22,1,0.36,1); }
    @keyframes cardIn { from { opacity: 0; transform: translateY(20px) scale(0.96); } to { opacity: 1; transform: translateY(0) scale(1); } }
    .ay-card-header { background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; padding: 28px; text-align: center; position: relative; overflow: hidden; }
    .ay-card-header::before { content: ''; position: absolute; top: -30px; right: -30px; width: 120px; height: 120px; background: rgba(255,255,255,0.08); border-radius: 50%; }
    .ay-card-header::after { content: ''; position: absolute; bottom: -40px; left: -20px; width: 100px; height: 100px; background: rgba(255,255,255,0.05); border-radius: 50%; }
    .ay-card-header svg { width: 44px; height: 44px; margin-bottom: 8px; opacity: 0.9; position: relative; z-index: 1; }
    .ay-card-header h2 { font-size: 26px; margin: 0 0 4px 0; display: flex; align-items: center; justify-content: center; gap: 10px; position: relative; z-index: 1; }
    .ay-card-header h2 .sep { width: 20px; height: 20px; opacity: 0.5; }
    .ay-card-header p { font-size: 12px; opacity: 0.75; position: relative; z-index: 1; }
    .ay-card-header .type-label { display: inline-block; background: rgba(255,255,255,0.2); padding: 4px 14px; border-radius: 12px; font-size: 11px; font-weight: 600; margin-top: 8px; position: relative; z-index: 1; }
    .ay-card-body { padding: 28px; }
    .stat-row { display: flex; align-items: center; gap: 14px; background: linear-gradient(135deg, rgba(0,0,0,0.02), rgba(0,0,0,0.01)); border-left: 3px solid var(--main); border-radius: 8px; padding: 16px; margin-bottom: 22px; }
    .stat-icon { background: linear-gradient(135deg, var(--main), var(--dark)); width: 44px; height: 44px; border-radius: 12px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
    .stat-icon svg { width: 22px; height: 22px; fill: #fff; }
    .stat-text span { display: block; font-size: 11px; color: #999; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 2px; }
    .stat-text strong { font-size: 22px; color: #2c3e50; }
    .action-row { display: flex; gap: 10px; }
    .btn-act {
      flex: 1; padding: 16px 10px; border-radius: 12px;
      text-decoration: none; font-size: 12px; font-weight: 600;
      transition: all 0.3s cubic-bezier(0.22,1,0.36,1);
      display: flex; flex-direction: column; align-items: center;
      justify-content: center; gap: 8px; border: 2px solid transparent;
      cursor: pointer; text-align: center; position: relative; overflow: hidden;
    }
    .btn-act svg { width: 24px; height: 24px; flex-shrink: 0; }
    .btn-act span { line-height: 1.3; }
    .btn-act:hover { transform: translateY(-3px); }
    .btn-act:active { transform: translateY(0); }

    .btn-toplu { background: #fff; color: var(--dark); border-color: var(--main); }
    .btn-toplu svg { fill: var(--dark); }
    .btn-toplu:hover { background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; box-shadow: 0 8px 24px rgba(0,0,0,0.15); }
    .btn-toplu:hover svg { fill: #fff; }

    .btn-tekil { background: #fff; color: var(--dark); border-color: var(--main); }
    .btn-tekil svg { fill: var(--dark); }
    .btn-tekil:hover { background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; box-shadow: 0 8px 24px rgba(0,0,0,0.15); }
    .btn-tekil:hover svg { fill: #fff; }

    .btn-sil { background: #fff; color: #e53935; border-color: #ef9a9a; }
    .btn-sil svg { fill: #e53935; }
    .btn-sil:hover { background: linear-gradient(135deg, #ef5350, #d32f2f); color: #fff; border-color: #ef5350; box-shadow: 0 8px 24px rgba(229,57,53,0.3); }
    .btn-sil:hover svg { fill: #fff; }

    .btn-geri-act { background: #fff; color: #78909c; border-color: #cfd8dc; }
    .btn-geri-act svg { fill: #78909c; }
    .btn-geri-act:hover { background: linear-gradient(135deg, #78909c, #546e7a); color: #fff; border-color: #78909c; box-shadow: 0 8px 24px rgba(84,110,122,0.3); }
    .btn-geri-act:hover svg { fill: #fff; }

    .log-toggle { position: fixed; right: 0; top: 50%; transform: translateY(-50%); z-index: 60; background: linear-gradient(180deg, #546e7a, #37474f); color: #fff; border: none; border-radius: 10px 0 0 10px; padding: 14px 10px; cursor: pointer; box-shadow: -3px 0 12px rgba(0,0,0,0.12); transition: all 0.3s; writing-mode: vertical-lr; font-size: 11px; font-weight: 600; display: flex; align-items: center; gap: 8px; letter-spacing: 0.5px; }
    .log-toggle:hover { padding-right: 14px; background: linear-gradient(180deg, #455a64, #263238); }
    .log-toggle svg { width: 16px; height: 16px; flex-shrink: 0; fill: currentColor; }
    .log-toggle-count { background: var(--main); padding: 2px 6px; border-radius: 8px; font-size: 10px; writing-mode: horizontal-tb; }
    .log-panel { position: fixed; top: 52px; right: 0; width: 310px; height: calc(100vh - 52px); background: #fff; box-shadow: -6px 0 24px rgba(0,0,0,0.1); display: flex; flex-direction: column; z-index: 50; transform: translateX(100%); transition: transform 0.35s cubic-bezier(0.22,1,0.36,1); }
    .log-panel.open { transform: translateX(0); }
    .log-header { background: linear-gradient(135deg, #546e7a, #37474f); color: #fff; padding: 14px 16px; display: flex; align-items: center; justify-content: space-between; flex-shrink: 0; }
    .log-header h3 { font-size: 14px; margin: 0; display: flex; align-items: center; gap: 6px; }
    .log-header h3 svg { width: 18px; height: 18px; flex-shrink: 0; }
    .log-count { background: rgba(255,255,255,0.2); padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: 600; }
    .log-close { background: none; border: none; color: #fff; cursor: pointer; padding: 4px; border-radius: 4px; display: flex; align-items: center; justify-content: center; transition: background 0.2s; }
    .log-close:hover { background: rgba(255,255,255,0.15); }
    .log-close svg { width: 18px; height: 18px; fill: currentColor; }
    .log-body { flex: 1; overflow-y: auto; }
    .log-item { padding: 12px 16px; border-bottom: 1px solid #f0f0f0; transition: background 0.2s; }
    .log-item:hover { background: #f5fffe; }
    .log-item:last-child { border-bottom: none; }
    .log-item-top { display: flex; align-items: center; justify-content: space-between; margin-bottom: 5px; }
    .log-date { font-size: 12px; font-weight: 600; color: #2c3e50; }
    .log-badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: 600; }
    .log-badge.toplu { background: #e0f2f1; color: #00796b; }
    .log-badge.tekil { background: #e3f2fd; color: #1565c0; }
    .log-item-bottom { display: flex; align-items: center; justify-content: space-between; }
    .log-user { font-size: 11px; color: #888; }
    .log-desc { font-size: 11px; color: #aaa; }
    .log-empty { text-align: center; padding: 40px 20px; color: #bbb; }
    .log-empty svg { width: 36px; height: 36px; fill: #ddd; margin-bottom: 10px; display: block; margin-left: auto; margin-right: auto; }
    .log-empty p { font-size: 12px; font-style: italic; line-height: 1.6; }
    @media (max-width: 600px) {
      .ay-card { width: calc(100vw - 30px); }
      .action-row { flex-wrap: wrap; }
      .btn-act { flex: 1 1 45%; }
      .log-panel { width: 260px; }
    }
  </style>
  <script>
    function topluSil() {
      if (confirm('T\u00dcM KAYITLARI silmek istedi\u011finize emin misiniz?\n\n<%= ay_adi %> <%= secilen_yil %> - <%= toplam_gun %> g\u00fcn silinecek!\n\nGER\u0130 ALINAMAZ!')) {
        window.location.href = 'yemek_liste_detay.asp?yil=<%= secilen_yil %>&ay=<%= secilen_ay %><%= tip_param %>&islem=toplu_sil';
      }
    }
    function toggleLog() {
      var p = document.getElementById('logPanel');
      var t = document.getElementById('logToggle');
      p.classList.toggle('open');
      t.style.display = p.classList.contains('open') ? 'none' : 'flex';
    }
    setTimeout(function(){ var a = document.querySelector('.alert'); if(a) a.style.display='none'; }, 4000);
  </script>
</head>
<body>
  <div id="header">
    <div id="nav-buttons">
      <a href="panel.asp" class="nav-button" title="Ana Sayfa"><svg viewBox="0 0 24 24"><path d="M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z"/></svg></a>
      <a href="yemek_liste.asp<% If menu_tipi = "diyet" Then %>?sekme=diyet<% End If %>" class="nav-button" title="Aylara D&#246;n"><svg viewBox="0 0 24 24"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg></a>
    </div>
    <h1>
      <%= ay_adi %>
      <svg viewBox="0 0 24 24"><path d="M9 11H7v2h2v-2zm4 0h-2v2h2v-2zm4 0h-2v2h2v-2zm2-7h-1V2h-2v2H8V2H6v2H5c-1.11 0-1.99.9-1.99 2L3 20c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 16H5V9h14v11z" fill="currentColor"/></svg>
      <%= secilen_yil %>
      <% If menu_tipi = "diyet" Then %><span class="header-badge">Diyet</span><% End If %>
    </h1>
    <div style="width:36px;"></div>
  </div>
  <% If basarili_mesaj <> "" Then %>
  <div class="alert alert-success"><svg viewBox="0 0 24 24" fill="currentColor"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><%= basarili_mesaj %></div>
  <% End If %>
  <% If hata_mesaj <> "" Then %>
  <div class="alert alert-error"><svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg><%= hata_mesaj %></div>
  <% End If %>
  <div class="page-wrapper">
    <div class="ay-card">
      <div class="ay-card-header">
        <% If menu_tipi = "diyet" Then %>
        <svg viewBox="0 0 24 24" fill="currentColor"><path d="M17.21 9l-4.38-6.56c-.19-.28-.51-.42-.83-.42-.32 0-.64.14-.83.43L6.79 9C6.3 9.71 6 10.57 6 11.5 6 14.53 8.47 17 11.5 17h1c3.03 0 5.5-2.47 5.5-5.5 0-.93-.3-1.79-.79-2.5zM12 15.5c-1.93 0-3.5-1.57-3.5-3.5 0-.53.2-1.33.57-1.88L12 5.97l2.93 4.15c.37.55.57 1.35.57 1.88 0 1.93-1.57 3.5-3.5 3.5z"/></svg>
        <% Else %>
        <svg viewBox="0 0 24 24" fill="currentColor"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10z"/></svg>
        <% End If %>
        <h2><%= ay_adi %> <svg viewBox="0 0 24 24" class="sep"><path d="M9 11H7v2h2v-2zm4 0h-2v2h2v-2zm4 0h-2v2h2v-2zm2-7h-1V2h-2v2H8V2H6v2H5c-1.11 0-1.99.9-1.99 2L3 20c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 16H5V9h14v11z" fill="currentColor"/></svg> <%= secilen_yil %></h2>
        <p><% If menu_tipi = "diyet" Then %>Ayl&#305;k Diyet Yemek Men&#252;s&#252; Y&#246;netimi<% Else %>Ayl&#305;k Yemek Men&#252;s&#252; Y&#246;netimi<% End If %></p>
        <% If menu_tipi = "diyet" Then %><span class="type-label">Diyet Men&#252;s&#252;</span><% End If %>
      </div>
      <div class="ay-card-body">
        <div class="stat-row">
          <div class="stat-icon"><svg viewBox="0 0 24 24"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10z"/></svg></div>
          <div class="stat-text"><span>Toplam G&#252;n</span><strong><%= toplam_gun %> G&#252;n</strong></div>
        </div>
        <div class="action-row">
          <a href="yemek_toplu_duzenle.asp?yil=<%= secilen_yil %>&ay=<%= secilen_ay %><%= tip_param %>" class="btn-act btn-toplu">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
            <span>Toplu<br>D&#252;zenle</span>
          </a>
          <a href="yemek_duzenle_ay.asp?yil=<%= secilen_yil %>&ay=<%= secilen_ay %><%= tip_param %>" class="btn-act btn-tekil">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8l-6-6zM6 20V4h7v5h5v11H6zm8-4H8v-2h6v2zm2-4H8v-2h8v2z"/></svg>
            <span>Tek Tek<br>D&#252;zenle</span>
          </a>
          <button onclick="topluSil();" class="btn-act btn-sil">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
            <span>Men&#252;y&#252;<br>Sil</span>
          </button>
          <a href="yemek_liste.asp<% If menu_tipi = "diyet" Then %>?sekme=diyet<% End If %>" class="btn-act btn-geri-act">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>
            <span>Listeye<br>D&#246;n</span>
          </a>
        </div>
      </div>
    </div>
  </div>
  <button class="log-toggle" id="logToggle" onclick="toggleLog()">
    <svg viewBox="0 0 24 24"><path d="M13 3c-4.97 0-9 4.03-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42C8.27 19.99 10.51 21 13 21c4.97 0 9-4.03 9-9s-4.03-9-9-9zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>
    G&#252;ncelleme Ge&#231;mi&#351;i
    <% If logSayisi > 0 Then %><span class="log-toggle-count"><%= logSayisi %></span><% End If %>
  </button>
  <div class="log-panel" id="logPanel">
    <div class="log-header">
      <h3><svg viewBox="0 0 24 24" fill="currentColor"><path d="M13 3c-4.97 0-9 4.03-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42C8.27 19.99 10.51 21 13 21c4.97 0 9-4.03 9-9s-4.03-9-9-9zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>G&#252;ncelleme Ge&#231;mi&#351;i</h3>
      <div style="display:flex;align-items:center;gap:8px;">
        <% If logSayisi > 0 Then %><span class="log-count"><%= logSayisi %> kay&#305;t</span><% End If %>
        <button class="log-close" onclick="toggleLog()"><svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
      </div>
    </div>
    <div class="log-body">
      <% If logHata Then %>
      <div class="log-empty">
        <svg viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg>
        <p>Log tablosu hen&#252;z olu&#351;turulmam&#305;&#351;.</p>
      </div>
      <% ElseIf rsLog.EOF Then %>
      <div class="log-empty">
        <svg viewBox="0 0 24 24"><path d="M13 3c-4.97 0-9 4.03-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42C8.27 19.99 10.51 21 13 21c4.97 0 9-4.03 9-9s-4.03-9-9-9zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>
        <p>Hen&#252;z g&#252;ncelleme yap&#305;lmam&#305;&#351;.</p>
      </div>
      <% Else %>
      <% Do While Not rsLog.EOF %>
      <div class="log-item">
        <div class="log-item-top">
          <span class="log-date"><%= FormatDateTime(rsLog("guncelleme_tarihi"), 0) %></span>
          <span class="log-badge <% If InStr(rsLog("guncelleme_tipi") & "", "Toplu") > 0 Then %>toplu<% Else %>tekil<% End If %>"><%= rsLog("guncelleme_tipi") %></span>
        </div>
        <div class="log-item-bottom">
          <span class="log-user"><%= rsLog("kullanici") %></span>
          <span class="log-desc"><%= rsLog("aciklama") %></span>
        </div>
      </div>
      <% rsLog.MoveNext
      Loop %>
      <% End If %>
    </div>
  </div>
</body>
</html>
<%
rsYemek.Close
Set rsYemek = Nothing
If Not logHata Then
  If Not rsLog Is Nothing Then rsLog.Close: Set rsLog = Nothing
End If
%>
