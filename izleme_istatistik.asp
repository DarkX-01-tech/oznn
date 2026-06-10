<!-- #include file="database/Connection.asp" -->
<%
If Session("yemek_izleme_giris") <> "OK" Then
    Response.Redirect "izleme_giris.asp"
End If
Session.Timeout = 120

If Request.QueryString("yil") = "" Or Request.QueryString("ay") = "" Then
    Response.Redirect "izleme_liste.asp"
End If

Dim secilen_yil, secilen_ay, menu_tipi, hedef_tablo, tip_param
secilen_yil = Request.QueryString("yil")
secilen_ay = Request.QueryString("ay")

If Request.QueryString("tip") = "diyet" Then
    menu_tipi = "diyet"
    hedef_tablo = "diyet_yemek_listesi"
    tip_param = "&tip=diyet"
Else
    menu_tipi = "normal"
    hedef_tablo = "yemek_listesi"
    tip_param = ""
End If

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
ay_adi = GetMonthName(CInt(secilen_ay))

Dim rsYemek, sqlYemek
sqlYemek = "SELECT * FROM " & hedef_tablo & " WHERE YEAR(tarih) = " & secilen_yil & " AND MONTH(tarih) = " & secilen_ay & " ORDER BY tarih ASC"
Set rsYemek = ConnYemek.Execute(sqlYemek)

If rsYemek.EOF Then
    If menu_tipi = "diyet" Then
        Response.Redirect "izleme_liste.asp?sekme=diyet"
    Else
        Response.Redirect "izleme_liste.asp"
    End If
End If

Dim dictOgleCorba, dictOgleAna, dictOgleYan, dictOgleTatli
Dim dictAksamCorba, dictAksamAna, dictAksamYan, dictAksamTatli
Dim dictTumu

Set dictOgleCorba = CreateObject("Scripting.Dictionary")
Set dictOgleAna = CreateObject("Scripting.Dictionary")
Set dictOgleYan = CreateObject("Scripting.Dictionary")
Set dictOgleTatli = CreateObject("Scripting.Dictionary")
Set dictAksamCorba = CreateObject("Scripting.Dictionary")
Set dictAksamAna = CreateObject("Scripting.Dictionary")
Set dictAksamYan = CreateObject("Scripting.Dictionary")
Set dictAksamTatli = CreateObject("Scripting.Dictionary")
Set dictTumu = CreateObject("Scripting.Dictionary")

Dim toplam_gun
toplam_gun = 0

Sub SayYemek(dict, yemekAdi)
    Dim y: y = Trim(yemekAdi & "")
    If y <> "" Then
        If dict.Exists(y) Then dict(y) = dict(y) + 1 Else dict.Add y, 1
    End If
End Sub

Sub SayTumu(yemekAdi)
    Dim y: y = Trim(yemekAdi & "")
    If y <> "" Then
        If dictTumu.Exists(y) Then dictTumu(y) = dictTumu(y) + 1 Else dictTumu.Add y, 1
    End If
End Sub

Do While Not rsYemek.EOF
    toplam_gun = toplam_gun + 1
    SayYemek dictOgleCorba, rsYemek("ogle_corba")
    SayYemek dictOgleAna, rsYemek("ogle_ana_yemek")
    SayYemek dictOgleYan, rsYemek("ogle_yan_urun")
    SayYemek dictOgleTatli, rsYemek("ogle_tatli")
    SayYemek dictAksamCorba, rsYemek("aksam_corba")
    SayYemek dictAksamAna, rsYemek("aksam_ana_yemek")
    SayYemek dictAksamYan, rsYemek("aksam_yan_urun")
    SayYemek dictAksamTatli, rsYemek("aksam_tatli")
    SayTumu rsYemek("ogle_corba"): SayTumu rsYemek("ogle_ana_yemek"): SayTumu rsYemek("ogle_yan_urun"): SayTumu rsYemek("ogle_tatli")
    SayTumu rsYemek("aksam_corba"): SayTumu rsYemek("aksam_ana_yemek"): SayTumu rsYemek("aksam_yan_urun"): SayTumu rsYemek("aksam_tatli")
    rsYemek.MoveNext
Loop

Function GetSortedKeys(dict)
    Dim keys, counts, i, j, tmpK, tmpC
    If dict.Count = 0 Then GetSortedKeys = Array(): Exit Function
    keys = dict.Keys
    ReDim counts(dict.Count - 1)
    For i = 0 To dict.Count - 1: counts(i) = dict(keys(i)): Next
    For i = 0 To UBound(keys) - 1
        For j = 0 To UBound(keys) - 1 - i
            If counts(j) < counts(j + 1) Then
                tmpK = keys(j): keys(j) = keys(j + 1): keys(j + 1) = tmpK
                tmpC = counts(j): counts(j) = counts(j + 1): counts(j + 1) = tmpC
            End If
        Next
    Next
    GetSortedKeys = keys
End Function

Function GetRankClass(idx)
    Select Case idx
        Case 0: GetRankClass = "r1"
        Case 1: GetRankClass = "r2"
        Case 2: GetRankClass = "r3"
        Case Else: GetRankClass = "rn"
    End Select
End Function

Dim toplam_cesit: toplam_cesit = dictTumu.Count
Dim en_cok_yemek, en_cok_sayi, kk
en_cok_yemek = "-": en_cok_sayi = 0
For Each kk In dictTumu.Keys
    If dictTumu(kk) > en_cok_sayi Then en_cok_sayi = dictTumu(kk): en_cok_yemek = kk
Next

Dim tek_kullanim: tek_kullanim = 0
For Each kk In dictTumu.Keys
    If dictTumu(kk) = 1 Then tek_kullanim = tek_kullanim + 1
Next

Sub RenderKategori(dict, maxVal)
    Dim sKeys, idx, pct, localMax
    sKeys = GetSortedKeys(dict)
    If IsArray(sKeys) And dict.Count > 0 Then
        localMax = maxVal
        If localMax = 0 Then localMax = dict(sKeys(0))
        For idx = 0 To UBound(sKeys)
            pct = 0
            If localMax > 0 Then pct = Round((dict(sKeys(idx)) / localMax) * 100)
%>
          <div class="yemek-row">
            <div class="yemek-rank <%= GetRankClass(idx) %>"><%= idx + 1 %></div>
            <div class="yemek-name"><%= sKeys(idx) %></div>
            <div class="yemek-count">
              <div class="yemek-bar-wrap"><div class="yemek-bar" style="width:<%= pct %>%;"></div></div>
              <span class="yemek-count-num"><%= dict(sKeys(idx)) %> Kez</span>
            </div>
          </div>
<%
        Next
    Else
%>
          <div class="empty-msg">Veri Bulunamad&#305;</div>
<%
    End If
End Sub
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>(<%= secilen_yil %>) <%= ay_adi %> <% If menu_tipi = "diyet" Then %>Diyet <% End If %>&#304;statistik - &#304;zleme</title>
  <link rel="icon" href="../../images/hastane_portal_logo.png" type="image/png">
  <style>
    :root { --main: <% If menu_tipi = "diyet" Then %>#4caf50<% Else %>#45b8c3<% End If %>; --dark: <% If menu_tipi = "diyet" Then %>#388e3c<% Else %>#2e8b91<% End If %>; --txt: #fff; }
    * { margin: 0; padding: 0; box-sizing: border-box; -webkit-user-select: none; -moz-user-select: none; -ms-user-select: none; user-select: none; }
    html, body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; min-height: 100vh; background: linear-gradient(135deg, #f0f9fa, #e6f4f1); }

    .content { padding: 20px; max-width: 1400px; margin: 0 auto; }

    .ozet-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
    .ozet-card { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 4px 14px rgba(0,0,0,0.05); display: flex; align-items: center; gap: 15px; transition: all 0.3s; }
    .ozet-card:hover { transform: translateY(-3px); box-shadow: 0 8px 22px rgba(0,0,0,0.1); }
    .ozet-icon { width: 50px; height: 50px; border-radius: 12px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
    .ozet-icon svg { width: 26px; height: 26px; fill: #fff; }
    .ozet-icon.bg1 { background: linear-gradient(135deg, var(--main), var(--dark)); }
    .ozet-icon.bg2 { background: linear-gradient(135deg, #ff9800, #e65100); }
    .ozet-icon.bg3 { background: linear-gradient(135deg, #e91e63, #ad1457); }
    .ozet-icon.bg4 { background: linear-gradient(135deg, <% If menu_tipi = "diyet" Then %>#45b8c3, #2e8b91<% Else %>#4caf50, #2e7d32<% End If %>); }
    .ozet-info span { display: block; font-size: 11px; color: #999; text-transform: uppercase; letter-spacing: 0.3px; }
    .ozet-info strong { font-size: 22px; color: #2c3e50; display: block; }
    .ozet-info small { font-size: 11px; color: #888; }

    .filter-bar { display: flex; gap: 8px; margin-bottom: 20px; flex-wrap: wrap; }
    .filter-btn { padding: 8px 16px; border: 2px solid #e0e0e0; border-radius: 20px; background: #fff; color: #666; font-size: 12px; font-weight: 600; cursor: pointer; transition: all 0.3s; display: inline-flex; align-items: center; gap: 5px; }
    .filter-btn:hover { border-color: var(--main); color: var(--main); }
    .filter-btn.active { background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; border-color: var(--main); }
    .filter-btn svg { width: 14px; height: 14px; fill: currentColor; }

    .kategori-section { display: none; animation: fadeIn 0.3s; }
    .kategori-section.active { display: block; }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }

    .kategori-card { background: #fff; border-radius: 12px; box-shadow: 0 4px 14px rgba(0,0,0,0.05); overflow: hidden; }
    .kategori-card-header { padding: 16px 20px; display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid #f0f0f0; }
    .kategori-card-header h3 { font-size: 15px; color: #2c3e50; display: flex; align-items: center; gap: 8px; }
    .kategori-card-header h3 svg { width: 18px; height: 18px; fill: var(--main); }
    .kategori-card-header .badge { background: var(--main); color: #fff; padding: 3px 12px; border-radius: 10px; font-size: 11px; font-weight: 600; }
    .kategori-card-body { max-height: 500px; overflow-y: auto; }
    .kategori-card-body::-webkit-scrollbar { width: 5px; }
    .kategori-card-body::-webkit-scrollbar-thumb { background: var(--main); border-radius: 10px; }

    .yemek-row { display: flex; align-items: center; padding: 11px 20px; border-bottom: 1px solid #f8f8f8; transition: background 0.2s; }
    .yemek-row:hover { background: #f5fffe; }
    .yemek-row:last-child { border-bottom: none; }
    .yemek-rank { width: 28px; height: 28px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 700; flex-shrink: 0; margin-right: 12px; }
    .yemek-rank.r1 { background: linear-gradient(135deg, #ffd700, #ffb300); color: #5d4037; }
    .yemek-rank.r2 { background: linear-gradient(135deg, #b0bec5, #78909c); color: #fff; }
    .yemek-rank.r3 { background: linear-gradient(135deg, #a1887f, #795548); color: #fff; }
    .yemek-rank.rn { background: #f0f0f0; color: #888; }
    .yemek-name { flex: 1; font-size: 13px; color: #333; font-weight: 500; }
    .yemek-count { display: flex; align-items: center; gap: 8px; }
    .yemek-count-num { background: <% If menu_tipi = "diyet" Then %>rgba(76,175,80,0.1)<% Else %>rgba(69,184,195,0.1)<% End If %>; color: var(--dark); padding: 4px 10px; border-radius: 8px; font-size: 12px; font-weight: 700; white-space: nowrap; }
    .yemek-bar-wrap { width: 80px; height: 6px; background: #f0f0f0; border-radius: 3px; overflow: hidden; }
    .yemek-bar { height: 100%; background: linear-gradient(90deg, var(--main), var(--dark)); border-radius: 3px; transition: width 0.6s ease; }

    .empty-msg { text-align: center; padding: 40px; color: #bbb; font-style: italic; }

    @media (max-width: 1000px) { .ozet-row { grid-template-columns: repeat(2, 1fr); } }
    @media (max-width: 600px) { .ozet-row { grid-template-columns: 1fr; } }
  </style>
</head>
<body>
<div class="content">

  <div class="ozet-row">
    <div class="ozet-card">
      <div class="ozet-icon bg1"><svg viewBox="0 0 24 24"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10z"/></svg></div>
      <div class="ozet-info"><span>Toplam G&#252;n</span><strong><%= toplam_gun %></strong><small><%= ay_adi %> / <%= secilen_yil %></small></div>
    </div>
    <div class="ozet-card">
      <div class="ozet-icon bg2"><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg></div>
      <div class="ozet-info"><span>Farkl&#305; Yemek</span><strong><%= toplam_cesit %></strong><small>&#199;e&#351;it Kullan&#305;lm&#305;&#351;</small></div>
    </div>
    <div class="ozet-card">
      <div class="ozet-icon bg3"><svg viewBox="0 0 24 24"><path d="M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z"/></svg></div>
      <div class="ozet-info"><span>En &#199;ok Tekrar</span><strong><%= en_cok_sayi %> Kez</strong><small><%= en_cok_yemek %></small></div>
    </div>
    <div class="ozet-card">
      <div class="ozet-icon bg4"><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg></div>
      <div class="ozet-info"><span>Tek Kullan&#305;m</span><strong><%= tek_kullanim %></strong><small>Yemek Sadece 1 Kez</small></div>
    </div>
  </div>

  <div class="filter-bar">
    <button class="filter-btn active" onclick="showTab('tumu', this)"><svg viewBox="0 0 24 24"><path d="M3 13h2v-2H3v2zm0 4h2v-2H3v2zm0-8h2V7H3v2zm4 4h14v-2H7v2zm0 4h14v-2H7v2zM7 7v2h14V7H7z"/></svg>T&#252;m&#252;</button>
    <button class="filter-btn" onclick="showTab('ogle_corba', this)">&#214;&#287;le &#199;orba</button>
    <button class="filter-btn" onclick="showTab('ogle_ana', this)">&#214;&#287;le Ana</button>
    <button class="filter-btn" onclick="showTab('ogle_yan', this)">&#214;&#287;le Yan</button>
    <button class="filter-btn" onclick="showTab('ogle_tatli', this)">&#214;&#287;le Tatl&#305;</button>
    <button class="filter-btn" onclick="showTab('aksam_corba', this)">Ak&#351;am &#199;orba</button>
    <button class="filter-btn" onclick="showTab('aksam_ana', this)">Ak&#351;am Ana</button>
    <button class="filter-btn" onclick="showTab('aksam_yan', this)">Ak&#351;am Yan</button>
    <button class="filter-btn" onclick="showTab('aksam_tatli', this)">Ak&#351;am Tatl&#305;</button>
  </div>

  <div class="kategori-section active" id="sec_tumu">
    <div class="kategori-card">
      <div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg>T&#252;m Yemekler - Tekrar Say&#305;s&#305;na G&#246;re</h3><span class="badge"><%= dictTumu.Count %> &#199;e&#351;it</span></div>
      <div class="kategori-card-body"><% RenderKategori dictTumu, en_cok_sayi %></div>
    </div>
  </div>

  <div class="kategori-section" id="sec_ogle_corba">
    <div class="kategori-card">
      <div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>&#214;&#287;le &#199;orbas&#305;</h3><span class="badge"><%= dictOgleCorba.Count %> &#199;e&#351;it</span></div>
      <div class="kategori-card-body"><% RenderKategori dictOgleCorba, 0 %></div>
    </div>
  </div>

  <div class="kategori-section" id="sec_ogle_ana">
    <div class="kategori-card">
      <div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>&#214;&#287;le Ana Yemek</h3><span class="badge"><%= dictOgleAna.Count %> &#199;e&#351;it</span></div>
      <div class="kategori-card-body"><% RenderKategori dictOgleAna, 0 %></div>
    </div>
  </div>

  <div class="kategori-section" id="sec_ogle_yan">
    <div class="kategori-card">
      <div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>&#214;&#287;le Yan &#220;r&#252;n</h3><span class="badge"><%= dictOgleYan.Count %> &#199;e&#351;it</span></div>
      <div class="kategori-card-body"><% RenderKategori dictOgleYan, 0 %></div>
    </div>
  </div>

  <div class="kategori-section" id="sec_ogle_tatli">
    <div class="kategori-card">
      <div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>&#214;&#287;le Tatl&#305;/Meyve</h3><span class="badge"><%= dictOgleTatli.Count %> &#199;e&#351;it</span></div>
      <div class="kategori-card-body"><% RenderKategori dictOgleTatli, 0 %></div>
    </div>
  </div>

  <div class="kategori-section" id="sec_aksam_corba">
    <div class="kategori-card">
      <div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M20 8.69V4h-4.69L12 .69 8.69 4H4v4.69L.69 12 4 15.31V20h4.69L12 23.31 15.31 20H20v-4.69L23.31 12 20 8.69zM12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6 6 2.69 6 6-2.69 6-6 6zm0-10c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4-1.79-4-4-4z"/></svg>Ak&#351;am &#199;orbas&#305;</h3><span class="badge"><%= dictAksamCorba.Count %> &#199;e&#351;it</span></div>
      <div class="kategori-card-body"><% RenderKategori dictAksamCorba, 0 %></div>
    </div>
  </div>

  <div class="kategori-section" id="sec_aksam_ana">
    <div class="kategori-card">
      <div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M20 8.69V4h-4.69L12 .69 8.69 4H4v4.69L.69 12 4 15.31V20h4.69L12 23.31 15.31 20H20v-4.69L23.31 12 20 8.69zM12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6 6 2.69 6 6-2.69 6-6 6zm0-10c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4-1.79-4-4-4z"/></svg>Ak&#351;am Ana Yemek</h3><span class="badge"><%= dictAksamAna.Count %> &#199;e&#351;it</span></div>
      <div class="kategori-card-body"><% RenderKategori dictAksamAna, 0 %></div>
    </div>
  </div>

  <div class="kategori-section" id="sec_aksam_yan">
    <div class="kategori-card">
      <div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M20 8.69V4h-4.69L12 .69 8.69 4H4v4.69L.69 12 4 15.31V20h4.69L12 23.31 15.31 20H20v-4.69L23.31 12 20 8.69zM12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6 6 2.69 6 6-2.69 6-6 6zm0-10c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4-1.79-4-4-4z"/></svg>Ak&#351;am Yan &#220;r&#252;n</h3><span class="badge"><%= dictAksamYan.Count %> &#199;e&#351;it</span></div>
      <div class="kategori-card-body"><% RenderKategori dictAksamYan, 0 %></div>
    </div>
  </div>

  <div class="kategori-section" id="sec_aksam_tatli">
    <div class="kategori-card">
      <div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M20 8.69V4h-4.69L12 .69 8.69 4H4v4.69L.69 12 4 15.31V20h4.69L12 23.31 15.31 20H20v-4.69L23.31 12 20 8.69zM12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6 6 2.69 6 6-2.69 6-6 6zm0-10c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4-1.79-4-4-4z"/></svg>Ak&#351;am Tatl&#305;/Meyve</h3><span class="badge"><%= dictAksamTatli.Count %> &#199;e&#351;it</span></div>
      <div class="kategori-card-body"><% RenderKategori dictAksamTatli, 0 %></div>
    </div>
  </div>

</div>

<script>
function showTab(tab, btn) {
  var secs = document.querySelectorAll('.kategori-section');
  for (var i = 0; i < secs.length; i++) secs[i].classList.remove('active');
  document.getElementById('sec_' + tab).classList.add('active');
  var btns = document.querySelectorAll('.filter-btn');
  for (var j = 0; j < btns.length; j++) btns[j].classList.remove('active');
  btn.classList.add('active');
}
</script>
</body>
</html>
<%
rsYemek.Close
Set rsYemek = Nothing
Set dictOgleCorba = Nothing: Set dictOgleAna = Nothing: Set dictOgleYan = Nothing: Set dictOgleTatli = Nothing
Set dictAksamCorba = Nothing: Set dictAksamAna = Nothing: Set dictAksamYan = Nothing: Set dictAksamTatli = Nothing
Set dictTumu = Nothing
%>
