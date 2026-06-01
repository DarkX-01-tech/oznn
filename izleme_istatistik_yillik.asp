<!-- #include file="database/Connection.asp" -->
<%
If Session("yemek_izleme_giris") <> "OK" Then
    Response.Redirect "izleme_giris.asp"
End If
Session.Timeout = 120

If Request.QueryString("yil") = "" Then
    Response.Redirect "izleme_liste.asp"
End If

Dim secilen_yil, menu_tipi, hedef_tablo, tip_param
secilen_yil = CInt(Request.QueryString("yil"))

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

Function GetDayName(d)
    Dim w: w = Weekday(d, 2)
    Select Case w
        Case 1: GetDayName = "Pazartesi"
        Case 2: GetDayName = "Sal&#305;"
        Case 3: GetDayName = "&#199;ar&#351;amba"
        Case 4: GetDayName = "Per&#351;embe"
        Case 5: GetDayName = "Cuma"
        Case 6: GetDayName = "Cumartesi"
        Case 7: GetDayName = "Pazar"
    End Select
End Function

Function SafeStr(val)
    If IsNull(val) Or IsEmpty(val) Then SafeStr = "" Else SafeStr = Trim(val & "")
End Function

Function SafeJsonStr(val)
    Dim s: s = SafeStr(val)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(10), "")
    SafeJsonStr = s
End Function

Dim rsYemek, sqlYemek
On Error Resume Next
sqlYemek = "SELECT * FROM " & hedef_tablo & " WHERE YEAR(tarih) = " & secilen_yil & " ORDER BY tarih ASC"
Set rsYemek = ConnYemek.Execute(sqlYemek)
If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Response.Redirect "izleme_liste.asp"
End If
On Error GoTo 0

If rsYemek.EOF Then
    If menu_tipi = "diyet" Then
        Response.Redirect "izleme_liste.asp?sekme=diyet"
    Else
        Response.Redirect "izleme_liste.asp"
    End If
End If

Dim dictTumu, toplam_gun, jsonArama
Set dictTumu = CreateObject("Scripting.Dictionary")
toplam_gun = 0
jsonArama = ""

Sub SayTumu(yemekAdi)
    Dim y: y = SafeStr(yemekAdi)
    If y <> "" Then
        Dim yUpper: yUpper = UCase(y)
        If dictTumu.Exists(yUpper) Then
            dictTumu(yUpper) = dictTumu(yUpper) + 1
        Else
            dictTumu.Add yUpper, 1
        End If
    End If
End Sub

Do While Not rsYemek.EOF
    toplam_gun = toplam_gun + 1
    SayTumu rsYemek("ogle_corba"): SayTumu rsYemek("ogle_ana_yemek"): SayTumu rsYemek("ogle_yan_urun"): SayTumu rsYemek("ogle_tatli")
    SayTumu rsYemek("aksam_corba"): SayTumu rsYemek("aksam_ana_yemek"): SayTumu rsYemek("aksam_yan_urun"): SayTumu rsYemek("aksam_tatli")

    Dim tmpTarih, tmpGun, tmpAy
    If Not IsNull(rsYemek("tarih")) And Not IsEmpty(rsYemek("tarih")) Then
        tmpTarih = Day(rsYemek("tarih")) & "." & Month(rsYemek("tarih")) & "." & Year(rsYemek("tarih"))
        tmpGun = GetDayName(CDate(rsYemek("tarih")))
        tmpAy = Month(rsYemek("tarih"))
    Else
        tmpTarih = "": tmpGun = "": tmpAy = 0
    End If

    If jsonArama <> "" Then jsonArama = jsonArama & ","
    jsonArama = jsonArama & "{""tarih"":""" & tmpTarih & """,""gun"":""" & tmpGun & """,""ay"":" & tmpAy & ","
    jsonArama = jsonArama & """ogle_corba"":""" & SafeJsonStr(rsYemek("ogle_corba")) & """,""ogle_ana"":""" & SafeJsonStr(rsYemek("ogle_ana_yemek")) & """,""ogle_yan"":""" & SafeJsonStr(rsYemek("ogle_yan_urun")) & """,""ogle_tatli"":""" & SafeJsonStr(rsYemek("ogle_tatli")) & ""","
    jsonArama = jsonArama & """aksam_corba"":""" & SafeJsonStr(rsYemek("aksam_corba")) & """,""aksam_ana"":""" & SafeJsonStr(rsYemek("aksam_ana_yemek")) & """,""aksam_yan"":""" & SafeJsonStr(rsYemek("aksam_yan_urun")) & """,""aksam_tatli"":""" & SafeJsonStr(rsYemek("aksam_tatli")) & """}"
    rsYemek.MoveNext
Loop
jsonArama = "[" & jsonArama & "]"

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

Dim sKeys
sKeys = GetSortedKeys(dictTumu)
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>(<%= secilen_yil %>) Y&#305;ll&#305;k <% If menu_tipi = "diyet" Then %>Diyet <% End If %>&#304;statistik</title>
  <link rel="icon" href="../../images/hastane_portal_logo.png" type="image/png">
  <style>
    :root { --main: <% If menu_tipi = "diyet" Then %>#4caf50<% Else %>#45b8c3<% End If %>; --dark: <% If menu_tipi = "diyet" Then %>#388e3c<% Else %>#2e8b91<% End If %>; --txt: #fff; }
    * { margin: 0; padding: 0; box-sizing: border-box; -webkit-user-select: none; -moz-user-select: none; -ms-user-select: none; user-select: none; }
    html, body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; min-height: 100vh; background: linear-gradient(135deg, #f0f9fa, #e6f4f1); }

    #header { background: linear-gradient(90deg, var(--main), var(--dark)); padding: 8px 15px; color: var(--txt); display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 10px rgba(0,0,0,0.15); position: sticky; top: 0; z-index: 100; }
    .header-left { display: flex; align-items: center; gap: 8px; }
    .nav-button { background: transparent; color: var(--txt); width: 36px; height: 36px; border: 2px solid var(--txt); border-radius: 50%; text-decoration: none; transition: all 0.3s; display: flex; align-items: center; justify-content: center; cursor: pointer; }
    .nav-button:hover { background: var(--dark); transform: translateY(-2px); }
    .nav-button svg { width: 20px; height: 20px; fill: currentColor; }
    #header h1 { font-size: 18px; margin: 0; text-shadow: 1px 1px 2px rgba(0,0,0,0.2); display: flex; align-items: center; gap: 8px; }
    #header h1 svg { width: 20px; height: 20px; opacity: 0.85; }
    .header-badge { background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 12px; font-size: 11px; font-weight: 600; }

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

    .kategori-card { background: #fff; border-radius: 12px; box-shadow: 0 4px 14px rgba(0,0,0,0.05); overflow: hidden; }
    .kategori-card-header { padding: 16px 20px; display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid #f0f0f0; }
    .kategori-card-header h3 { font-size: 15px; color: #2c3e50; display: flex; align-items: center; gap: 8px; }
    .kategori-card-header h3 svg { width: 18px; height: 18px; fill: var(--main); }
    .kategori-card-header .badge { background: var(--main); color: #fff; padding: 3px 12px; border-radius: 10px; font-size: 11px; font-weight: 600; }
    .kategori-card-body { max-height: 600px; overflow-y: auto; }
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
    .yemek-name { flex: 1; font-size: 13px; color: #333; font-weight: 500; cursor: pointer; }
    .yemek-name:hover { color: var(--main); text-decoration: underline; }
    .yemek-count { display: flex; align-items: center; gap: 8px; margin-right: 10px; }
    .yemek-count-num { background: <% If menu_tipi = "diyet" Then %>rgba(76,175,80,0.1)<% Else %>rgba(69,184,195,0.1)<% End If %>; color: var(--dark); padding: 4px 10px; border-radius: 8px; font-size: 12px; font-weight: 700; white-space: nowrap; }
    .yemek-bar-wrap { width: 80px; height: 6px; background: #f0f0f0; border-radius: 3px; overflow: hidden; }
    .yemek-bar { height: 100%; background: linear-gradient(90deg, var(--main), var(--dark)); border-radius: 3px; transition: width 0.6s ease; }

    .detay-btn { padding: 6px 14px; border: 2px solid #e0e0e0; border-radius: 20px; background: #fff; color: #666; font-size: 11px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; transition: all 0.3s; flex-shrink: 0; }
    .detay-btn:hover { border-color: var(--main); color: var(--main); transform: translateY(-2px); }
    .detay-btn svg { width: 13px; height: 13px; fill: currentColor; }

    .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); backdrop-filter: blur(3px); z-index: 1000; justify-content: center; align-items: flex-start; padding: 30px 15px; overflow-y: auto; }
    .modal-overlay.show { display: flex; }
    .modal-box { background: #fff; border-radius: 16px; width: 100%; max-width: 800px; box-shadow: 0 20px 60px rgba(0,0,0,0.25); overflow: hidden; animation: modalIn 0.3s ease; }
    @keyframes modalIn { from { opacity: 0; transform: translateY(-30px) scale(0.95); } to { opacity: 1; transform: translateY(0) scale(1); } }
    .modal-header { background: linear-gradient(90deg, var(--main), var(--dark)); color: #fff; padding: 14px 20px; display: flex; align-items: center; justify-content: space-between; }
    .modal-header h2 { font-size: 15px; margin: 0; display: flex; align-items: center; gap: 8px; font-weight: 600; }
    .modal-header h2 svg { width: 18px; height: 18px; fill: currentColor; }
    .modal-close { background: rgba(255,255,255,0.2); border: 2px solid rgba(255,255,255,0.4); color: #fff; width: 32px; height: 32px; border-radius: 50%; font-size: 18px; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: all 0.3s; }
    .modal-close:hover { background: rgba(255,255,255,0.35); transform: rotate(90deg); }
    .modal-body { padding: 18px; max-height: 70vh; overflow-y: auto; }
    .modal-body::-webkit-scrollbar { width: 4px; }
    .modal-body::-webkit-scrollbar-thumb { background: var(--main); border-radius: 10px; }

    .modal-ozet { display: flex; gap: 8px; margin-bottom: 14px; flex-wrap: wrap; }
    .modal-ozet-card { background: #f8f9fa; border-radius: 10px; padding: 10px 14px; text-align: center; border: 1px solid #e8e8e8; flex: 1; min-width: 70px; }
    .modal-ozet-card .mo-sayi { font-size: 18px; font-weight: 700; color: var(--dark); line-height: 1.2; }
    .modal-ozet-card .mo-baslik { font-size: 9px; color: #999; text-transform: uppercase; margin-top: 2px; }
    .modal-ozet-card.vurgulu { background: #fff8e1; border-color: #ffe082; }
    .modal-ozet-card.vurgulu .mo-sayi { color: #e65100; font-size: 13px; }

    .modal-tablo-wrap { background: #fff; border-radius: 10px; border: 1px solid #e8e8e8; overflow: hidden; }
    .modal-tablo { width: 100%; border-collapse: collapse; }
    .modal-tablo thead th { background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; padding: 9px 12px; font-size: 10px; text-align: left; font-weight: 600; text-transform: uppercase; }
    .modal-tablo tbody tr { border-bottom: 1px solid #f5f5f5; transition: background 0.2s; cursor: pointer; }
    .modal-tablo tbody tr:hover { background: #f0fffe; }
    .modal-tablo tbody td { padding: 8px 12px; font-size: 12px; color: #444; }
    .ay-link { color: var(--main); font-weight: 600; text-decoration: underline; cursor: pointer; }
    .ay-link:hover { color: var(--dark); }
    .ogun-badge { display: inline-block; padding: 3px 10px; border-radius: 8px; font-size: 10px; font-weight: 600; }
    .ogun-badge.ogle { background: rgba(69,184,195,0.15); color: #2e8b91; }
    .ogun-badge.aksam { background: rgba(156,39,176,0.12); color: #7b1fa2; }
    .kategori-badge { display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 10px; font-weight: 600; background: rgba(76,175,80,0.12); color: #2e7d32; }

    @media (max-width: 1000px) { .ozet-row { grid-template-columns: repeat(2, 1fr); } }
    @media (max-width: 600px) { .ozet-row { grid-template-columns: 1fr; } .modal-box { margin: 5px; } }
  </style>
</head>
<body>
<div id="header">
  <div class="header-left">
    <a href="izleme_panel.asp" class="nav-button" title="Panel"><svg viewBox="0 0 24 24"><path d="M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z"/></svg></a>
    <a href="izleme_liste.asp<% If menu_tipi = "diyet" Then %>?sekme=diyet<% End If %>" class="nav-button" title="Geri D&#246;n"><svg viewBox="0 0 24 24"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg></a>
  </div>
  <h1>
    <svg viewBox="0 0 24 24" fill="currentColor"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg>
    (<%= secilen_yil %>) Y&#305;ll&#305;k <% If menu_tipi = "diyet" Then %>Diyet <% End If %>Yemek &#304;statistikleri
    <% If menu_tipi = "diyet" Then %><span class="header-badge">Diyet</span><% End If %>
  </h1>
  <div style="width:36px;"></div>
</div>

<div class="content">
  <div class="ozet-row">
    <div class="ozet-card"><div class="ozet-icon bg1"><svg viewBox="0 0 24 24"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10z"/></svg></div><div class="ozet-info"><span>Toplam G&#252;n</span><strong><%= toplam_gun %></strong><small><%= secilen_yil %> Y&#305;l&#305;</small></div></div>
    <div class="ozet-card"><div class="ozet-icon bg2"><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg></div><div class="ozet-info"><span>Farkl&#305; Yemek</span><strong><%= toplam_cesit %></strong><small>&#199;e&#351;it Kullan&#305;lm&#305;&#351;</small></div></div>
    <div class="ozet-card"><div class="ozet-icon bg3"><svg viewBox="0 0 24 24"><path d="M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z"/></svg></div><div class="ozet-info"><span>En &#199;ok Tekrar</span><strong><%= en_cok_sayi %> Kez</strong><small><%= en_cok_yemek %></small></div></div>
    <div class="ozet-card"><div class="ozet-icon bg4"><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg></div><div class="ozet-info"><span>Tek Kullan&#305;m</span><strong><%= tek_kullanim %></strong><small>Yemek Sadece 1 Kez</small></div></div>
  </div>

  <div class="kategori-card">
    <div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg><%= secilen_yil %> Y&#305;l&#305; - T&#252;m Yemekler</h3><span class="badge"><%= dictTumu.Count %> &#199;e&#351;it</span></div>
    <div class="kategori-card-body">
      <% If IsArray(sKeys) And dictTumu.Count > 0 Then
        Dim idx, pct
        For idx = 0 To UBound(sKeys)
          pct = 0
          If en_cok_sayi > 0 Then pct = Round((dictTumu(sKeys(idx)) / en_cok_sayi) * 100)
      %>
      <div class="yemek-row">
        <div class="yemek-rank <%= GetRankClass(idx) %>"><%= idx + 1 %></div>
        <div class="yemek-name" onclick="detayGoster('<%= Replace(Replace(sKeys(idx), "'", "&#39;"), """", "&quot;") %>')"><%= sKeys(idx) %></div>
        <div class="yemek-count">
          <div class="yemek-bar-wrap"><div class="yemek-bar" style="width:<%= pct %>%;"></div></div>
          <span class="yemek-count-num"><%= dictTumu(sKeys(idx)) %> Kez</span>
        </div>
        <button class="detay-btn" onclick="detayGoster('<%= Replace(Replace(sKeys(idx), "'", "&#39;"), """", "&quot;") %>')">
          <svg viewBox="0 0 24 24"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>Detay
        </button>
      </div>
      <% Next
      End If %>
    </div>
  </div>
</div>

<div class="modal-overlay" id="detayModal" onclick="if(event.target===this) modalKapat();">
  <div class="modal-box">
    <div class="modal-header"><h2><svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg><span id="modalBaslik">Detay</span></h2><button class="modal-close" onclick="modalKapat();">&times;</button></div>
    <div class="modal-body" id="modalIcerik"></div>
  </div>
</div>

<script>
var tumVeri = <%= jsonArama %>;
var ayAdlari = ['','Ocak','\u015eubat','Mart','Nisan','May\u0131s','Haziran','Temmuz','A\u011fustos','Eyl\u00fcl','Ekim','Kas\u0131m','Aral\u0131k'];
var ogunKeys = ['ogle_corba','ogle_ana','ogle_yan','ogle_tatli','aksam_corba','aksam_ana','aksam_yan','aksam_tatli'];
var ogunLabels = {'ogle_corba':'\u00D6\u011Fle \u00C7orba','ogle_ana':'\u00D6\u011Fle Ana','ogle_yan':'\u00D6\u011Fle Yan','ogle_tatli':'\u00D6\u011Fle Tatl\u0131','aksam_corba':'Ak\u015Fam \u00C7orba','aksam_ana':'Ak\u015Fam Ana','aksam_yan':'Ak\u015Fam Yan','aksam_tatli':'Ak\u015Fam Tatl\u0131'};
var kategoriLabels = {'ogle_corba':'\u00C7orba','ogle_ana':'Ana Yemek','ogle_yan':'Yan \u00DCr\u00FCn','ogle_tatli':'Tatl\u0131/Meyve','aksam_corba':'\u00C7orba','aksam_ana':'Ana Yemek','aksam_yan':'Yan \u00DCr\u00FCn','aksam_tatli':'Tatl\u0131/Meyve'};

function decodeHtml(h){var d=document.createElement('div');d.innerHTML=h;return d.textContent||d.innerText||'';}

function detayGoster(yr) {
  var ya = decodeHtml(yr), al = ya.toUpperCase();
  var ayBazli = {}, toplamB = 0, ogunD = {};
  for (var o = 0; o < ogunKeys.length; o++) ogunD[ogunKeys[o]] = 0;

  for (var i = 0; i < tumVeri.length; i++) {
    var r = tumVeri[i];
    for (var k = 0; k < ogunKeys.length; k++) {
      if (r[ogunKeys[k]] && r[ogunKeys[k]].toUpperCase() === al) {
        toplamB++;
        ogunD[ogunKeys[k]]++;
        var ayNo = r.ay;
        if (!ayBazli[ayNo]) ayBazli[ayNo] = [];
        ayBazli[ayNo].push({ tarih: r.tarih, gun: r.gun, ogun: ogunKeys[k] });
      }
    }
  }

  document.getElementById('modalBaslik').innerHTML = '"' + ya + '" Y\u0131ll\u0131k Detay';

  if (toplamB === 0) { document.getElementById('modalIcerik').innerHTML = '<p style="text-align:center;color:#999;padding:40px;">Veri bulunamad\u0131.</p>'; document.getElementById('detayModal').classList.add('show'); document.body.style.overflow='hidden'; return; }

  var ayS = Object.keys(ayBazli).length;
  var ecO = '', ecOS = 0;
  for (var m = 0; m < ogunKeys.length; m++) { if (ogunD[ogunKeys[m]] > ecOS) { ecOS = ogunD[ogunKeys[m]]; ecO = ogunLabels[ogunKeys[m]]; } }

  var h = '<div class="modal-ozet">';
  h += '<div class="modal-ozet-card"><div class="mo-sayi">' + toplamB + '</div><div class="mo-baslik">Toplam</div></div>';
  h += '<div class="modal-ozet-card"><div class="mo-sayi">' + ayS + '</div><div class="mo-baslik">Farkl\u0131 Ay</div></div>';
  h += '<div class="modal-ozet-card vurgulu"><div class="mo-sayi">' + ecO + '</div><div class="mo-baslik">En \u00C7ok \u00D6\u011F\u00FCn</div></div>';
  h += '</div>';

  // Ay bazli tablo
  h += '<div class="modal-tablo-wrap"><table class="modal-tablo"><thead><tr><th>#</th><th>Ay</th><th>Tekrar</th><th></th></tr></thead><tbody>';
  var sira = 0;
  var ayKeys = Object.keys(ayBazli).sort(function(a,b){return parseInt(a)-parseInt(b);});
  for (var ai = 0; ai < ayKeys.length; ai++) {
    sira++;
    var ayK = ayKeys[ai];
    h += '<tr onclick="gunDetayGoster(\'' + ya.replace(/'/g,"\\'") + '\',' + ayK + ')" style="cursor:pointer;">';
    h += '<td>' + sira + '</td>';
    h += '<td><span class="ay-link">' + ayAdlari[parseInt(ayK)] + '</span></td>';
    h += '<td><strong>' + ayBazli[ayK].length + ' Kez</strong></td>';
    h += '<td><button class="detay-btn" onclick="event.stopPropagation();gunDetayGoster(\'' + ya.replace(/'/g,"\\'") + '\',' + ayK + ')"><svg viewBox="0 0 24 24"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>G\u00fcn Baz</button></td>';
    h += '</tr>';
  }
  h += '</tbody></table></div>';

  document.getElementById('modalIcerik').innerHTML = h;
  document.getElementById('detayModal').classList.add('show');
  document.body.style.overflow = 'hidden';
}

function gunDetayGoster(yemekAdi, ayNo) {
  var al = yemekAdi.toUpperCase();
  var bulunanlar = [];
  for (var i = 0; i < tumVeri.length; i++) {
    var r = tumVeri[i];
    if (r.ay !== ayNo) continue;
    for (var k = 0; k < ogunKeys.length; k++) {
      if (r[ogunKeys[k]] && r[ogunKeys[k]].toUpperCase() === al) {
        bulunanlar.push({ tarih: r.tarih, gun: r.gun, ogun: ogunKeys[k] });
      }
    }
  }

  document.getElementById('modalBaslik').innerHTML = '"' + yemekAdi + '" - ' + ayAdlari[ayNo] + ' G\u00fcn Detay\u0131';

  var h = '<div style="margin-bottom:10px;"><button class="detay-btn" onclick="detayGoster(\'' + yemekAdi.replace(/'/g,"&#39;") + '\')"><svg viewBox="0 0 24 24"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>Ay Baz\u0131na D\u00f6n</button></div>';
  h += '<div class="modal-tablo-wrap"><table class="modal-tablo"><thead><tr><th>#</th><th>Tarih</th><th>G\u00fcn</th><th>\u00D6\u011F\u00FCn</th><th>Kategori</th></tr></thead><tbody>';

  for (var p = 0; p < bulunanlar.length; p++) {
    var b = bulunanlar[p];
    var ot = b.ogun.indexOf('ogle') > -1 ? 'ogle' : 'aksam';
    var oa = ot === 'ogle' ? '\u00D6\u011Fle' : 'Ak\u015Fam';
    h += '<tr><td>' + (p+1) + '</td><td><strong>' + b.tarih + '</strong></td><td>' + b.gun + '</td><td><span class="ogun-badge ' + ot + '">' + oa + '</span></td><td><span class="kategori-badge">' + kategoriLabels[b.ogun] + '</span></td></tr>';
  }
  h += '</tbody></table></div>';

  document.getElementById('modalIcerik').innerHTML = h;
}

function modalKapat() { document.getElementById('detayModal').classList.remove('show'); document.body.style.overflow = ''; }
document.addEventListener('keydown', function(e) { if (e.key === 'Escape') modalKapat(); });
</script>
</body>
</html>
<%
rsYemek.Close: Set rsYemek = Nothing
Set dictTumu = Nothing
%>
