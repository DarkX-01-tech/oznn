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

Function GetDayName(d)
    Dim w
    w = Weekday(d, 2)
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

Dim toplam_gun: toplam_gun = 0

Sub SayYemek(dict, yemekAdi)
    Dim y: y = SafeStr(yemekAdi)
    If y <> "" Then
        If dict.Exists(y) Then dict(y) = dict(y) + 1 Else dict.Add y, 1
    End If
End Sub

Sub SayTumu(yemekAdi)
    Dim y: y = SafeStr(yemekAdi)
    If y <> "" Then
        If dictTumu.Exists(y) Then dictTumu(y) = dictTumu(y) + 1 Else dictTumu.Add y, 1
    End If
End Sub

Dim jsonArama: jsonArama = ""

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

    Dim tmpTarih, tmpGun
    If Not IsNull(rsYemek("tarih")) And Not IsEmpty(rsYemek("tarih")) Then
        tmpTarih = Day(rsYemek("tarih")) & "." & Month(rsYemek("tarih")) & "." & Year(rsYemek("tarih"))
        tmpGun = GetDayName(CDate(rsYemek("tarih")))
    Else
        tmpTarih = "": tmpGun = ""
    End If

    If jsonArama <> "" Then jsonArama = jsonArama & ","
    jsonArama = jsonArama & "{""tarih"":""" & tmpTarih & """,""gun"":""" & tmpGun & """,""ogle_corba"":""" & SafeJsonStr(rsYemek("ogle_corba")) & """,""ogle_ana"":""" & SafeJsonStr(rsYemek("ogle_ana_yemek")) & """,""ogle_yan"":""" & SafeJsonStr(rsYemek("ogle_yan_urun")) & """,""ogle_tatli"":""" & SafeJsonStr(rsYemek("ogle_tatli")) & """,""aksam_corba"":""" & SafeJsonStr(rsYemek("aksam_corba")) & """,""aksam_ana"":""" & SafeJsonStr(rsYemek("aksam_ana_yemek")) & """,""aksam_yan"":""" & SafeJsonStr(rsYemek("aksam_yan_urun")) & """,""aksam_tatli"":""" & SafeJsonStr(rsYemek("aksam_tatli")) & """}"
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
            <button class="detay-btn" onclick="detayGoster('<%= Replace(Replace(sKeys(idx), "'", "&#39;"), """", "&quot;") %>')">
              <svg viewBox="0 0 24 24"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>
              Detay
            </button>
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
  <title>(<%= secilen_yil %>) <%= ay_adi %> <% If menu_tipi = "diyet" Then %>Diyet <% End If %>&#304;statistik</title>
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
    .yemek-count { display: flex; align-items: center; gap: 8px; margin-right: 10px; }
    .yemek-count-num { background: <% If menu_tipi = "diyet" Then %>rgba(76,175,80,0.1)<% Else %>rgba(69,184,195,0.1)<% End If %>; color: var(--dark); padding: 4px 10px; border-radius: 8px; font-size: 12px; font-weight: 700; white-space: nowrap; }
    .yemek-bar-wrap { width: 80px; height: 6px; background: #f0f0f0; border-radius: 3px; overflow: hidden; }
    .yemek-bar { height: 100%; background: linear-gradient(90deg, var(--main), var(--dark)); border-radius: 3px; transition: width 0.6s ease; }
    .empty-msg { text-align: center; padding: 40px; color: #bbb; font-style: italic; }

    .detay-btn { padding: 6px 14px; border: 2px solid #e0e0e0; border-radius: 20px; background: #fff; color: #666; font-size: 11px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; transition: all 0.3s; flex-shrink: 0; }
    .detay-btn:hover { border-color: var(--main); color: var(--main); transform: translateY(-2px); box-shadow: 0 4px 12px <% If menu_tipi = "diyet" Then %>rgba(76,175,80,0.15)<% Else %>rgba(69,184,195,0.15)<% End If %>; }
    .detay-btn:active { transform: translateY(0); }
    .detay-btn svg { width: 13px; height: 13px; fill: currentColor; }

    .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.45); backdrop-filter: blur(3px); -webkit-backdrop-filter: blur(3px); z-index: 1000; justify-content: center; align-items: flex-start; padding: 40px 15px; overflow-y: auto; }
    .modal-overlay.show { display: flex; }
    .modal-box { background: linear-gradient(135deg, #f0f9fa, #e6f4f1); border-radius: 16px; width: 100%; max-width: 750px; box-shadow: 0 20px 60px rgba(0,0,0,0.2); overflow: hidden; animation: modalSlide 0.35s cubic-bezier(0.34, 1.56, 0.64, 1); }
    @keyframes modalSlide { from { opacity: 0; transform: translateY(-40px) scale(0.95); } to { opacity: 1; transform: translateY(0) scale(1); } }
    .modal-header { background: linear-gradient(90deg, var(--main), var(--dark)); color: #fff; padding: 14px 20px; display: flex; align-items: center; justify-content: space-between; }
    .modal-header h2 { font-size: 15px; display: flex; align-items: center; gap: 8px; margin: 0; font-weight: 600; }
    .modal-header h2 svg { width: 18px; height: 18px; fill: currentColor; opacity: 0.9; }
    .modal-close { background: rgba(255,255,255,0.2); border: 2px solid rgba(255,255,255,0.4); color: #fff; width: 32px; height: 32px; border-radius: 50%; font-size: 16px; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: all 0.3s; }
    .modal-close:hover { background: rgba(255,255,255,0.35); transform: rotate(90deg); }
    .modal-body { padding: 18px; max-height: 65vh; overflow-y: auto; }
    .modal-body::-webkit-scrollbar { width: 4px; }
    .modal-body::-webkit-scrollbar-thumb { background: var(--main); border-radius: 10px; }

    .modal-ozet { display: flex; gap: 8px; margin-bottom: 14px; flex-wrap: nowrap; overflow-x: auto; }
    .modal-ozet-card { background: #fff; border-radius: 10px; padding: 10px 14px; text-align: center; border: 1px solid #e8e8e8; flex: 1; min-width: 70px; }
    .modal-ozet-card .mo-sayi { font-size: 18px; font-weight: 700; color: var(--dark); line-height: 1.2; }
    .modal-ozet-card .mo-baslik { font-size: 9px; color: #999; text-transform: uppercase; margin-top: 2px; letter-spacing: 0.3px; }
    .modal-ozet-card.vurgulu { background: #fff8e1; border-color: #ffe082; }
    .modal-ozet-card.vurgulu .mo-sayi { color: #e65100; font-size: 13px; font-weight: 700; }

    .modal-ogun-bar { background: #fff; border-radius: 10px; border: 1px solid #e8e8e8; padding: 10px 14px; margin-bottom: 14px; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
    .modal-ogun-label { font-size: 10px; color: #999; text-transform: uppercase; font-weight: 600; margin-right: 4px; }
    .ogun-badge { display: inline-block; padding: 3px 10px; border-radius: 8px; font-size: 10px; font-weight: 600; }
    .ogun-badge.ogle { background: rgba(69,184,195,0.15); color: #2e8b91; }
    .ogun-badge.aksam { background: rgba(156,39,176,0.12); color: #7b1fa2; }

    .modal-tablo-wrap { background: #fff; border-radius: 10px; border: 1px solid #e8e8e8; overflow: hidden; }
    .modal-tablo { width: 100%; border-collapse: collapse; }
    .modal-tablo thead th { background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; padding: 9px 12px; font-size: 10px; text-align: left; font-weight: 600; text-transform: uppercase; letter-spacing: 0.4px; }
    .modal-tablo tbody tr { border-bottom: 1px solid #f5f5f5; transition: background 0.2s; }
    .modal-tablo tbody tr:last-child { border-bottom: none; }
    .modal-tablo tbody tr:hover { background: #f0fffe; }
    .modal-tablo tbody td { padding: 8px 12px; font-size: 12px; color: #444; }
    .modal-tablo tbody td:first-child { color: #ccc; font-weight: 700; font-size: 11px; }
    .kategori-badge { display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 10px; font-weight: 600; background: rgba(76,175,80,0.12); color: #2e7d32; }

    @media (max-width: 1000px) { .ozet-row { grid-template-columns: repeat(2, 1fr); } }
    @media (max-width: 600px) { .ozet-row { grid-template-columns: 1fr; } .modal-box { margin: 5px; } .modal-ozet { flex-wrap: wrap; } .modal-ozet-card { min-width: 60px; } }
  </style>
</head>
<body>
<div id="header">
  <div class="header-left">
    <a href="izleme_panel.asp" class="nav-button" title="Panel"><svg viewBox="0 0 24 24"><path d="M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z"/></svg></a>
    <a href="izleme_liste.asp<% If menu_tipi = "diyet" Then %>?sekme=diyet<% End If %>" class="nav-button" title="Yemek Listeleri"><svg viewBox="0 0 24 24"><path d="M3 13h2v-2H3v2zm0 4h2v-2H3v2zm0-8h2V7H3v2zm4 4h14v-2H7v2zm0 4h14v-2H7v2zM7 7v2h14V7H7z"/></svg></a>
    <a href="izleme_liste.asp<% If menu_tipi = "diyet" Then %>?sekme=diyet<% End If %>" class="nav-button" title="Geri D&#246;n"><svg viewBox="0 0 24 24"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg></a>
  </div>
  <h1>
    <svg viewBox="0 0 24 24" fill="currentColor"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg>
    (<%= secilen_yil %>) <%= ay_adi %> <% If menu_tipi = "diyet" Then %>Diyet <% End If %>Yemek &#304;statistikleri
    <% If menu_tipi = "diyet" Then %><span class="header-badge">Diyet</span><% End If %>
  </h1>
  <div style="width:36px;"></div>
</div>

<div class="content">
  <div class="ozet-row">
    <div class="ozet-card"><div class="ozet-icon bg1"><svg viewBox="0 0 24 24"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10z"/></svg></div><div class="ozet-info"><span>Toplam G&#252;n</span><strong><%= toplam_gun %></strong><small><%= ay_adi %> / <%= secilen_yil %></small></div></div>
    <div class="ozet-card"><div class="ozet-icon bg2"><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg></div><div class="ozet-info"><span>Farkl&#305; Yemek</span><strong><%= toplam_cesit %></strong><small>&#199;e&#351;it Kullan&#305;lm&#305;&#351;</small></div></div>
    <div class="ozet-card"><div class="ozet-icon bg3"><svg viewBox="0 0 24 24"><path d="M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z"/></svg></div><div class="ozet-info"><span>En &#199;ok Tekrar</span><strong><%= en_cok_sayi %> Kez</strong><small><%= en_cok_yemek %></small></div></div>
    <div class="ozet-card"><div class="ozet-icon bg4"><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg></div><div class="ozet-info"><span>Tek Kullan&#305;m</span><strong><%= tek_kullanim %></strong><small>Yemek Sadece 1 Kez</small></div></div>
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

  <div class="kategori-section active" id="sec_tumu"><div class="kategori-card"><div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg>T&#252;m Yemekler</h3><span class="badge"><%= dictTumu.Count %> &#199;e&#351;it</span></div><div class="kategori-card-body"><% RenderKategori dictTumu, en_cok_sayi %></div></div></div>
  <div class="kategori-section" id="sec_ogle_corba"><div class="kategori-card"><div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>&#214;&#287;le &#199;orbas&#305;</h3><span class="badge"><%= dictOgleCorba.Count %> &#199;e&#351;it</span></div><div class="kategori-card-body"><% RenderKategori dictOgleCorba, 0 %></div></div></div>
  <div class="kategori-section" id="sec_ogle_ana"><div class="kategori-card"><div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>&#214;&#287;le Ana Yemek</h3><span class="badge"><%= dictOgleAna.Count %> &#199;e&#351;it</span></div><div class="kategori-card-body"><% RenderKategori dictOgleAna, 0 %></div></div></div>
  <div class="kategori-section" id="sec_ogle_yan"><div class="kategori-card"><div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>&#214;&#287;le Yan &#220;r&#252;n</h3><span class="badge"><%= dictOgleYan.Count %> &#199;e&#351;it</span></div><div class="kategori-card-body"><% RenderKategori dictOgleYan, 0 %></div></div></div>
  <div class="kategori-section" id="sec_ogle_tatli"><div class="kategori-card"><div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>&#214;&#287;le Tatl&#305;/Meyve</h3><span class="badge"><%= dictOgleTatli.Count %> &#199;e&#351;it</span></div><div class="kategori-card-body"><% RenderKategori dictOgleTatli, 0 %></div></div></div>
  <div class="kategori-section" id="sec_aksam_corba"><div class="kategori-card"><div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M20 8.69V4h-4.69L12 .69 8.69 4H4v4.69L.69 12 4 15.31V20h4.69L12 23.31 15.31 20H20v-4.69L23.31 12 20 8.69zM12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6 6 2.69 6 6-2.69 6-6 6zm0-10c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4-1.79-4-4-4z"/></svg>Ak&#351;am &#199;orbas&#305;</h3><span class="badge"><%= dictAksamCorba.Count %> &#199;e&#351;it</span></div><div class="kategori-card-body"><% RenderKategori dictAksamCorba, 0 %></div></div></div>
  <div class="kategori-section" id="sec_aksam_ana"><div class="kategori-card"><div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M20 8.69V4h-4.69L12 .69 8.69 4H4v4.69L.69 12 4 15.31V20h4.69L12 23.31 15.31 20H20v-4.69L23.31 12 20 8.69zM12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6 6 2.69 6 6-2.69 6-6 6zm0-10c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4-1.79-4-4-4z"/></svg>Ak&#351;am Ana Yemek</h3><span class="badge"><%= dictAksamAna.Count %> &#199;e&#351;it</span></div><div class="kategori-card-body"><% RenderKategori dictAksamAna, 0 %></div></div></div>
  <div class="kategori-section" id="sec_aksam_yan"><div class="kategori-card"><div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M20 8.69V4h-4.69L12 .69 8.69 4H4v4.69L.69 12 4 15.31V20h4.69L12 23.31 15.31 20H20v-4.69L23.31 12 20 8.69zM12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6 6 2.69 6 6-2.69 6-6 6zm0-10c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4-1.79-4-4-4z"/></svg>Ak&#351;am Yan &#220;r&#252;n</h3><span class="badge"><%= dictAksamYan.Count %> &#199;e&#351;it</span></div><div class="kategori-card-body"><% RenderKategori dictAksamYan, 0 %></div></div></div>
  <div class="kategori-section" id="sec_aksam_tatli"><div class="kategori-card"><div class="kategori-card-header"><h3><svg viewBox="0 0 24 24"><path d="M20 8.69V4h-4.69L12 .69 8.69 4H4v4.69L.69 12 4 15.31V20h4.69L12 23.31 15.31 20H20v-4.69L23.31 12 20 8.69zM12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6 6 2.69 6 6-2.69 6-6 6zm0-10c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4-1.79-4-4-4z"/></svg>Ak&#351;am Tatl&#305;/Meyve</h3><span class="badge"><%= dictAksamTatli.Count %> &#199;e&#351;it</span></div><div class="kategori-card-body"><% RenderKategori dictAksamTatli, 0 %></div></div></div>
</div>

<div class="modal-overlay" id="detayModal" onclick="if(event.target===this) modalKapat();">
  <div class="modal-box">
    <div class="modal-header"><h2><svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg><span id="modalBaslik">Yemek Detay</span></h2><button class="modal-close" onclick="modalKapat();">&times;</button></div>
    <div class="modal-body" id="modalIcerik"></div>
  </div>
</div>

<script>
var tumVeri = <%= jsonArama %>;
var ogunLabels = {'ogle_corba':'\u00D6\u011Fle \u00C7orba','ogle_ana':'\u00D6\u011Fle Ana Yemek','ogle_yan':'\u00D6\u011Fle Yan \u00DCr\u00FCn','ogle_tatli':'\u00D6\u011Fle Tatl\u0131','aksam_corba':'Ak\u015Fam \u00C7orba','aksam_ana':'Ak\u015Fam Ana Yemek','aksam_yan':'Ak\u015Fam Yan \u00DCr\u00FCn','aksam_tatli':'Ak\u015Fam Tatl\u0131'};
var kategoriLabels = {'ogle_corba':'\u00C7orba','ogle_ana':'Ana Yemek','ogle_yan':'Yan \u00DCr\u00FCn','ogle_tatli':'Tatl\u0131/Meyve','aksam_corba':'\u00C7orba','aksam_ana':'Ana Yemek','aksam_yan':'Yan \u00DCr\u00FCn','aksam_tatli':'Tatl\u0131/Meyve'};
var ogunKeys = ['ogle_corba','ogle_ana','ogle_yan','ogle_tatli','aksam_corba','aksam_ana','aksam_yan','aksam_tatli'];
function decodeHtml(html){var d=document.createElement('div');d.innerHTML=html;return d.textContent||d.innerText||'';}
function detayGoster(yr){var ya=decodeHtml(yr),al=ya.toLowerCase(),bl=[],od={},tb=0,gs=0,gd={};for(var o=0;o<ogunKeys.length;o++)od[ogunKeys[o]]=0;for(var i=0;i<tumVeri.length;i++){var r=tumVeri[i],gb=false,go=[];for(var k=0;k<ogunKeys.length;k++){var key=ogunKeys[k],val=r[key];if(val&&val.toLowerCase()===al){gb=true;tb++;od[key]++;go.push({ogun:key});var gt=decodeHtml(r.gun);if(gd[gt])gd[gt]++;else gd[gt]=1;}}if(gb){gs++;bl.push({tarih:r.tarih,gun:r.gun,ogunler:go});}}document.getElementById('modalBaslik').innerHTML='&quot;'+ya+'&quot; Detayl\u0131 \u0130statistik';if(bl.length===0){document.getElementById('modalIcerik').innerHTML='<div class="empty-msg">Detay bulunamad\u0131.</div>';document.getElementById('detayModal').classList.add('show');document.body.style.overflow='hidden';return;}var eco='',ecs=0;for(var m=0;m<ogunKeys.length;m++){if(od[ogunKeys[m]]>ecs){ecs=od[ogunKeys[m]];eco=ogunLabels[ogunKeys[m]];}}var ecg='',ecgs=0;for(var g in gd){if(gd[g]>ecgs){ecgs=gd[g];ecg=g;}}var oc=0,ac=0;for(var a=0;a<ogunKeys.length;a++){if(ogunKeys[a].indexOf('ogle')>-1)oc+=od[ogunKeys[a]];else ac+=od[ogunKeys[a]];}var h='<div class="modal-ozet">';h+='<div class="modal-ozet-card"><div class="mo-sayi">'+tb+'</div><div class="mo-baslik">Toplam</div></div>';h+='<div class="modal-ozet-card"><div class="mo-sayi">'+gs+'</div><div class="mo-baslik">Farkl\u0131 G\u00FCn</div></div>';if(oc>0)h+='<div class="modal-ozet-card"><div class="mo-sayi">'+oc+'</div><div class="mo-baslik">\u00D6\u011Fle</div></div>';if(ac>0)h+='<div class="modal-ozet-card"><div class="mo-sayi">'+ac+'</div><div class="mo-baslik">Ak\u015Fam</div></div>';h+='<div class="modal-ozet-card vurgulu"><div class="mo-sayi">'+eco+'</div><div class="mo-baslik">En \u00C7ok \u00D6\u011F\u00FCn</div></div>';if(ecg)h+='<div class="modal-ozet-card vurgulu"><div class="mo-sayi">'+ecg+'</div><div class="mo-baslik">En \u00C7ok G\u00FCn</div></div>';h+='</div>';var odh='';for(var n=0;n<ogunKeys.length;n++){if(od[ogunKeys[n]]>0){var ot=ogunKeys[n].indexOf('ogle')>-1?'ogle':'aksam';odh+='<span class="ogun-badge '+ot+'">'+ogunLabels[ogunKeys[n]]+': '+od[ogunKeys[n]]+'</span> ';}}if(odh){h+='<div class="modal-ogun-bar"><span class="modal-ogun-label">\u00D6\u011F\u00FCn Da\u011F\u0131l\u0131m\u0131</span>'+odh+'</div>';}h+='<div class="modal-tablo-wrap"><table class="modal-tablo"><thead><tr><th>#</th><th>Tarih</th><th>G\u00FCn</th><th>\u00D6\u011F\u00FCn</th><th>Kategori</th></tr></thead><tbody>';var s=0;for(var p=0;p<bl.length;p++){var b=bl[p];for(var q=0;q<b.ogunler.length;q++){s++;var ot2=b.ogunler[q].ogun.indexOf('ogle')>-1?'ogle':'aksam';var oa=ot2==='ogle'?'\u00D6\u011Fle':'Ak\u015Fam';h+='<tr><td>'+s+'</td><td><strong>'+b.tarih+'</strong></td><td>'+b.gun+'</td><td><span class="ogun-badge '+ot2+'">'+oa+'</span></td><td><span class="kategori-badge">'+kategoriLabels[b.ogunler[q].ogun]+'</span></td></tr>';}}h+='</tbody></table></div>';document.getElementById('modalIcerik').innerHTML=h;document.getElementById('detayModal').classList.add('show');document.body.style.overflow='hidden';}
function modalKapat(){document.getElementById('detayModal').classList.remove('show');document.body.style.overflow='';}
document.addEventListener('keydown',function(e){if(e.keyCode===27)modalKapat();});
function showTab(tab,btn){var s=document.querySelectorAll('.kategori-section');for(var i=0;i<s.length;i++)s[i].classList.remove('active');document.getElementById('sec_'+tab).classList.add('active');var b=document.querySelectorAll('.filter-btn');for(var j=0;j<b.length;j++)b[j].classList.remove('active');btn.classList.add('active');}
</script>
</body>
</html>
<%
rsYemek.Close: Set rsYemek = Nothing
Set dictOgleCorba = Nothing: Set dictOgleAna = Nothing: Set dictOgleYan = Nothing: Set dictOgleTatli = Nothing
Set dictAksamCorba = Nothing: Set dictAksamAna = Nothing: Set dictAksamYan = Nothing: Set dictAksamTatli = Nothing
Set dictTumu = Nothing
%>
