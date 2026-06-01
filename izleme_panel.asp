<!-- #include file="database/Connection.asp" -->
<%
If Session("yemek_izleme_giris") <> "OK" Then
    Response.Redirect "izleme_giris.asp"
End If

Dim izleme_kullanici
izleme_kullanici = Session("yemek_izleme_kullanici")
Session.Timeout = 120

If Request.QueryString("islem") = "cikis" Then
    Session.Abandon
    Response.Redirect "izleme_giris.asp"
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

Function GetDaysInMonth(m, y)
    Select Case m
        Case 1, 3, 5, 7, 8, 10, 12: GetDaysInMonth = 31
        Case 4, 6, 9, 11: GetDaysInMonth = 30
        Case 2
            If (y Mod 4 = 0 And y Mod 100 <> 0) Or (y Mod 400 = 0) Then
                GetDaysInMonth = 29
            Else
                GetDaysInMonth = 28
            End If
    End Select
End Function

' ============================================
' NORMAL ISTATISTIKLER
' ============================================
Dim rsToplam, toplamAylik, toplamKayit, dictYillar
Set dictYillar = CreateObject("Scripting.Dictionary")
toplamAylik = 0

Set rsToplam = ConnYemek.Execute("SELECT DISTINCT Year(tarih) as yil, Month(tarih) as ay FROM yemek_listesi ORDER BY Year(tarih) DESC, Month(tarih) DESC")
If Not rsToplam.EOF Then
    Do While Not rsToplam.EOF
        Dim yilStr
        yilStr = CStr(rsToplam("yil"))
        If dictYillar.Exists(yilStr) Then
            dictYillar(yilStr) = dictYillar(yilStr) + 1
        Else
            dictYillar.Add yilStr, 1
        End If
        toplamAylik = toplamAylik + 1
        rsToplam.MoveNext
    Loop
End If
rsToplam.Close
Set rsToplam = Nothing

Dim rsToplamKayit
Set rsToplamKayit = ConnYemek.Execute("SELECT COUNT(*) as toplam FROM yemek_listesi")
toplamKayit = 0
If Not rsToplamKayit.EOF Then toplamKayit = rsToplamKayit("toplam")
rsToplamKayit.Close
Set rsToplamKayit = Nothing

' ============================================
' DIYET ISTATISTIKLER
' ============================================
Dim toplamAylikDiyet, toplamKayitDiyet, dictYillarDiyet
Set dictYillarDiyet = CreateObject("Scripting.Dictionary")
toplamAylikDiyet = 0
toplamKayitDiyet = 0

On Error Resume Next
Dim rsToplamDiyet
Set rsToplamDiyet = ConnYemek.Execute("SELECT DISTINCT Year(tarih) as yil, Month(tarih) as ay FROM diyet_yemek_listesi ORDER BY Year(tarih) DESC, Month(tarih) DESC")
If Err.Number = 0 Then
    If Not rsToplamDiyet.EOF Then
        Do While Not rsToplamDiyet.EOF
            Dim yilStrD
            yilStrD = CStr(rsToplamDiyet("yil"))
            If dictYillarDiyet.Exists(yilStrD) Then
                dictYillarDiyet(yilStrD) = dictYillarDiyet(yilStrD) + 1
            Else
                dictYillarDiyet.Add yilStrD, 1
            End If
            toplamAylikDiyet = toplamAylikDiyet + 1
            rsToplamDiyet.MoveNext
        Loop
    End If
    rsToplamDiyet.Close
    Set rsToplamDiyet = Nothing

    Dim rsToplamKayitDiyet
    Set rsToplamKayitDiyet = ConnYemek.Execute("SELECT COUNT(*) as toplam FROM diyet_yemek_listesi")
    If Err.Number = 0 And Not rsToplamKayitDiyet.EOF Then toplamKayitDiyet = rsToplamKayitDiyet("toplam")
    If Not rsToplamKayitDiyet Is Nothing Then rsToplamKayitDiyet.Close: Set rsToplamKayitDiyet = Nothing
End If
Err.Clear
On Error GoTo 0

' ============================================
' BU AY DURUMU
' ============================================
Dim bugun_yil, bugun_ay
bugun_yil = Year(Now())
bugun_ay = Month(Now())

Dim rsBuAy, buAyKayit, buAyDurum, buAyGun
buAyGun = GetDaysInMonth(bugun_ay, bugun_yil)
Set rsBuAy = ConnYemek.Execute("SELECT COUNT(*) as toplam FROM yemek_listesi WHERE Year(tarih) = " & bugun_yil & " AND Month(tarih) = " & bugun_ay)
buAyKayit = 0
If Not rsBuAy.EOF Then buAyKayit = rsBuAy("toplam")
rsBuAy.Close
Set rsBuAy = Nothing
If buAyKayit > 0 Then buAyDurum = "girildi" Else buAyDurum = "girilmedi"

Dim buAyKayitDiyet, buAyDurumDiyet
buAyKayitDiyet = 0
On Error Resume Next
Dim rsBuAyD
Set rsBuAyD = ConnYemek.Execute("SELECT COUNT(*) as toplam FROM diyet_yemek_listesi WHERE Year(tarih) = " & bugun_yil & " AND Month(tarih) = " & bugun_ay)
If Err.Number = 0 And Not rsBuAyD.EOF Then buAyKayitDiyet = rsBuAyD("toplam")
If Not rsBuAyD Is Nothing Then rsBuAyD.Close: Set rsBuAyD = Nothing
Err.Clear
On Error GoTo 0
If buAyKayitDiyet > 0 Then buAyDurumDiyet = "girildi" Else buAyDurumDiyet = "girilmedi"

' ============================================
' GELECEK AY DURUMU
' ============================================
Dim gelecekAy, gelecekYil, gelecekAyGun
gelecekAy = bugun_ay + 1
gelecekYil = bugun_yil
If gelecekAy > 12 Then
    gelecekAy = 1
    gelecekYil = bugun_yil + 1
End If
gelecekAyGun = GetDaysInMonth(gelecekAy, gelecekYil)

Dim rsGelecek, gelecekKayit, gelecekDurum
Set rsGelecek = ConnYemek.Execute("SELECT COUNT(*) as toplam FROM yemek_listesi WHERE Year(tarih) = " & gelecekYil & " AND Month(tarih) = " & gelecekAy)
gelecekKayit = 0
If Not rsGelecek.EOF Then gelecekKayit = rsGelecek("toplam")
rsGelecek.Close
Set rsGelecek = Nothing
If gelecekKayit > 0 Then gelecekDurum = "girildi" Else gelecekDurum = "girilmedi"

Dim gelecekKayitDiyet, gelecekDurumDiyet
gelecekKayitDiyet = 0
On Error Resume Next
Dim rsGelecekD
Set rsGelecekD = ConnYemek.Execute("SELECT COUNT(*) as toplam FROM diyet_yemek_listesi WHERE Year(tarih) = " & gelecekYil & " AND Month(tarih) = " & gelecekAy)
If Err.Number = 0 And Not rsGelecekD.EOF Then gelecekKayitDiyet = rsGelecekD("toplam")
If Not rsGelecekD Is Nothing Then rsGelecekD.Close: Set rsGelecekD = Nothing
Err.Clear
On Error GoTo 0
If gelecekKayitDiyet > 0 Then gelecekDurumDiyet = "girildi" Else gelecekDurumDiyet = "girilmedi"

%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Y&#246;netim Paneli - Yemek Sistemi</title>
  <link rel="icon" href="../../images/hastane_portal_logo.png" type="image/png">
  <style>
    :root { --main: #45b8c3; --dark: #2e8b91; --txt: #fff; --diyet: #4caf50; --diyet-dark: #388e3c; }
    * { margin: 0; padding: 0; box-sizing: border-box; -webkit-user-select: none; -moz-user-select: none; -ms-user-select: none; user-select: none; }
    html, body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; height: 100vh; overflow: hidden; background: linear-gradient(135deg, #f0f9fa, #e6f4f1); }

    #header { background: linear-gradient(90deg, var(--main), var(--dark)); padding: 8px 15px; color: var(--txt); display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 10px rgba(0,0,0,0.15); }
    .header-left { display: flex; align-items: center; gap: 10px; }
    .nav-button { background: transparent; color: var(--txt); width: 36px; height: 36px; border: 2px solid var(--txt); border-radius: 50%; text-decoration: none; transition: all 0.3s; display: flex; align-items: center; justify-content: center; cursor: pointer; }
    .nav-button:hover { background: var(--dark); transform: translateY(-2px); }
    .nav-button svg { width: 20px; height: 20px; fill: currentColor; }
    #header h1 { font-size: 18px; margin: 0; text-shadow: 1px 1px 2px rgba(0,0,0,0.2); display: flex; align-items: center; gap: 8px; }
    #header h1 svg { width: 20px; height: 20px; opacity: 0.85; }
    #user-info { display: flex; align-items: center; gap: 8px; font-size: 13px; background: rgba(255,255,255,0.15); padding: 6px 14px; border-radius: 20px; }
    #user-info svg { width: 18px; height: 18px; fill: currentColor; }

    .content { height: calc(100vh - 52px); overflow-y: auto; padding: 20px; }
    .content::-webkit-scrollbar { width: 6px; }
    .content::-webkit-scrollbar-track { background: #f1f1f1; }
    .content::-webkit-scrollbar-thumb { background: var(--main); border-radius: 10px; }

    .mini-stats { display: flex; gap: 12px; margin-bottom: 18px; flex-wrap: wrap; }
    .mini-stat { flex: 1; min-width: 140px; background: #fff; border-radius: 10px; padding: 14px 16px; display: flex; align-items: center; gap: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); border: 1px solid #eef5f5; transition: all 0.3s; }
    .mini-stat:hover { transform: translateY(-2px); box-shadow: 0 4px 14px rgba(0,0,0,0.08); }
    .mini-stat-icon { width: 36px; height: 36px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
    .mini-stat-icon.normal-icon { background: linear-gradient(135deg, rgba(69,184,195,0.12), rgba(46,139,145,0.08)); }
    .mini-stat-icon.normal-icon svg { fill: var(--dark); }
    .mini-stat-icon.diyet-icon { background: linear-gradient(135deg, rgba(76,175,80,0.12), rgba(56,142,60,0.08)); }
    .mini-stat-icon.diyet-icon svg { fill: var(--diyet-dark); }
    .mini-stat-icon svg { width: 18px; height: 18px; }
    .mini-stat-info { display: flex; flex-direction: column; }
    .mini-stat-value { font-size: 20px; font-weight: 700; color: #2c3e50; line-height: 1.2; }
    .mini-stat-label { font-size: 10px; color: #999; text-transform: uppercase; letter-spacing: 0.4px; }

    .ay-durum-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-bottom: 18px; }
    .ay-durum-card { background: #fff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 14px rgba(0,0,0,0.05); transition: all 0.3s; }
    .ay-durum-card:hover { transform: translateY(-3px); box-shadow: 0 8px 22px rgba(0,0,0,0.1); }
    .ay-durum-top { padding: 8px 0; text-align: center; }
    .ay-durum-card.girildi .ay-durum-top { background: linear-gradient(135deg, var(--main), var(--dark)); }
    .ay-durum-card.girilmedi .ay-durum-top { background: linear-gradient(135deg, #78909c, #546e7a); }
    .ay-durum-label { font-size: 10px; color: #fff; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; }
    .ay-durum-body { padding: 18px 16px; text-align: center; }
    .ay-durum-icon { width: 44px; height: 44px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 10px; }
    .ay-durum-icon svg { width: 24px; height: 24px; fill: #fff; }
    .ay-durum-card.girildi .ay-durum-icon { background: linear-gradient(135deg, var(--main), var(--dark)); }
    .ay-durum-card.girilmedi .ay-durum-icon { background: linear-gradient(135deg, #78909c, #546e7a); }
    .ay-durum-title { font-size: 16px; font-weight: 700; color: #2c3e50; margin-bottom: 4px; }
    .ay-durum-title .sep { color: #ccc; font-weight: 300; margin: 0 4px; }
    .ay-durum-gun { font-size: 11px; color: #999; margin-bottom: 10px; }
    .ay-durum-badges { display: flex; gap: 6px; justify-content: center; flex-wrap: wrap; }
    .ay-badge { display: inline-flex; align-items: center; gap: 4px; padding: 4px 12px; border-radius: 16px; font-size: 11px; font-weight: 700; }
    .ay-badge.normal-g { background: rgba(69,184,195,0.12); color: var(--dark); }
    .ay-badge.normal-x { background: rgba(120,144,156,0.12); color: #546e7a; }
    .ay-badge.diyet-g { background: rgba(76,175,80,0.12); color: var(--diyet-dark); }
    .ay-badge.diyet-x { background: rgba(120,144,156,0.08); color: #90a4ae; }
    .ay-badge svg { width: 12px; height: 12px; fill: currentColor; }

    .yil-section { margin-bottom: 18px; }
    .yil-section-header { background: #fff; border-radius: 12px 12px 0 0; padding: 16px 22px; border-bottom: 2px solid var(--main); box-shadow: 0 2px 8px rgba(0,0,0,0.04); }
    .yil-section-header h2 { font-size: 15px; color: #2c3e50; margin: 0; display: flex; align-items: center; gap: 10px; }
    .yil-section-header h2 svg { width: 20px; height: 20px; fill: var(--main); }
    .yil-section-body { background: #fff; border-radius: 0 0 12px 12px; padding: 18px 22px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); }
    .yil-cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 12px; }
    .yil-card { background: linear-gradient(135deg, rgba(69,184,195,0.04), rgba(46,139,145,0.02)); border: 2px solid #e8e8e8; border-radius: 12px; padding: 16px 14px; text-align: center; transition: all 0.3s; cursor: default; }
    .yil-card:hover { border-color: var(--main); transform: translateY(-3px); box-shadow: 0 8px 20px rgba(69,184,195,0.15); }
    .yil-card.diyet-card { background: linear-gradient(135deg, rgba(76,175,80,0.04), rgba(56,142,60,0.02)); }
    .yil-card.diyet-card:hover { border-color: var(--diyet); box-shadow: 0 8px 20px rgba(76,175,80,0.15); }
    .yil-card-year { font-size: 20px; font-weight: 700; color: #2c3e50; margin-bottom: 6px; }
    .yil-card-count { display: inline-block; background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; padding: 3px 12px; border-radius: 12px; font-size: 11px; font-weight: 700; }
    .yil-card.diyet-card .yil-card-count { background: linear-gradient(135deg, var(--diyet), var(--diyet-dark)); }
    .yil-card-bar { margin-top: 8px; height: 4px; background: #f0f0f0; border-radius: 2px; overflow: hidden; }
    .yil-card-bar-fill { height: 100%; background: linear-gradient(90deg, var(--main), var(--dark)); border-radius: 2px; transition: width 0.6s ease; }
    .yil-card.diyet-card .yil-card-bar-fill { background: linear-gradient(90deg, var(--diyet), var(--diyet-dark)); }
    .yil-card-type { font-size: 9px; color: #999; text-transform: uppercase; margin-top: 6px; letter-spacing: 0.5px; }

    .menu-container { background: #fff; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); overflow: hidden; }
    .menu-header { padding: 16px 22px; border-bottom: 2px solid var(--main); }
    .menu-header h2 { font-size: 15px; color: #2c3e50; margin: 0; display: flex; align-items: center; gap: 10px; }
    .menu-header h2 svg { width: 20px; height: 20px; fill: var(--main); }
    .menu-body { padding: 18px 22px; display: flex; gap: 12px; flex-wrap: wrap; }
    .menu-btn { display: flex; align-items: center; gap: 14px; padding: 16px 18px; background: #fff; border: 2px solid #e8e8e8; border-radius: 12px; text-decoration: none; color: #2c3e50; transition: all 0.3s; cursor: pointer; min-width: 240px; }
    .menu-btn:hover { border-color: var(--main); transform: translateY(-3px); box-shadow: 0 8px 22px rgba(69,184,195,0.2); }
    .menu-btn.diyet-btn:hover { border-color: var(--diyet); box-shadow: 0 8px 22px rgba(76,175,80,0.2); }
    .menu-btn-icon { width: 44px; height: 44px; border-radius: 10px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; background: linear-gradient(135deg, var(--main), var(--dark)); }
    .menu-btn.diyet-btn .menu-btn-icon { background: linear-gradient(135deg, var(--diyet), var(--diyet-dark)); }
    .menu-btn-icon svg { width: 22px; height: 22px; fill: #fff; }
    .menu-btn-text strong { display: block; font-size: 14px; margin-bottom: 2px; }
    .menu-btn-text span { font-size: 11px; color: #999; }

    @media (max-width: 1000px) { .mini-stats { flex-wrap: wrap; } .mini-stat { min-width: calc(50% - 6px); } }
    @media (max-width: 768px) { #header { flex-wrap: wrap; gap: 8px; padding: 10px; } #header h1 { font-size: 15px; } .mini-stats { flex-direction: column; } .ay-durum-row { grid-template-columns: 1fr; } .yil-cards { grid-template-columns: 1fr 1fr; } .menu-body { flex-direction: column; } }
  </style>
</head>
<body>
  <div id="header">
    <div class="header-left">
      <a href="izleme_liste.asp" class="nav-button" title="Yemek Listeleri">
        <svg viewBox="0 0 24 24"><path d="M3 13h2v-2H3v2zm0 4h2v-2H3v2zm0-8h2V7H3v2zm4 4h14v-2H7v2zm0 4h14v-2H7v2zM7 7v2h14V7H7z"/></svg>
      </a>
    </div>
    <h1>
      <svg viewBox="0 0 24 24" fill="currentColor"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>
      Yemek Sistemi - Y&#246;netim Paneli
    </h1>
    <div style="display:flex;align-items:center;gap:10px;">
      <div id="user-info">
        <svg viewBox="0 0 24 24"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>
        <strong><%= izleme_kullanici %></strong>
      </div>
      <a href="izleme_panel.asp?islem=cikis" class="nav-button" title="&#199;&#305;k&#305;&#351; Yap" onclick="return confirm('\u00c7\u0131k\u0131\u015f Yapmak \u0130stedi\u011finize Emin Misiniz?')">
        <svg viewBox="0 0 24 24"><path d="M17 7l-1.41 1.41L18.17 11H8v2h10.17l-2.58 2.58L17 17l5-5zM4 5h8V3H4c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h8v-2H4V5z"/></svg>
      </a>
    </div>
  </div>

  <div class="content">

    <div class="mini-stats">
      <div class="mini-stat">
        <div class="mini-stat-icon normal-icon"><svg viewBox="0 0 24 24"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10z"/></svg></div>
        <div class="mini-stat-info"><div class="mini-stat-value"><%= toplamAylik %></div><div class="mini-stat-label">Normal Ayl&#305;k Men&#252;</div></div>
      </div>
      <div class="mini-stat">
        <div class="mini-stat-icon normal-icon"><svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg></div>
        <div class="mini-stat-info"><div class="mini-stat-value"><%= toplamKayit %></div><div class="mini-stat-label">Normal G&#252;nl&#252;k Kay&#305;t</div></div>
      </div>
      <div class="mini-stat">
        <div class="mini-stat-icon diyet-icon"><svg viewBox="0 0 24 24"><path d="M17.21 9l-4.38-6.56c-.19-.28-.51-.42-.83-.42-.32 0-.64.14-.83.43L6.79 9C6.3 9.71 6 10.57 6 11.5 6 14.53 8.47 17 11.5 17h1c3.03 0 5.5-2.47 5.5-5.5 0-.93-.3-1.79-.79-2.5z"/></svg></div>
        <div class="mini-stat-info"><div class="mini-stat-value"><%= toplamAylikDiyet %></div><div class="mini-stat-label">Diyet Ayl&#305;k Men&#252;</div></div>
      </div>
      <div class="mini-stat">
        <div class="mini-stat-icon diyet-icon"><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg></div>
        <div class="mini-stat-info"><div class="mini-stat-value"><%= toplamKayitDiyet %></div><div class="mini-stat-label">Diyet G&#252;nl&#252;k Kay&#305;t</div></div>
      </div>
    </div>

    <div class="ay-durum-row">
      <div class="ay-durum-card <%= buAyDurum %>">
        <div class="ay-durum-top"><div class="ay-durum-label">Bu Ay</div></div>
        <div class="ay-durum-body">
          <div class="ay-durum-icon"><% If buAyDurum = "girildi" Then %><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><% Else %><svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg><% End If %></div>
          <div class="ay-durum-title"><%= GetMonthName(bugun_ay) %> <span class="sep">|</span> <%= bugun_yil %></div>
          <div class="ay-durum-gun">(<%= buAyGun %> G&#252;n)</div>
          <div class="ay-durum-badges">
            <span class="ay-badge <% If buAyDurum = "girildi" Then %>normal-g<% Else %>normal-x<% End If %>">
              <% If buAyDurum = "girildi" Then %><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><% Else %><svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg><% End If %>
              Normal
            </span>
            <span class="ay-badge <% If buAyDurumDiyet = "girildi" Then %>diyet-g<% Else %>diyet-x<% End If %>">
              <% If buAyDurumDiyet = "girildi" Then %><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><% Else %><svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg><% End If %>
              Diyet
            </span>
          </div>
        </div>
      </div>
      <div class="ay-durum-card <%= gelecekDurum %>">
        <div class="ay-durum-top"><div class="ay-durum-label">Gelecek Ay</div></div>
        <div class="ay-durum-body">
          <div class="ay-durum-icon"><% If gelecekDurum = "girildi" Then %><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><% Else %><svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg><% End If %></div>
          <div class="ay-durum-title"><%= GetMonthName(gelecekAy) %> <span class="sep">|</span> <%= gelecekYil %></div>
          <div class="ay-durum-gun">(<%= gelecekAyGun %> G&#252;n)</div>
          <div class="ay-durum-badges">
            <span class="ay-badge <% If gelecekDurum = "girildi" Then %>normal-g<% Else %>normal-x<% End If %>">
              <% If gelecekDurum = "girildi" Then %><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><% Else %><svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg><% End If %>
              Normal
            </span>
            <span class="ay-badge <% If gelecekDurumDiyet = "girildi" Then %>diyet-g<% Else %>diyet-x<% End If %>">
              <% If gelecekDurumDiyet = "girildi" Then %><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><% Else %><svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg><% End If %>
              Diyet
            </span>
          </div>
        </div>
      </div>
    </div>

    <div class="yil-section">
      <div class="yil-section-header">
        <h2><svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg>Y&#305;l Bazl&#305; &#304;statistikler</h2>
      </div>
      <div class="yil-section-body">
        <div class="yil-cards">
          <% Dim yKey, barPct
          For Each yKey In dictYillar.Keys
            barPct = Round((dictYillar(yKey) / 12) * 100)
          %>
          <div class="yil-card">
            <div class="yil-card-year"><%= yKey %></div>
            <div class="yil-card-count"><%= dictYillar(yKey) %> Ay</div>
            <div class="yil-card-bar"><div class="yil-card-bar-fill" style="width:<%= barPct %>%;"></div></div>
            <div class="yil-card-type">Normal</div>
          </div>
          <% Next %>
          <% Dim yKeyD, barPctD
          For Each yKeyD In dictYillarDiyet.Keys
            barPctD = Round((dictYillarDiyet(yKeyD) / 12) * 100)
          %>
          <div class="yil-card diyet-card">
            <div class="yil-card-year"><%= yKeyD %></div>
            <div class="yil-card-count"><%= dictYillarDiyet(yKeyD) %> Ay</div>
            <div class="yil-card-bar"><div class="yil-card-bar-fill" style="width:<%= barPctD %>%;"></div></div>
            <div class="yil-card-type">Diyet</div>
          </div>
          <% Next %>
        </div>
      </div>
    </div>

    <div class="menu-container">
      <div class="menu-header">
        <h2><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>Men&#252; G&#246;r&#252;nt&#252;le</h2>
      </div>
      <div class="menu-body">
        <a href="izleme_liste.asp" class="menu-btn">
          <div class="menu-btn-icon"><svg viewBox="0 0 24 24"><path d="M3 13h2v-2H3v2zm0 4h2v-2H3v2zm0-8h2V7H3v2zm4 4h14v-2H7v2zm0 4h14v-2H7v2zM7 7v2h14V7H7z"/></svg></div>
          <div class="menu-btn-text"><strong>Yemek Listeleri</strong><span>Normal men&#252;leri g&#246;r&#252;nt&#252;le</span></div>
        </a>
        <a href="izleme_liste.asp?sekme=diyet" class="menu-btn diyet-btn">
          <div class="menu-btn-icon"><svg viewBox="0 0 24 24"><path d="M17.21 9l-4.38-6.56c-.19-.28-.51-.42-.83-.42-.32 0-.64.14-.83.43L6.79 9C6.3 9.71 6 10.57 6 11.5 6 14.53 8.47 17 11.5 17h1c3.03 0 5.5-2.47 5.5-5.5 0-.93-.3-1.79-.79-2.5z"/></svg></div>
          <div class="menu-btn-text"><strong>Diyet Yemek Listeleri</strong><span>Diyet men&#252;lerini g&#246;r&#252;nt&#252;le</span></div>
        </a>
      </div>
    </div>

  </div>

  <script>
    var inactivityTimeout;
    function resetInactivityTimer() {
      clearTimeout(inactivityTimeout);
      inactivityTimeout = setTimeout(function() { window.location.href = 'izleme_panel.asp?islem=cikis'; }, 300000);
    }
    document.addEventListener('mousemove', resetInactivityTimer);
    document.addEventListener('keypress', resetInactivityTimer);
    document.addEventListener('click', resetInactivityTimer);
    document.addEventListener('scroll', resetInactivityTimer);
    document.addEventListener('touchstart', resetInactivityTimer);
    resetInactivityTimer();
    document.addEventListener('visibilitychange', function() { if (!document.hidden) resetInactivityTimer(); });
  </script>
</body>
</html>
<%
Set dictYillar = Nothing
Set dictYillarDiyet = Nothing
%>
