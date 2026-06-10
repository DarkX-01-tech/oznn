<!-- #include file="database/Connection.asp" -->
<%
If Session("yemek_admin_giris") <> "OK" Then
    Response.Redirect "giris.asp"
End If

Dim admin_kullanici
admin_kullanici = Session("yemek_admin_kullanici")
Session.Timeout = 2

If Request.QueryString("islem") = "cikis" Then
    Session.Abandon
    Response.Redirect "giris.asp"
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
' GENEL ISTATISTIKLER - NORMAL
' ============================================
Dim rsToplam, toplamAylik, dictYillar
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

' ============================================
' GENEL ISTATISTIKLER - DIYET
' ============================================
Dim rsToplamDiyet, toplamAylikDiyet, dictYillarDiyet
Set dictYillarDiyet = CreateObject("Scripting.Dictionary")
toplamAylikDiyet = 0

On Error Resume Next
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

If buAyKayit > 0 Then
    buAyDurum = "girildi"
Else
    buAyDurum = "girilmedi"
End If

' BU AY DIYET DURUMU
Dim rsBuAyDiyet, buAyKayitDiyet, buAyDurumDiyet
buAyKayitDiyet = 0
On Error Resume Next
Set rsBuAyDiyet = ConnYemek.Execute("SELECT COUNT(*) as toplam FROM diyet_yemek_listesi WHERE Year(tarih) = " & bugun_yil & " AND Month(tarih) = " & bugun_ay)
If Err.Number = 0 And Not rsBuAyDiyet.EOF Then buAyKayitDiyet = rsBuAyDiyet("toplam")
If Not rsBuAyDiyet Is Nothing Then rsBuAyDiyet.Close: Set rsBuAyDiyet = Nothing
Err.Clear
On Error GoTo 0

If buAyKayitDiyet > 0 Then
    buAyDurumDiyet = "girildi"
Else
    buAyDurumDiyet = "girilmedi"
End If

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

If gelecekKayit > 0 Then
    gelecekDurum = "girildi"
Else
    gelecekDurum = "girilmedi"
End If

' GELECEK AY DIYET DURUMU
Dim rsGelecekDiyet, gelecekKayitDiyet, gelecekDurumDiyet
gelecekKayitDiyet = 0
On Error Resume Next
Set rsGelecekDiyet = ConnYemek.Execute("SELECT COUNT(*) as toplam FROM diyet_yemek_listesi WHERE Year(tarih) = " & gelecekYil & " AND Month(tarih) = " & gelecekAy)
If Err.Number = 0 And Not rsGelecekDiyet.EOF Then gelecekKayitDiyet = rsGelecekDiyet("toplam")
If Not rsGelecekDiyet Is Nothing Then rsGelecekDiyet.Close: Set rsGelecekDiyet = Nothing
Err.Clear
On Error GoTo 0

If gelecekKayitDiyet > 0 Then
    gelecekDurumDiyet = "girildi"
Else
    gelecekDurumDiyet = "girilmedi"
End If

' ============================================
' GECMIS DONEM LISTESI - NORMAL
' ============================================
Dim rsGecmis, sqlGecmis
sqlGecmis = "SELECT DISTINCT Year(tarih) as yil, Month(tarih) as ay FROM yemek_listesi " & _
            "WHERE (Year(tarih) < " & bugun_yil & ") " & _
            "OR (Year(tarih) = " & bugun_yil & " AND Month(tarih) < " & bugun_ay & ") " & _
            "ORDER BY Year(tarih) DESC, Month(tarih) DESC"
Set rsGecmis = ConnYemek.Execute(sqlGecmis)

' ============================================
' GECMIS DONEM LISTESI - DIYET
' ============================================
Dim rsGecmisDiyet, sqlGecmisDiyet, diyetGecmisVar
diyetGecmisVar = False
On Error Resume Next
sqlGecmisDiyet = "SELECT DISTINCT Year(tarih) as yil, Month(tarih) as ay FROM diyet_yemek_listesi " & _
            "WHERE (Year(tarih) < " & bugun_yil & ") " & _
            "OR (Year(tarih) = " & bugun_yil & " AND Month(tarih) < " & bugun_ay & ") " & _
            "ORDER BY Year(tarih) DESC, Month(tarih) DESC"
Set rsGecmisDiyet = ConnYemek.Execute(sqlGecmisDiyet)
If Err.Number = 0 Then diyetGecmisVar = True
Err.Clear
On Error GoTo 0
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Y&#246;netici Paneli - Yemek Sistemi</title>
  <link rel="icon" href="../../images/hastane_portal_logo.png" type="image/png">
  <style>
    :root { --main: #45b8c3; --dark: #2e8b91; --txt: #fff; --diyet: #4caf50; --diyet-dark: #388e3c; }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    html, body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      color: #333; height: 100vh; overflow: hidden;
      background: linear-gradient(135deg, #f0f9fa, #e6f4f1);
    }

    #header {
      background: linear-gradient(90deg, var(--main), var(--dark));
      padding: 8px 15px; color: var(--txt);
      display: flex; align-items: center; justify-content: space-between;
      box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    }
    #nav-buttons { display: flex; gap: 8px; }
    .nav-button {
      background: transparent; color: var(--txt);
      width: 36px; height: 36px; border: 2px solid var(--txt);
      border-radius: 50%; text-decoration: none; transition: all 0.3s;
      display: flex; align-items: center; justify-content: center; cursor: pointer;
    }
    .nav-button:hover { background: var(--dark); transform: translateY(-2px); }
    .nav-button.active { background: rgba(255,255,255,0.2); }
    .nav-button svg { width: 20px; height: 20px; fill: currentColor; }
    #header h1 {
      font-size: 18px; margin: 0; text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
      display: flex; align-items: center; gap: 8px;
    }
    #header h1 svg { width: 20px; height: 20px; opacity: 0.85; }
    #user-info {
      display: flex; align-items: center; gap: 8px; font-size: 13px;
      background: rgba(255,255,255,0.15); padding: 6px 14px; border-radius: 20px;
    }
    #user-info svg { width: 18px; height: 18px; fill: currentColor; }

    .content { height: calc(100vh - 52px); overflow-y: auto; padding: 20px; }

    .stats-row { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; margin-bottom: 22px; }

    .stat-genel {
      background: #fff; border-radius: 12px;
      box-shadow: 0 4px 14px rgba(0,0,0,0.05); overflow: hidden;
    }
    .stat-genel-header {
      background: linear-gradient(135deg, var(--main), var(--dark));
      padding: 16px 18px; color: #fff;
      display: flex; align-items: center; gap: 12px;
    }
    .stat-genel-header.diyet-header { background: linear-gradient(135deg, var(--diyet), var(--diyet-dark)); }
    .stat-genel-header .sg-icon {
      width: 40px; height: 40px; border-radius: 10px;
      background: rgba(255,255,255,0.2);
      display: flex; align-items: center; justify-content: center; flex-shrink: 0;
    }
    .stat-genel-header .sg-icon svg { width: 22px; height: 22px; fill: #fff; }
    .stat-genel-header .sg-info { color: #fff; }
    .stat-genel-header .sg-info span { display: block; font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px; opacity: 0.85; }
    .stat-genel-header .sg-info strong { font-size: 20px; }
    .stat-genel-header .sg-info strong small { font-size: 12px; font-weight: 400; opacity: 0.85; }
    .yil-liste { padding: 12px; display: flex; flex-direction: column; gap: 6px; max-height: 120px; overflow-y: auto; }
    .yil-liste::-webkit-scrollbar { width: 4px; }
    .yil-liste::-webkit-scrollbar-thumb { background: var(--main); border-radius: 10px; }
    .yil-item {
      display: flex; align-items: center; justify-content: space-between;
      padding: 9px 12px; background: linear-gradient(135deg, rgba(69,184,195,0.06), rgba(46,139,145,0.03));
      border-radius: 8px; border: 1px solid #eef5f4; transition: all 0.2s;
    }
    .yil-item:hover { background: linear-gradient(135deg, rgba(69,184,195,0.12), rgba(46,139,145,0.06)); border-color: var(--main); }
    .yil-item.diyet-item { background: linear-gradient(135deg, rgba(76,175,80,0.06), rgba(56,142,60,0.03)); border-color: #e8f5e9; }
    .yil-item.diyet-item:hover { background: linear-gradient(135deg, rgba(76,175,80,0.12), rgba(56,142,60,0.06)); border-color: var(--diyet); }
    .yil-item-left { display: flex; align-items: center; gap: 8px; }
    .yil-item-left svg { width: 14px; height: 14px; fill: var(--main); }
    .yil-item.diyet-item .yil-item-left svg { fill: var(--diyet); }
    .yil-item-left span { font-size: 13px; font-weight: 600; color: #2c3e50; }
    .yil-item-right {
      background: linear-gradient(135deg, var(--main), var(--dark));
      color: #fff; padding: 3px 10px; border-radius: 10px;
      font-size: 11px; font-weight: 700;
    }
    .yil-item.diyet-item .yil-item-right { background: linear-gradient(135deg, var(--diyet), var(--diyet-dark)); }

    .stat-ay {
      background: #fff; border-radius: 12px;
      box-shadow: 0 4px 14px rgba(0,0,0,0.05);
      overflow: hidden; transition: all 0.3s;
    }
    .stat-ay:hover { transform: translateY(-3px); box-shadow: 0 8px 22px rgba(0,0,0,0.1); }
    .stat-ay-top { padding: 8px 0; text-align: center; }
    .stat-ay.girildi .stat-ay-top { background: linear-gradient(135deg, var(--main), var(--dark)); }
    .stat-ay.girilmedi .stat-ay-top { background: linear-gradient(135deg, #78909c, #546e7a); }
    .stat-ay-label { font-size: 10px; color: #fff; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; }
    .stat-ay-body { padding: 18px 16px; text-align: center; }
    .stat-ay-icon {
      width: 44px; height: 44px; border-radius: 50%;
      display: flex; align-items: center; justify-content: center;
      margin: 0 auto 10px;
    }
    .stat-ay-icon svg { width: 24px; height: 24px; fill: #fff; }
    .stat-ay.girildi .stat-ay-icon { background: linear-gradient(135deg, var(--main), var(--dark)); }
    .stat-ay.girilmedi .stat-ay-icon { background: linear-gradient(135deg, #78909c, #546e7a); }
    .stat-ay-title { font-size: 16px; font-weight: 700; color: #2c3e50; margin-bottom: 4px; }
    .stat-ay-title .sep { color: #ccc; font-weight: 300; margin: 0 4px; }
    .stat-ay-gun { font-size: 11px; color: #999; margin-bottom: 10px; }
    .stat-ay-badges { display: flex; gap: 6px; justify-content: center; flex-wrap: wrap; }
    .stat-ay-badge { display: inline-flex; align-items: center; gap: 4px; padding: 4px 12px; border-radius: 16px; font-size: 11px; font-weight: 700; letter-spacing: 0.3px; }
    .stat-ay-badge.normal-badge.girildi { background: rgba(69,184,195,0.12); color: var(--dark); }
    .stat-ay-badge.normal-badge.girilmedi { background: rgba(120,144,156,0.12); color: #546e7a; }
    .stat-ay-badge.diyet-badge.girildi { background: rgba(76,175,80,0.12); color: var(--diyet-dark); }
    .stat-ay-badge.diyet-badge.girilmedi { background: rgba(120,144,156,0.08); color: #90a4ae; }
    .stat-ay-badge svg { width: 12px; height: 12px; fill: currentColor; }

    .menu-container {
      display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
      gap: 15px; margin-bottom: 22px;
    }
    .menu-card {
      background: #fff; border-radius: 12px; padding: 22px;
      box-shadow: 0 4px 14px rgba(0,0,0,0.05);
      text-align: center; text-decoration: none; color: #333;
      transition: all 0.35s cubic-bezier(0.22,1,0.36,1);
      border: 2px solid transparent; position: relative; overflow: hidden;
    }
    .menu-card::after {
      content: ''; position: absolute; bottom: 0; left: 0; right: 0;
      height: 3px; background: linear-gradient(90deg, var(--main), var(--dark));
      transform: scaleX(0); transition: transform 0.3s;
    }
    .menu-card:hover {
      transform: translateY(-5px);
      box-shadow: 0 10px 25px rgba(69,184,195,0.18);
      border-color: var(--main);
    }
    .menu-card:hover::after { transform: scaleX(1); }
    .menu-card-icon {
      width: 52px; height: 52px; margin: 0 auto 14px;
      border-radius: 14px; display: flex; align-items: center; justify-content: center;
      background: linear-gradient(135deg, rgba(69,184,195,0.1), rgba(46,139,145,0.05));
    }
    .menu-card-icon svg { width: 28px; height: 28px; fill: var(--main); }
    .menu-card h2 { font-size: 16px; font-weight: 700; margin: 0 0 6px 0; color: #2c3e50; }
    .menu-card p { font-size: 12px; color: #999; line-height: 1.5; margin: 0; }

    .gecmis-container {
      background: #fff; border-radius: 12px; padding: 22px;
      box-shadow: 0 4px 14px rgba(0,0,0,0.05); margin-bottom: 22px;
    }
    .gecmis-container h2 {
      font-size: 16px; color: #2c3e50; margin: 0 0 15px 0;
      padding-bottom: 12px; border-bottom: 2px solid var(--main);
      display: flex; align-items: center; gap: 10px;
    }
    .gecmis-container h2 svg { width: 22px; height: 22px; fill: var(--main); }
    .gecmis-container h2.diyet-title { border-bottom-color: var(--diyet); }
    .gecmis-container h2.diyet-title svg { fill: var(--diyet); }
    .gecmis-grid {
      display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
      gap: 10px; max-height: 200px; overflow-y: auto;
    }
    .gecmis-grid::-webkit-scrollbar { width: 5px; }
    .gecmis-grid::-webkit-scrollbar-thumb { background: var(--main); border-radius: 10px; }
    .gecmis-item {
      background: #fff; color: var(--dark); border: 2px solid var(--main);
      padding: 12px; border-radius: 10px;
      text-align: center; text-decoration: none; font-weight: 600;
      font-size: 13px; transition: all 0.3s; cursor: pointer;
      white-space: nowrap; display: flex; align-items: center;
      justify-content: center; gap: 6px;
    }
    .gecmis-item:hover {
      background: linear-gradient(135deg, var(--main), var(--dark));
      color: #fff; transform: translateY(-3px);
      box-shadow: 0 6px 16px rgba(69,184,195,0.35);
    }
    .gecmis-item.diyet-item { color: var(--diyet-dark); border-color: var(--diyet); }
    .gecmis-item.diyet-item:hover {
      background: linear-gradient(135deg, var(--diyet), var(--diyet-dark));
      color: #fff; box-shadow: 0 6px 16px rgba(76,175,80,0.35);
    }
    .gecmis-separator { opacity: 0.5; font-weight: 300; font-size: 16px; }

    .modal {
      display: none; position: fixed; z-index: 9999;
      left: 0; top: 0; width: 100%; height: 100%;
      background: rgba(0,0,0,0); transition: background 0.3s;
    }
    .modal.show { background: rgba(0,0,0,0.7); }
    .modal-content {
      background: #fff; margin: 2% auto; border-radius: 14px;
      width: 900px; max-width: 95%; height: 90vh;
      box-shadow: 0 20px 50px rgba(0,0,0,0.3); overflow: hidden;
      transform: scale(0.7); opacity: 0;
      transition: transform 0.3s, opacity 0.3s;
    }
    .modal.show .modal-content { transform: scale(1); opacity: 1; }
    .modal.hide .modal-content { transform: scale(0.7); opacity: 0; }
    .modal.hide { background: rgba(0,0,0,0); }
    .modal-header {
      background: linear-gradient(90deg, var(--main), var(--dark));
      color: #fff; padding: 15px 20px;
      display: flex; justify-content: space-between; align-items: center;
    }
    .modal-header.diyet-modal-header { background: linear-gradient(90deg, var(--diyet), var(--diyet-dark)); }
    .modal-header h2 { margin: 0; font-size: 17px; display: flex; align-items: center; gap: 10px; }
    .modal-header h2 svg { width: 22px; height: 22px; }
    .close {
      color: #fff; font-size: 28px; font-weight: bold; cursor: pointer;
      transition: all 0.3s; width: 35px; height: 35px;
      display: flex; align-items: center; justify-content: center; border-radius: 50%;
    }
    .close:hover { background: rgba(255,255,255,0.2); transform: rotate(90deg); }
    .modal-body { height: calc(90vh - 62px); overflow: hidden; }
    .modal-body iframe { width: 100%; height: 100%; border: none; }

    .content::-webkit-scrollbar { width: 6px; }
    .content::-webkit-scrollbar-track { background: #f1f1f1; }
    .content::-webkit-scrollbar-thumb { background: var(--main); border-radius: 10px; }

    @media (max-width: 900px) { .stats-row { grid-template-columns: 1fr; } }
    @media (max-width: 768px) {
      #header { flex-wrap: wrap; gap: 8px; padding: 10px; }
      #header h1 { font-size: 15px; }
      .modal-content { width: 98%; height: 95vh; }
      .menu-container { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <div id="header">
    <div id="nav-buttons">
      <a href="panel.asp" class="nav-button active" title="Ana Sayfa">
        <svg viewBox="0 0 24 24"><path d="M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z"/></svg>
      </a>
      <a href="yemek_liste.asp" class="nav-button" title="Yemek Listesi">
        <svg viewBox="0 0 24 24"><path d="M3 13h2v-2H3v2zm0 4h2v-2H3v2zm0-8h2V7H3v2zm4 4h14v-2H7v2zm0 4h14v-2H7v2zM7 7v2h14V7H7z"/></svg>
      </a>
      <a href="panel.asp?islem=cikis" class="nav-button" title="&#199;&#305;k&#305;&#351; Yap" onclick="return confirm('\u00c7\u0131k\u0131\u015f yapmak istedi\u011finize emin misiniz?')">
        <svg viewBox="0 0 24 24"><path d="M17 7l-1.41 1.41L18.17 11H8v2h10.17l-2.58 2.58L17 17l5-5zM4 5h8V3H4c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h8v-2H4V5z"/></svg>
      </a>
    </div>
    <h1>
      <svg viewBox="0 0 24 24" fill="currentColor"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>
      Yemek Sistemi - Y&#246;netici Paneli
    </h1>
    <div id="user-info">
      <svg viewBox="0 0 24 24"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>
      <strong><%= admin_kullanici %></strong>
    </div>
  </div>

  <div class="content">

    <div class="stats-row">

      <div class="stat-genel">
        <div class="stat-genel-header">
          <div class="sg-icon"><svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg></div>
          <div class="sg-info">
            <span>Normal Men&#252;</span>
            <strong><%= toplamAylik %> <small>ayl&#305;k men&#252;</small></strong>
          </div>
        </div>
        <div class="yil-liste">
          <% Dim yKey: For Each yKey In dictYillar.Keys %>
          <div class="yil-item">
            <div class="yil-item-left">
              <svg viewBox="0 0 24 24"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10z"/></svg>
              <span><%= yKey %></span>
            </div>
            <div class="yil-item-right"><%= dictYillar(yKey) %> ay</div>
          </div>
          <% Next %>
          <% If toplamAylikDiyet > 0 Then %>
          <% Dim yKeyD: For Each yKeyD In dictYillarDiyet.Keys %>
          <div class="yil-item diyet-item">
            <div class="yil-item-left">
              <svg viewBox="0 0 24 24"><path d="M17.21 9l-4.38-6.56c-.19-.28-.51-.42-.83-.42-.32 0-.64.14-.83.43L6.79 9C6.3 9.71 6 10.57 6 11.5 6 14.53 8.47 17 11.5 17h1c3.03 0 5.5-2.47 5.5-5.5 0-.93-.3-1.79-.79-2.5z"/></svg>
              <span><%= yKeyD %> (Diyet)</span>
            </div>
            <div class="yil-item-right"><%= dictYillarDiyet(yKeyD) %> ay</div>
          </div>
          <% Next %>
          <% End If %>
        </div>
      </div>

      <div class="stat-ay <%= buAyDurum %>">
        <div class="stat-ay-top">
          <div class="stat-ay-label">Bu Ay</div>
        </div>
        <div class="stat-ay-body">
          <div class="stat-ay-icon">
            <% If buAyDurum = "girildi" Then %>
            <svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>
            <% Else %>
            <svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
            <% End If %>
          </div>
          <div class="stat-ay-title"><%= GetMonthName(bugun_ay) %> <span class="sep">|</span> <%= bugun_yil %></div>
          <div class="stat-ay-gun">(<%= buAyGun %> G&#252;n)</div>
          <div class="stat-ay-badges">
            <span class="stat-ay-badge normal-badge <%= buAyDurum %>">
              <% If buAyDurum = "girildi" Then %><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><% Else %><svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg><% End If %>
              Normal
            </span>
            <span class="stat-ay-badge diyet-badge <%= buAyDurumDiyet %>">
              <% If buAyDurumDiyet = "girildi" Then %><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><% Else %><svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg><% End If %>
              Diyet
            </span>
          </div>
        </div>
      </div>

      <div class="stat-ay <%= gelecekDurum %>">
        <div class="stat-ay-top">
          <div class="stat-ay-label">Gelecek Ay</div>
        </div>
        <div class="stat-ay-body">
          <div class="stat-ay-icon">
            <% If gelecekDurum = "girildi" Then %>
            <svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>
            <% Else %>
            <svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
            <% End If %>
          </div>
          <div class="stat-ay-title"><%= GetMonthName(gelecekAy) %> <span class="sep">|</span> <%= gelecekYil %></div>
          <div class="stat-ay-gun">(<%= gelecekAyGun %> G&#252;n)</div>
          <div class="stat-ay-badges">
            <span class="stat-ay-badge normal-badge <%= gelecekDurum %>">
              <% If gelecekDurum = "girildi" Then %><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><% Else %><svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg><% End If %>
              Normal
            </span>
            <span class="stat-ay-badge diyet-badge <%= gelecekDurumDiyet %>">
              <% If gelecekDurumDiyet = "girildi" Then %><svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><% Else %><svg viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg><% End If %>
              Diyet
            </span>
          </div>
        </div>
      </div>

    </div>

    <div class="menu-container">
      <a href="yemek_liste.asp" class="menu-card">
        <div class="menu-card-icon"><svg viewBox="0 0 24 24"><path d="M3 13h2v-2H3v2zm0 4h2v-2H3v2zm0-8h2V7H3v2zm4 4h14v-2H7v2zm0 4h14v-2H7v2zM7 7v2h14V7H7z"/></svg></div>
        <h2>Yemek Listesi</h2>
        <p>T&#252;m kay&#305;tlar&#305; g&#246;r&#252;nt&#252;le, d&#252;zenle veya sil</p>
      </a>
      <a href="yemek_ekle.asp" class="menu-card">
        <div class="menu-card-icon"><svg viewBox="0 0 24 24"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg></div>
        <h2>Yeni Kay&#305;t Ekle</h2>
        <p>Yeni yemek men&#252;s&#252; kayd&#305; olu&#351;tur</p>
      </a>
      <a href="javascript:void(0);" onclick="openPreview()" class="menu-card">
        <div class="menu-card-icon"><svg viewBox="0 0 24 24"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg></div>
        <h2>&#214;n &#304;zleme</h2>
        <p>Aktif listenin ziyaret&#231;i g&#246;r&#252;n&#252;m&#252;</p>
      </a>
      <a href="yemek_kisi_giris.asp" class="menu-card">
        <div class="menu-card-icon"><svg viewBox="0 0 24 24"><path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg></div>
        <h2>Ki&#351;i Say&#305;s&#305;</h2>
        <p>G&#252;nl&#252;k yemek yiyen ki&#351;i say&#305;s&#305; giri&#351;i</p>
      </a>
    </div>

    <!-- GECMIS DONEM - NORMAL -->
    <div class="gecmis-container">
      <h2>
        <svg viewBox="0 0 24 24"><path d="M20.54 5.23l-1.39-1.68C18.88 3.21 18.47 3 18 3H6c-.47 0-.88.21-1.16.55L3.46 5.23C3.17 5.57 3 6.02 3 6.5V19c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V6.5c0-.48-.17-.93-.46-1.27zM12 17.5L6.5 12H10v-2h4v2h3.5L12 17.5zM5.12 5l.81-1h12l.94 1H5.12z"/></svg>
        Ge&#231;mi&#351; D&#246;nem Listeleri
      </h2>
      <div class="gecmis-grid">
        <% If Not rsGecmis.EOF Then
          Do While Not rsGecmis.EOF
            Dim ay_adi, yil_no
            yil_no = rsGecmis("yil")
            ay_adi = GetMonthName(rsGecmis("ay"))
        %>
        <a href="javascript:void(0);" onclick="openGecmis(<%= rsGecmis("yil") %>, <%= rsGecmis("ay") %>, 'normal')" class="gecmis-item">
          <span><%= ay_adi %></span>
          <span class="gecmis-separator">|</span>
          <span><%= yil_no %></span>
        </a>
        <%  rsGecmis.MoveNext
          Loop
        Else %>
        <p style="text-align:center; color:#999; grid-column: 1/-1; padding: 20px;">Hen&#252;z ge&#231;mi&#351; d&#246;nem kayd&#305; bulunmuyor</p>
        <% End If %>
      </div>
    </div>

    <!-- GECMIS DONEM - DIYET -->
    <% If diyetGecmisVar Then %>
    <div class="gecmis-container">
      <h2 class="diyet-title">
        <svg viewBox="0 0 24 24"><path d="M17.21 9l-4.38-6.56c-.19-.28-.51-.42-.83-.42-.32 0-.64.14-.83.43L6.79 9C6.3 9.71 6 10.57 6 11.5 6 14.53 8.47 17 11.5 17h1c3.03 0 5.5-2.47 5.5-5.5 0-.93-.3-1.79-.79-2.5z"/></svg>
        Ge&#231;mi&#351; D&#246;nem Diyet Listeleri
      </h2>
      <div class="gecmis-grid">
        <% If Not rsGecmisDiyet.EOF Then
          Do While Not rsGecmisDiyet.EOF
            Dim ay_adi_d, yil_no_d
            yil_no_d = rsGecmisDiyet("yil")
            ay_adi_d = GetMonthName(rsGecmisDiyet("ay"))
        %>
        <a href="javascript:void(0);" onclick="openGecmis(<%= rsGecmisDiyet("yil") %>, <%= rsGecmisDiyet("ay") %>, 'diyet')" class="gecmis-item diyet-item">
          <span><%= ay_adi_d %></span>
          <span class="gecmis-separator">|</span>
          <span><%= yil_no_d %></span>
        </a>
        <%  rsGecmisDiyet.MoveNext
          Loop
        Else %>
        <p style="text-align:center; color:#999; grid-column: 1/-1; padding: 20px;">Hen&#252;z ge&#231;mi&#351; diyet d&#246;nem kayd&#305; bulunmuyor</p>
        <% End If %>
      </div>
    </div>
    <% End If %>

  </div>

  <div id="previewModal" class="modal">
    <div class="modal-content">
      <div class="modal-header">
        <h2>
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>
          Yemek Listesi &#214;n &#304;zleme
        </h2>
        <span class="close" onclick="closePreview()">&times;</span>
      </div>
      <div class="modal-body">
        <iframe id="previewFrame" src=""></iframe>
      </div>
    </div>
  </div>

  <div id="gecmisModal" class="modal">
    <div class="modal-content">
      <div class="modal-header" id="gecmisModalHeader">
        <h2 id="gecmisTitle">
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M20.54 5.23l-1.39-1.68C18.88 3.21 18.47 3 18 3H6c-.47 0-.88.21-1.16.55L3.46 5.23C3.17 5.57 3 6.02 3 6.5V19c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V6.5c0-.48-.17-.93-.46-1.27zM12 17.5L6.5 12H10v-2h4v2h3.5L12 17.5zM5.12 5l.81-1h12l.94 1H5.12z"/></svg>
          <span id="gecmisTitleText">Ge&#231;mi&#351; D&#246;nem Listesi</span>
        </h2>
        <span class="close" onclick="closeGecmis()">&times;</span>
      </div>
      <div class="modal-body">
        <iframe id="gecmisFrame" src=""></iframe>
      </div>
    </div>
  </div>

  <script>
    function openPreview() {
      var modal = document.getElementById('previewModal');
      modal.style.display = 'block';
      document.body.style.overflow = 'hidden';
      setTimeout(function() { modal.classList.add('show'); }, 10);
      setTimeout(function() { document.getElementById('previewFrame').src = 'yemek_index.asp'; }, 350);
    }

    function closePreview() {
      var modal = document.getElementById('previewModal');
      modal.classList.remove('show');
      modal.classList.add('hide');
      setTimeout(function() {
        modal.style.display = 'none';
        modal.classList.remove('hide');
        document.getElementById('previewFrame').src = '';
        document.body.style.overflow = 'auto';
      }, 300);
    }

    function openGecmis(yil, ay, tip) {
      var modal = document.getElementById('gecmisModal');
      var header = document.getElementById('gecmisModalHeader');
      var ayAdlari = ['', 'Ocak', '\u015eubat', 'Mart', 'Nisan', 'May\u0131s', 'Haziran', 'Temmuz', 'A\u011fustos', 'Eyl\u00fcl', 'Ekim', 'Kas\u0131m', 'Aral\u0131k'];

      if (tip === 'diyet') {
        document.getElementById('gecmisTitleText').textContent = ayAdlari[ay] + ' | ' + yil + ' - Diyet Yemek Listesi';
        header.className = 'modal-header diyet-modal-header';
      } else {
        document.getElementById('gecmisTitleText').textContent = ayAdlari[ay] + ' | ' + yil + ' - Yemek Listesi';
        header.className = 'modal-header';
      }

      modal.style.display = 'block';
      document.body.style.overflow = 'hidden';
      setTimeout(function() { modal.classList.add('show'); }, 10);

      var url = 'yemek_gecmis.asp?yil=' + yil + '&ay=' + ay;
      if (tip === 'diyet') url += '&tip=diyet';
      setTimeout(function() { document.getElementById('gecmisFrame').src = url; }, 350);
    }

    function closeGecmis() {
      var modal = document.getElementById('gecmisModal');
      modal.classList.remove('show');
      modal.classList.add('hide');
      setTimeout(function() {
        modal.style.display = 'none';
        modal.classList.remove('hide');
        document.getElementById('gecmisFrame').src = '';
        document.body.style.overflow = 'auto';
      }, 300);
    }

    window.onclick = function(event) {
      if (event.target.className.indexOf('modal') > -1 && event.target.className.indexOf('show') > -1) {
        if (event.target.id === 'previewModal') closePreview();
        else if (event.target.id === 'gecmisModal') closeGecmis();
      }
    };

    document.addEventListener('keydown', function(event) {
      if (event.key === 'Escape') { closePreview(); closeGecmis(); }
    });

    var inactivityTimeout;
    var inactivityLimit = 120000;
    function resetInactivityTimer() {
      clearTimeout(inactivityTimeout);
      inactivityTimeout = setTimeout(function() {
        window.location.href = 'panel.asp?islem=cikis';
      }, inactivityLimit);
    }
    document.addEventListener('mousemove', resetInactivityTimer);
    document.addEventListener('keypress', resetInactivityTimer);
    document.addEventListener('click', resetInactivityTimer);
    document.addEventListener('scroll', resetInactivityTimer);
    document.addEventListener('touchstart', resetInactivityTimer);
    resetInactivityTimer();
    document.addEventListener('visibilitychange', function() {
      if (!document.hidden) resetInactivityTimer();
    });
  </script>
</body>
</html>
<%
rsGecmis.Close
Set rsGecmis = Nothing
If diyetGecmisVar Then
  If Not rsGecmisDiyet Is Nothing Then rsGecmisDiyet.Close: Set rsGecmisDiyet = Nothing
End If
Set dictYillar = Nothing
Set dictYillarDiyet = Nothing
%>
