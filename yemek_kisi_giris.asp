<!-- #include file="database/Connection.asp" -->
<%
If Session("yemek_admin_giris") <> "OK" Then
    Response.Redirect "giris.asp"
End If
Dim admin_kullanici
admin_kullanici = Session("yemek_admin_kullanici")
Session.Timeout = 30

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

Dim basarili_mesaj, hata_mesaj

If Request.QueryString("sil") <> "" Then
    Dim silId
    silId = Request.QueryString("sil")
    If IsNumeric(silId) Then
        On Error Resume Next
        ConnYemek.Execute "DELETE FROM yemek_kisi_sayisi WHERE id = " & CInt(silId)
        If Err.Number = 0 Then
            basarili_mesaj = "Kay&#305;t silindi!"
        Else
            hata_mesaj = "Silme s&#305;ras&#305;nda hata olu&#351;tu!"
        End If
        Err.Clear
        On Error GoTo 0
    End If
End If

If Request.Form("btn_kaydet") <> "" Then
    Dim frm_tarih, frm_tip, frm_kisi
    frm_tarih = Request.Form("tarih")
    frm_tip = Request.Form("menu_tipi")
    frm_kisi = Request.Form("kisi_sayisi")

    If frm_tarih = "" Or frm_tip = "" Or frm_kisi = "" Or Not IsNumeric(frm_kisi) Then
        hata_mesaj = "T&#252;m alanlar&#305; do&#287;ru doldurun!"
    Else
        Dim sqlKontrol, rsKontrol
        sqlKontrol = "SELECT COUNT(*) AS adet FROM yemek_kisi_sayisi WHERE tarih = #" & frm_tarih & "# AND menu_tipi = '" & frm_tip & "'"
        On Error Resume Next
        Set rsKontrol = ConnYemek.Execute(sqlKontrol)
        If Err.Number <> 0 Then
            hata_mesaj = "yemek_kisi_sayisi tablosu bulunamad&#305;! L&#252;tfen tabloyu olu&#351;turun."
            Err.Clear
            On Error GoTo 0
        Else
            On Error GoTo 0
            If rsKontrol("adet") > 0 Then
                Dim sqlUpdate
                sqlUpdate = "UPDATE yemek_kisi_sayisi SET kisi_sayisi = " & CInt(frm_kisi) & ", kullanici = '" & Replace(admin_kullanici, "'", "''") & "', kayit_tarihi = Now() WHERE tarih = #" & frm_tarih & "# AND menu_tipi = '" & frm_tip & "'"
                ConnYemek.Execute sqlUpdate
                basarili_mesaj = "Kay&#305;t g&#252;ncellendi!"
            Else
                Dim sqlInsert
                sqlInsert = "INSERT INTO yemek_kisi_sayisi (tarih, menu_tipi, kisi_sayisi, kullanici, kayit_tarihi) VALUES (#" & frm_tarih & "#, '" & frm_tip & "', " & CInt(frm_kisi) & ", '" & Replace(admin_kullanici, "'", "''") & "', Now())"
                ConnYemek.Execute sqlInsert
                basarili_mesaj = "Kay&#305;t eklendi!"
            End If
            rsKontrol.Close
            Set rsKontrol = Nothing
        End If
    End If
End If

Dim bugun_tarih
bugun_tarih = Year(Now()) & "-" & Right("0" & Month(Now()), 2) & "-" & Right("0" & Day(Now()), 2)

Dim rsNormal, rsDiyet, bugunNormal, bugunDiyet
bugunNormal = ""
bugunDiyet = ""
On Error Resume Next
Set rsNormal = ConnYemek.Execute("SELECT kisi_sayisi FROM yemek_kisi_sayisi WHERE tarih = #" & Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()) & "# AND menu_tipi = 'normal'")
If Err.Number = 0 And Not rsNormal.EOF Then bugunNormal = rsNormal("kisi_sayisi")
If Not rsNormal Is Nothing Then rsNormal.Close: Set rsNormal = Nothing
Set rsDiyet = ConnYemek.Execute("SELECT kisi_sayisi FROM yemek_kisi_sayisi WHERE tarih = #" & Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()) & "# AND menu_tipi = 'diyet'")
If Err.Number = 0 And Not rsDiyet.EOF Then bugunDiyet = rsDiyet("kisi_sayisi")
If Not rsDiyet Is Nothing Then rsDiyet.Close: Set rsDiyet = Nothing
Err.Clear
On Error GoTo 0

Dim rsSon10, sqlSon10, son10Var
son10Var = False
On Error Resume Next
sqlSon10 = "SELECT TOP 10 * FROM yemek_kisi_sayisi ORDER BY tarih DESC, menu_tipi ASC"
Set rsSon10 = ConnYemek.Execute(sqlSon10)
If Err.Number = 0 Then son10Var = True
Err.Clear
On Error GoTo 0
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Ki&#351;i Say&#305;s&#305; Giri&#351;i</title>
  <link rel="icon" href="../../images/hastane_portal_logo.png" type="image/png">
  <style>
    :root { --main: #45b8c3; --dark: #2e8b91; --txt: #fff; --diyet: #4caf50; --diyet-dark: #388e3c; }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    html, body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; min-height: 100vh; background: linear-gradient(135deg, #f0f9fa, #e6f4f1); }
    #header { background: linear-gradient(90deg, var(--main), var(--dark)); padding: 8px 15px; color: var(--txt); display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 10px rgba(0,0,0,0.15); position: sticky; top: 0; z-index: 100; }
    #nav-buttons { display: flex; gap: 8px; }
    .nav-button { background: transparent; color: var(--txt); width: 36px; height: 36px; border: 2px solid var(--txt); border-radius: 50%; text-decoration: none; transition: all 0.3s; display: flex; align-items: center; justify-content: center; }
    .nav-button:hover { background: var(--dark); transform: translateY(-2px); }
    .nav-button svg { width: 20px; height: 20px; fill: currentColor; }
    #header h1 { font-size: 18px; margin: 0; display: flex; align-items: center; gap: 8px; }
    #header h1 svg { width: 20px; height: 20px; opacity: 0.85; }
    .content { padding: 20px; max-width: 900px; margin: 0 auto; }
    .alert { padding: 12px 20px; border-radius: 8px; margin-bottom: 15px; font-size: 14px; display: flex; align-items: center; gap: 10px; animation: slideDown 0.3s; }
    @keyframes slideDown { from { opacity: 0; transform: translateY(-10px); } to { opacity: 1; transform: translateY(0); } }
    .alert-success { background: #e8f5e9; color: #2e7d32; border-left: 4px solid #4caf50; }
    .alert-error { background: #ffebee; color: #c62828; border-left: 4px solid #f44336; }
    .alert svg { width: 20px; height: 20px; flex-shrink: 0; }

    .form-card { background: #fff; border-radius: 14px; box-shadow: 0 6px 20px rgba(0,0,0,0.06); overflow: hidden; margin-bottom: 20px; }
    .form-card-header { background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; padding: 18px 22px; display: flex; align-items: center; gap: 10px; }
    .form-card-header h2 { font-size: 17px; margin: 0; }
    .form-card-header svg { width: 22px; height: 22px; }
    .form-card-body { padding: 24px; }
    .form-row { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; margin-bottom: 20px; }
    .form-group label { display: block; font-size: 12px; font-weight: 600; color: #555; margin-bottom: 6px; text-transform: uppercase; }
    .form-group input, .form-group select { width: 100%; padding: 11px 14px; border: 2px solid #e8e8e8; border-radius: 10px; font-size: 14px; outline: none; transition: all 0.3s; font-family: inherit; }
    .form-group input:focus, .form-group select:focus { border-color: var(--main); box-shadow: 0 0 0 3px rgba(69,184,195,0.15); }
    .btn-kaydet { width: 100%; padding: 14px; border: none; border-radius: 10px; font-size: 15px; font-weight: 600; cursor: pointer; background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; transition: all 0.3s; display: flex; align-items: center; justify-content: center; gap: 8px; }
    .btn-kaydet:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(69,184,195,0.35); }
    .btn-kaydet svg { width: 20px; height: 20px; }

    .bugun-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-bottom: 20px; }
    .bugun-card { background: #fff; border-radius: 12px; padding: 18px; box-shadow: 0 4px 14px rgba(0,0,0,0.05); display: flex; align-items: center; gap: 14px; border-left: 4px solid var(--main); }
    .bugun-card.diyet { border-left-color: var(--diyet); }
    .bugun-card-icon { width: 44px; height: 44px; border-radius: 10px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
    .bugun-card-icon.normal { background: linear-gradient(135deg, var(--main), var(--dark)); }
    .bugun-card-icon.diyet { background: linear-gradient(135deg, var(--diyet), var(--diyet-dark)); }
    .bugun-card-icon svg { width: 22px; height: 22px; fill: #fff; }
    .bugun-card-info span { display: block; font-size: 10px; color: #999; text-transform: uppercase; }
    .bugun-card-info strong { font-size: 22px; color: #2c3e50; }
    .bugun-card-info small { font-size: 11px; color: #888; }

    .son-kayitlar { background: #fff; border-radius: 14px; box-shadow: 0 6px 20px rgba(0,0,0,0.06); overflow: hidden; }
    .son-kayitlar-header { padding: 16px 22px; border-bottom: 1px solid #f0f0f0; display: flex; align-items: center; gap: 8px; }
    .son-kayitlar-header h3 { font-size: 15px; color: #2c3e50; margin: 0; }
    .son-kayitlar-header svg { width: 18px; height: 18px; fill: var(--main); }
    .son-tablo { width: 100%; border-collapse: collapse; }
    .son-tablo th { background: #f8f9fa; padding: 10px 16px; font-size: 11px; text-align: left; color: #666; text-transform: uppercase; font-weight: 600; }
    .son-tablo td { padding: 10px 16px; font-size: 13px; border-top: 1px solid #f5f5f5; }
    .son-tablo tr:hover td { background: #f0fffe; }
    .tip-badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 10px; font-weight: 600; }
    .tip-badge.normal { background: rgba(69,184,195,0.12); color: var(--dark); }
    .tip-badge.diyet { background: rgba(76,175,80,0.12); color: var(--diyet-dark); }

    .islem-td { display: flex; gap: 6px; }
    .btn-duzenle, .btn-sil { width: 30px; height: 30px; border-radius: 8px; border: none; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: all 0.3s; }
    .btn-duzenle { background: rgba(69,184,195,0.1); }
    .btn-duzenle svg { width: 14px; height: 14px; fill: var(--dark); }
    .btn-duzenle:hover { background: var(--main); }
    .btn-duzenle:hover svg { fill: #fff; }
    .btn-sil { background: rgba(229,57,53,0.1); }
    .btn-sil svg { width: 14px; height: 14px; fill: #e53935; }
    .btn-sil:hover { background: #e53935; }
    .btn-sil:hover svg { fill: #fff; }

    @media (max-width: 768px) { .form-row { grid-template-columns: 1fr; } .bugun-row { grid-template-columns: 1fr; } }
  </style>
  <script>
    function duzenleSatir(tarih, tip, kisi) {
      document.querySelector('input[name="tarih"]').value = tarih;
      document.querySelector('select[name="menu_tipi"]').value = tip;
      document.querySelector('input[name="kisi_sayisi"]').value = kisi;
      document.querySelector('.form-card').scrollIntoView({ behavior: 'smooth', block: 'center' });
      document.querySelector('input[name="kisi_sayisi"]').focus();
    }
    function silSatir(id) {
      if (confirm('Bu kayd\u0131 silmek istedi\u011finize emin misiniz?')) {
        window.location.href = 'yemek_kisi_giris.asp?sil=' + id;
      }
    }
  </script>
</head>
<body>
<div id="header">
  <div id="nav-buttons">
    <a href="panel.asp" class="nav-button" title="Panel"><svg viewBox="0 0 24 24"><path d="M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z"/></svg></a>
    <a href="panel.asp" class="nav-button" title="Geri"><svg viewBox="0 0 24 24"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg></a>
  </div>
  <h1><svg viewBox="0 0 24 24" fill="currentColor"><path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg>G&#252;nl&#252;k Ki&#351;i Say&#305;s&#305; Giri&#351;i</h1>
  <div style="width:36px;"></div>
</div>

<div class="content">
  <% If basarili_mesaj <> "" Then %>
  <div class="alert alert-success"><svg viewBox="0 0 24 24" fill="currentColor"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg><%= basarili_mesaj %></div>
  <% End If %>
  <% If hata_mesaj <> "" Then %>
  <div class="alert alert-error"><svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg><%= hata_mesaj %></div>
  <% End If %>

  <div class="bugun-row">
    <div class="bugun-card">
      <div class="bugun-card-icon normal"><svg viewBox="0 0 24 24"><path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg></div>
      <div class="bugun-card-info"><span>Bug&#252;n Normal</span><strong><% If bugunNormal <> "" Then %><%= bugunNormal %><% Else %>-<% End If %></strong><small>Ki&#351;i</small></div>
    </div>
    <div class="bugun-card diyet">
      <div class="bugun-card-icon diyet"><svg viewBox="0 0 24 24"><path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg></div>
      <div class="bugun-card-info"><span>Bug&#252;n Diyet</span><strong><% If bugunDiyet <> "" Then %><%= bugunDiyet %><% Else %>-<% End If %></strong><small>Ki&#351;i</small></div>
    </div>
  </div>

  <div class="form-card">
    <div class="form-card-header">
      <svg viewBox="0 0 24 24" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
      <h2>Ki&#351;i Say&#305;s&#305; Ekle / G&#252;ncelle</h2>
    </div>
    <div class="form-card-body">
      <form method="post" action="yemek_kisi_giris.asp">
        <div class="form-row">
          <div class="form-group">
            <label>Tarih</label>
            <input type="date" name="tarih" value="<%= bugun_tarih %>" required>
          </div>
          <div class="form-group">
            <label>Men&#252; Tipi</label>
            <select name="menu_tipi" required>
              <option value="normal">Normal Men&#252;</option>
              <option value="diyet">Diyet Men&#252;s&#252;</option>
            </select>
          </div>
          <div class="form-group">
            <label>Ki&#351;i Say&#305;s&#305;</label>
            <input type="number" name="kisi_sayisi" min="0" placeholder="&#214;rn: 350" required>
          </div>
        </div>
        <button type="submit" name="btn_kaydet" value="1" class="btn-kaydet">
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M17 3H5c-1.11 0-2 .9-2 2v14c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V7l-4-4zm-5 16c-1.66 0-3-1.34-3-3s1.34-3 3-3 3 1.34 3 3-1.34 3-3 3zm3-10H5V5h10v4z"/></svg>
          Kaydet
        </button>
      </form>
    </div>
  </div>

  <% If son10Var Then %>
  <div class="son-kayitlar">
    <div class="son-kayitlar-header">
      <svg viewBox="0 0 24 24"><path d="M13 3c-4.97 0-9 4.03-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42C8.27 19.99 10.51 21 13 21c4.97 0 9-4.03 9-9s-4.03-9-9-9zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>
      <h3>Son Kay&#305;tlar</h3>
    </div>
    <table class="son-tablo">
      <thead><tr><th>Tarih</th><th>Tip</th><th>Ki&#351;i</th><th>&#304;&#351;lem</th></tr></thead>
      <tbody>
        <% If Not rsSon10.EOF Then
          Do While Not rsSon10.EOF
            Dim kayitTarihStr
            kayitTarihStr = Year(rsSon10("tarih")) & "-" & Right("0" & Month(rsSon10("tarih")), 2) & "-" & Right("0" & Day(rsSon10("tarih")), 2)
        %>
        <tr>
          <td><strong><%= Right("0" & Day(rsSon10("tarih")), 2) %>.<%= Right("0" & Month(rsSon10("tarih")), 2) %>.<%= Year(rsSon10("tarih")) %></strong></td>
          <td><span class="tip-badge <%= rsSon10("menu_tipi") %>"><% If rsSon10("menu_tipi") = "diyet" Then %>Diyet<% Else %>Normal<% End If %></span></td>
          <td><strong><%= rsSon10("kisi_sayisi") %></strong></td>
          <td class="islem-td">
            <button class="btn-duzenle" onclick="duzenleSatir('<%= kayitTarihStr %>', '<%= rsSon10("menu_tipi") %>', <%= rsSon10("kisi_sayisi") %>)" title="D&#252;zenle">
              <svg viewBox="0 0 24 24"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
            </button>
            <button class="btn-sil" onclick="silSatir(<%= rsSon10("id") %>)" title="Sil">
              <svg viewBox="0 0 24 24"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
            </button>
          </td>
        </tr>
        <%  rsSon10.MoveNext
          Loop
        End If %>
      </tbody>
    </table>
  </div>
  <% End If %>
</div>
</body>
</html>
<%
If son10Var Then
  If Not rsSon10 Is Nothing Then rsSon10.Close: Set rsSon10 = Nothing
End If
%>
