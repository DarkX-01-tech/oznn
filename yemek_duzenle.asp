<!-- #include file="database/Connection.asp" -->
<%
'====================================
' Session Kontrolü
'====================================
If Session("yemek_admin_giris") <> "OK" Then
    Response.Redirect "giris.asp"
End If

Dim admin_kullanici
admin_kullanici = Session("yemek_admin_kullanici")

'====================================
' ID ve Tip Kontrolü
'====================================
If Request.QueryString("id") = "" Then
    Response.Redirect "yemek_liste.asp"
End If

Dim kayit_id, menu_tipi, hedef_tablo, tip_param
kayit_id = Request.QueryString("id")

If Request.QueryString("tip") = "diyet" Then
    menu_tipi = "diyet"
    hedef_tablo = "diyet_yemek_listesi"
    tip_param = "&tip=diyet"
Else
    menu_tipi = "normal"
    hedef_tablo = "yemek_listesi"
    tip_param = ""
End If

'====================================
' Kayıt Bilgilerini Çek
'====================================
Dim rsKayit, sqlKayit
sqlKayit = "SELECT * FROM " & hedef_tablo & " WHERE id = " & kayit_id
Set rsKayit = ConnYemek.Execute(sqlKayit)

If rsKayit.EOF Then
    If menu_tipi = "diyet" Then
        Response.Redirect "yemek_liste.asp?sekme=diyet&durum=hata"
    Else
        Response.Redirect "yemek_liste.asp?durum=hata"
    End If
End If

Dim kayit_yil, kayit_ay
kayit_yil = Year(rsKayit("tarih"))
kayit_ay = Month(rsKayit("tarih"))

'====================================
' Form İşlemleri
'====================================
Dim basarili_mesaj, hata_mesaj

If Request.Form("btnGuncelle") <> "" Then
    Dim tarih, gun_adi, ogle_corba, ogle_ana_yemek, ogle_yan_urun, ogle_tatli
    Dim aksam_corba, aksam_ana_yemek, aksam_yan_urun, aksam_tatli, aktif

    tarih = Request.Form("tarih")
    gun_adi = Trim(Request.Form("gun_adi"))
    ogle_corba = Trim(Request.Form("ogle_corba"))
    ogle_ana_yemek = Trim(Request.Form("ogle_ana_yemek"))
    ogle_yan_urun = Trim(Request.Form("ogle_yan_urun"))
    ogle_tatli = Trim(Request.Form("ogle_tatli"))
    aksam_corba = Trim(Request.Form("aksam_corba"))
    aksam_ana_yemek = Trim(Request.Form("aksam_ana_yemek"))
    aksam_yan_urun = Trim(Request.Form("aksam_yan_urun"))
    aksam_tatli = Trim(Request.Form("aksam_tatli"))
    aktif = Request.Form("aktif")

    If aktif = "on" Or aktif = "1" Then
        aktif = True
    Else
        aktif = False
    End If

    If tarih = "" Or gun_adi = "" Then
        hata_mesaj = "Tarih ve G&#252;n Ad&#305; alanlar&#305; zorunludur!"
    Else
        gun_adi = Replace(gun_adi, "'", "''")
        ogle_corba = Replace(ogle_corba, "'", "''")
        ogle_ana_yemek = Replace(ogle_ana_yemek, "'", "''")
        ogle_yan_urun = Replace(ogle_yan_urun, "'", "''")
        ogle_tatli = Replace(ogle_tatli, "'", "''")
        aksam_corba = Replace(aksam_corba, "'", "''")
        aksam_ana_yemek = Replace(aksam_ana_yemek, "'", "''")
        aksam_yan_urun = Replace(aksam_yan_urun, "'", "''")
        aksam_tatli = Replace(aksam_tatli, "'", "''")

        Dim sqlUpdate
        sqlUpdate = "UPDATE " & hedef_tablo & " SET " & _
                    "tarih = #" & tarih & "#, " & _
                    "gun_adi = '" & gun_adi & "', " & _
                    "ogle_corba = '" & ogle_corba & "', " & _
                    "ogle_ana_yemek = '" & ogle_ana_yemek & "', " & _
                    "ogle_yan_urun = '" & ogle_yan_urun & "', " & _
                    "ogle_tatli = '" & ogle_tatli & "', " & _
                    "aksam_corba = '" & aksam_corba & "', " & _
                    "aksam_ana_yemek = '" & aksam_ana_yemek & "', " & _
                    "aksam_yan_urun = '" & aksam_yan_urun & "', " & _
                    "aksam_tatli = '" & aksam_tatli & "', " & _
                    "aktif = " & aktif & " " & _
                    "WHERE id = " & kayit_id

        On Error Resume Next
        ConnYemek.Execute sqlUpdate

        If Err.Number = 0 Then
            Response.Redirect "yemek_liste_detay.asp?yil=" & kayit_yil & "&ay=" & kayit_ay & tip_param & "&durum=guncellendi"
        Else
            hata_mesaj = "G&#252;ncelleme s&#305;ras&#305;nda hata olu&#351;tu: " & Err.Description
        End If
        On Error GoTo 0
    End If

    If hata_mesaj = "" Then
        rsKayit.Close
        Set rsKayit = ConnYemek.Execute(sqlKayit)
        kayit_yil = Year(rsKayit("tarih"))
        kayit_ay = Month(rsKayit("tarih"))
    End If
End If
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Kay&#305;t D&#252;zenle - <% If menu_tipi = "diyet" Then %>Diyet <% End If %>Y&#246;netim Paneli</title>
  <link rel="icon" href="../../images/hastane_portal_logo.png" type="image/png">
  <style>
    :root {
      --main-bg-color: <% If menu_tipi = "diyet" Then %>#4caf50<% Else %>#45b8c3<% End If %>;
      --hover-bg-color: <% If menu_tipi = "diyet" Then %>#388e3c<% Else %>#2e8b91<% End If %>;
      --main-text-color: #ffffff;
      --reset-bg-color: #ff4d4d;
    }

    * { margin: 0; padding: 0; box-sizing: border-box; }

    html, body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      color: #333;
      height: 100vh;
      overflow: hidden;
      background: linear-gradient(to right, #f2f6f7, #eefbf8);
    }

    #header {
      background: linear-gradient(90deg, var(--main-bg-color), var(--hover-bg-color));
      padding: 8px 15px;
      color: var(--main-text-color);
      display: flex;
      align-items: center;
      justify-content: space-between;
      box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    }

    #nav-buttons { display: flex; gap: 8px; }

    .nav-button {
      background-color: transparent;
      color: var(--main-text-color);
      width: 36px; height: 36px;
      border: 2px solid var(--main-text-color);
      border-radius: 50%;
      text-decoration: none;
      transition: all 0.3s ease;
      display: flex; align-items: center; justify-content: center;
      cursor: pointer;
    }

    .nav-button:hover { background-color: var(--hover-bg-color); transform: translateY(-2px); }
    .nav-button svg { width: 20px; height: 20px; fill: currentColor; }

    #header h1 { font-size: 18px; margin: 0; text-shadow: 1px 1px 2px rgba(0,0,0,0.2); display: flex; align-items: center; gap: 8px; }
    .header-badge { background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 12px; font-size: 11px; font-weight: 600; }

    .content { height: calc(100vh - 52px); overflow-y: auto; padding: 15px; }

    .alert {
      padding: 12px 20px; border-radius: 8px; margin-bottom: 15px;
      font-size: 14px; animation: slideDown 0.3s ease;
      display: flex; align-items: center; gap: 10px;
    }
    @keyframes slideDown { from { opacity: 0; transform: translateY(-10px); } to { opacity: 1; transform: translateY(0); } }
    .alert-error { background: #ffebee; color: #c62828; border-left: 4px solid #f44336; }
    .alert svg { width: 20px; height: 20px; flex-shrink: 0; }

    .form-container {
      max-width: 900px; margin: 0 auto;
      background: #fff; border-radius: 10px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.08); overflow: hidden;
    }

    .form-header {
      background: linear-gradient(135deg, var(--main-bg-color), var(--hover-bg-color));
      color: #fff; padding: 15px 20px;
      display: flex; align-items: center; gap: 10px;
    }
    .form-header h2 { font-size: 18px; margin: 0; }
    .form-header svg { width: 24px; height: 24px; }
    .form-header .type-badge { background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 12px; font-size: 11px; font-weight: 600; margin-left: auto; }

    .form-body { padding: 25px; }

    .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-bottom: 15px; }
    .form-group { margin-bottom: 15px; }
    .form-group label { display: block; color: #555; font-size: 13px; font-weight: 600; margin-bottom: 6px; }
    .form-group label.required:after { content: " *"; color: #f44336; }

    .form-group input[type="date"],
    .form-group input[type="text"],
    .form-group select {
      width: 100%; padding: 10px 12px;
      border: 2px solid #e0e0e0; border-radius: 6px;
      font-size: 13px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      transition: all 0.3s ease; outline: none;
    }
    .form-group input:focus, .form-group select:focus {
      border-color: var(--main-bg-color);
      box-shadow: 0 0 0 3px <% If menu_tipi = "diyet" Then %>rgba(76,175,80,0.1)<% Else %>rgba(69,184,195,0.1)<% End If %>;
    }

    .section-header {
      background: linear-gradient(135deg, <% If menu_tipi = "diyet" Then %>rgba(76,175,80,0.1), rgba(56,142,60,0.1)<% Else %>rgba(69,184,195,0.1), rgba(46,139,145,0.1)<% End If %>);
      padding: 10px 15px; border-radius: 6px;
      margin: 20px 0 15px 0;
      border-left: 4px solid var(--main-bg-color);
    }
    .section-header h3 { font-size: 15px; color: #2c3e50; font-weight: 600; margin: 0; }

    .checkbox-group {
      display: flex; align-items: center; gap: 10px;
      padding: 12px; background: #f8f9fa; border-radius: 6px;
      margin-top: 20px; cursor: pointer; transition: all 0.3s ease;
    }
    .checkbox-group:hover { background: #e9ecef; }
    .checkbox-group input[type="checkbox"] { width: 20px; height: 20px; cursor: pointer; accent-color: var(--main-bg-color); }
    .checkbox-group label { margin: 0; cursor: pointer; font-size: 13px; font-weight: 600; color: #2c3e50; }

    .form-buttons {
      display: flex; gap: 10px; margin-top: 25px;
      padding-top: 20px; border-top: 2px solid #f0f0f0;
    }

    .btn {
      padding: 12px 30px; border: none; border-radius: 6px;
      font-size: 14px; font-weight: 600; cursor: pointer;
      transition: all 0.3s ease; display: inline-flex;
      align-items: center; gap: 8px; text-decoration: none;
    }
    .btn svg { width: 18px; height: 18px; }

    .btn-primary { background: linear-gradient(135deg, var(--main-bg-color), var(--hover-bg-color)); color: #fff; }
    .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 6px 15px <% If menu_tipi = "diyet" Then %>rgba(76,175,80,0.4)<% Else %>rgba(69,184,195,0.4)<% End If %>; }

    .btn-back { background: #fff; color: var(--main-bg-color); border: 2px solid var(--main-bg-color); }
    .btn-back:hover { background: var(--main-bg-color); color: #fff; }

    .content::-webkit-scrollbar { width: 6px; }
    .content::-webkit-scrollbar-track { background: #f1f1f1; }
    .content::-webkit-scrollbar-thumb { background: var(--main-bg-color); border-radius: 10px; }

    @media (max-width: 768px) {
      #header { flex-wrap: wrap; padding: 10px; }
      #header h1 { font-size: 16px; }
      .form-row { grid-template-columns: 1fr; }
      .form-buttons { flex-direction: column; }
      .btn { width: 100%; justify-content: center; }
    }
  </style>
</head>
<body>
  <div id="header">
    <div id="nav-buttons">
      <a href="panel.asp" class="nav-button" title="Ana Sayfa">
        <svg viewBox="0 0 24 24"><path d="M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z"/></svg>
      </a>
      <a href="yemek_liste.asp<% If menu_tipi = "diyet" Then %>?sekme=diyet<% End If %>" class="nav-button" title="Yemek Listesi">
        <svg viewBox="0 0 24 24"><path d="M3 13h2v-2H3v2zm0 4h2v-2H3v2zm0-8h2V7H3v2zm4 4h14v-2H7v2zm0 4h14v-2H7v2zM7 7v2h14V7H7z"/></svg>
      </a>
      <a href="yemek_liste_detay.asp?yil=<%= kayit_yil %>&ay=<%= kayit_ay %><%= tip_param %>" class="nav-button" title="Detaya D&#246;n">
        <svg viewBox="0 0 24 24"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>
      </a>
    </div>
    <h1>
      Kay&#305;t D&#252;zenleme
      <% If menu_tipi = "diyet" Then %><span class="header-badge">Diyet</span><% End If %>
    </h1>
    <div style="width: 36px;"></div>
  </div>

  <div class="content">

    <% If hata_mesaj <> "" Then %>
    <div class="alert alert-error">
      <svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg>
      <%= hata_mesaj %>
    </div>
    <% End If %>

    <div class="form-container">
      <div class="form-header">
        <svg viewBox="0 0 24 24" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
        <h2><% If menu_tipi = "diyet" Then %>Diyet <% End If %>Yemek Men&#252;s&#252; Kayd&#305;n&#305; D&#252;zenle</h2>
        <% If menu_tipi = "diyet" Then %><span class="type-badge">Diyet Men&#252;s&#252;</span><% End If %>
      </div>

      <div class="form-body">
        <form method="post" action="yemek_duzenle.asp?id=<%= kayit_id %><%= tip_param %>">

          <div class="form-row">
            <div class="form-group">
              <label for="tarih" class="required">Tarih</label>
              <input type="date" id="tarih" name="tarih" value="<%= Year(rsKayit("tarih")) %>-<%= Right("0" & Month(rsKayit("tarih")), 2) %>-<%= Right("0" & Day(rsKayit("tarih")), 2) %>" required>
            </div>
            <div class="form-group">
              <label for="gun_adi" class="required">G&#252;n Ad&#305;</label>
              <select id="gun_adi" name="gun_adi" required>
                <option value="">Se&#231;iniz...</option>
                <option value="PAZAR" <% If rsKayit("gun_adi") = "PAZAR" Then Response.Write "selected" %>>PAZAR</option>
                <option value="PAZARTES&#304;" <% If InStr(rsKayit("gun_adi"), "PAZARTES") > 0 Then Response.Write "selected" %>>PAZARTES&#304;</option>
                <option value="SALI" <% If rsKayit("gun_adi") = "SALI" Then Response.Write "selected" %>>SALI</option>
                <option value="&#199;AR&#350;AMBA" <% If InStr(rsKayit("gun_adi"), "AR") > 0 And InStr(rsKayit("gun_adi"), "AMBA") > 0 Then Response.Write "selected" %>>&#199;AR&#350;AMBA</option>
                <option value="PER&#350;EMBE" <% If InStr(rsKayit("gun_adi"), "PER") > 0 Then Response.Write "selected" %>>PER&#350;EMBE</option>
                <option value="CUMA" <% If rsKayit("gun_adi") = "CUMA" Then Response.Write "selected" %>>CUMA</option>
                <option value="CUMARTES&#304;" <% If InStr(rsKayit("gun_adi"), "CUMARTES") > 0 Then Response.Write "selected" %>>CUMARTES&#304;</option>
              </select>
            </div>
          </div>

          <div class="section-header">
            <h3>&#214;&#286;LE YEME&#286;&#304;</h3>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label for="ogle_corba">&#199;orba</label>
              <input type="text" id="ogle_corba" name="ogle_corba" value="<%= rsKayit("ogle_corba") %>" placeholder="&#214;rn: Mercimek &#199;orba">
            </div>
            <div class="form-group">
              <label for="ogle_ana_yemek">Ana Yemek</label>
              <input type="text" id="ogle_ana_yemek" name="ogle_ana_yemek" value="<%= rsKayit("ogle_ana_yemek") %>" placeholder="&#214;rn: Tavuklu Pilav">
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label for="ogle_yan_urun">Yan &#220;r&#252;n</label>
              <input type="text" id="ogle_yan_urun" name="ogle_yan_urun" value="<%= rsKayit("ogle_yan_urun") %>" placeholder="&#214;rn: Bulgur Pilav&#305;">
            </div>
            <div class="form-group">
              <label for="ogle_tatli">Tatl&#305; / Meyve</label>
              <input type="text" id="ogle_tatli" name="ogle_tatli" value="<%= rsKayit("ogle_tatli") %>" placeholder="&#214;rn: S&#252;tla&#231;">
            </div>
          </div>

          <div class="section-header">
            <h3>AK&#350;AM YEME&#286;&#304;</h3>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label for="aksam_corba">&#199;orba</label>
              <input type="text" id="aksam_corba" name="aksam_corba" value="<%= rsKayit("aksam_corba") %>" placeholder="&#214;rn: Domates &#199;orba">
            </div>
            <div class="form-group">
              <label for="aksam_ana_yemek">Ana Yemek</label>
              <input type="text" id="aksam_ana_yemek" name="aksam_ana_yemek" value="<%= rsKayit("aksam_ana_yemek") %>" placeholder="&#214;rn: K&#246;fte">
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label for="aksam_yan_urun">Yan &#220;r&#252;n</label>
              <input type="text" id="aksam_yan_urun" name="aksam_yan_urun" value="<%= rsKayit("aksam_yan_urun") %>" placeholder="&#214;rn: Makarna">
            </div>
            <div class="form-group">
              <label for="aksam_tatli">Tatl&#305; / Meyve</label>
              <input type="text" id="aksam_tatli" name="aksam_tatli" value="<%= rsKayit("aksam_tatli") %>" placeholder="&#214;rn: Ayran">
            </div>
          </div>

          <div class="checkbox-group">
            <input type="checkbox" id="aktif" name="aktif" <% If rsKayit("aktif") = True Then Response.Write "checked" %>>
            <label for="aktif">Bu kay&#305;t aktif olsun (Ziyaret&#231;ilere g&#246;sterilsin)</label>
          </div>

          <div class="form-buttons">
            <button type="submit" name="btnGuncelle" value="1" class="btn btn-primary">
              <svg viewBox="0 0 24 24" fill="currentColor"><path d="M17 3H5c-1.11 0-2 .9-2 2v14c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V7l-4-4zm-5 16c-1.66 0-3-1.34-3-3s1.34-3 3-3 3 1.34 3 3-1.34 3-3 3zm3-10H5V5h10v4z"/></svg>
              G&#252;ncelle
            </button>
            <a href="yemek_liste_detay.asp?yil=<%= kayit_yil %>&ay=<%= kayit_ay %><%= tip_param %>" class="btn btn-back">
              <svg viewBox="0 0 24 24" fill="currentColor"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
              &#304;ptal
            </a>
          </div>

        </form>
      </div>
    </div>

  </div>
</body>
</html>
<%
rsKayit.Close
Set rsKayit = Nothing
%>
