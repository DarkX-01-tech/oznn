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
Session.Timeout = 30

'====================================
' Yıl, Ay ve Tip Parametresi Kontrolü
'====================================
If Request.QueryString("yil") = "" Or Request.QueryString("ay") = "" Then
    Response.Redirect "yemek_liste.asp"
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

'====================================
' Ay adı fonksiyonu
'====================================
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

'====================================
' Gün sayısını hesapla
'====================================
Dim gun_sayisi_hesap
Dim rsGunSayisi
Set rsGunSayisi = ConnYemek.Execute("SELECT COUNT(*) as toplam FROM " & hedef_tablo & " WHERE YEAR(tarih) = " & secilen_yil & " AND MONTH(tarih) = " & secilen_ay)
gun_sayisi_hesap = rsGunSayisi("toplam")
rsGunSayisi.Close
Set rsGunSayisi = Nothing

'====================================
' Seçilen Aya Ait Verileri Çek
'====================================
Dim rsYemek, sqlYemek
sqlYemek = "SELECT * FROM " & hedef_tablo & " " & _
           "WHERE YEAR(tarih) = " & secilen_yil & " AND MONTH(tarih) = " & secilen_ay & " " & _
           "ORDER BY tarih ASC"
Set rsYemek = ConnYemek.Execute(sqlYemek)

If rsYemek.EOF Then
    If menu_tipi = "diyet" Then
        Response.Redirect "yemek_liste.asp?sekme=diyet&durum=hata"
    Else
        Response.Redirect "yemek_liste.asp?durum=hata"
    End If
End If

'====================================
' TOPLU GÜNCELLEME İŞLEMİ
'====================================
Dim basarili_mesaj, hata_mesaj

If Request.Form("toplu_guncelle") <> "" Then
    Dim hata_sayisi, basarili_sayisi
    hata_sayisi = 0
    basarili_sayisi = 0

    On Error Resume Next

    Dim kayit_id, gun_adi_str
    Dim ogle_corba, ogle_ana_yemek, ogle_yan_urun, ogle_tatli
    Dim aksam_corba, aksam_ana_yemek, aksam_yan_urun, aksam_tatli

    rsYemek.MoveFirst
    Do While Not rsYemek.EOF
        kayit_id = rsYemek("id")

        gun_adi_str = Trim(Request.Form("gun_adi_" & kayit_id))
        ogle_corba = Trim(Request.Form("ogle_corba_" & kayit_id))
        ogle_ana_yemek = Trim(Request.Form("ogle_ana_" & kayit_id))
        ogle_yan_urun = Trim(Request.Form("ogle_yan_" & kayit_id))
        ogle_tatli = Trim(Request.Form("ogle_tatli_" & kayit_id))
        aksam_corba = Trim(Request.Form("aksam_corba_" & kayit_id))
        aksam_ana_yemek = Trim(Request.Form("aksam_ana_" & kayit_id))
        aksam_yan_urun = Trim(Request.Form("aksam_yan_" & kayit_id))
        aksam_tatli = Trim(Request.Form("aksam_tatli_" & kayit_id))

        gun_adi_str = Replace(gun_adi_str, "'", "''")
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
                    "gun_adi = '" & gun_adi_str & "', " & _
                    "ogle_corba = '" & ogle_corba & "', " & _
                    "ogle_ana_yemek = '" & ogle_ana_yemek & "', " & _
                    "ogle_yan_urun = '" & ogle_yan_urun & "', " & _
                    "ogle_tatli = '" & ogle_tatli & "', " & _
                    "aksam_corba = '" & aksam_corba & "', " & _
                    "aksam_ana_yemek = '" & aksam_ana_yemek & "', " & _
                    "aksam_yan_urun = '" & aksam_yan_urun & "', " & _
                    "aksam_tatli = '" & aksam_tatli & "' " & _
                    "WHERE id = " & kayit_id

        ConnYemek.Execute sqlUpdate

        If Err.Number = 0 Then
            basarili_sayisi = basarili_sayisi + 1
        Else
            hata_sayisi = hata_sayisi + 1
        End If

        Err.Clear
        rsYemek.MoveNext
    Loop

    On Error GoTo 0

    If hata_sayisi = 0 Then
        Response.Redirect "yemek_liste_detay.asp?yil=" & secilen_yil & "&ay=" & secilen_ay & tip_param & "&durum=guncellendi"
    Else
        hata_mesaj = "Toplam " & basarili_sayisi & " kay&#305;t g&#252;ncellendi, " & hata_sayisi & " kay&#305;t g&#252;ncellenemedi!"
    End If

    rsYemek.Close
    Set rsYemek = ConnYemek.Execute(sqlYemek)
End If
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>(<%= secilen_yil %>) <%= ay_adi %> Ay&#305; <% If menu_tipi = "diyet" Then %>Diyet <% End If %>Men&#252; D&#252;zenleme</title>
  <link rel="icon" href="../../images/hastane_portal_logo.png" type="image/png">
  <style>
    :root {
      --main-bg-color: <% If menu_tipi = "diyet" Then %>#4caf50<% Else %>#45b8c3<% End If %>;
      --hover-bg-color: <% If menu_tipi = "diyet" Then %>#388e3c<% Else %>#2e8b91<% End If %>;
      --main-text-color: #ffffff;
    }

    * { margin: 0; padding: 0; box-sizing: border-box; }

    html, body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      color: #333;
      min-height: 100vh;
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
      position: sticky;
      top: 0;
      z-index: 100;
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

    #header h1 {
      font-size: 18px; margin: 0;
      text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
      display: flex; align-items: center; gap: 8px;
    }
    #header h1 svg { width: 20px; height: 20px; opacity: 0.8; }
    .header-badge { background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 12px; font-size: 11px; font-weight: 600; }

    .content { padding: 20px; max-width: 100%; margin: 0 auto; }

    .alert {
      padding: 12px 20px; border-radius: 8px; margin-bottom: 15px;
      font-size: 14px; animation: slideDown 0.3s ease;
      display: flex; align-items: center; gap: 10px;
      max-width: 1200px; margin-left: auto; margin-right: auto;
    }
    @keyframes slideDown { from { opacity: 0; transform: translateY(-10px); } to { opacity: 1; transform: translateY(0); } }
    .alert-error { background: #ffebee; color: #c62828; border-left: 4px solid #f44336; }
    .alert svg { width: 20px; height: 20px; flex-shrink: 0; }

    .toplu-form-container {
      background: #fff; border-radius: 10px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.08);
      overflow: hidden; max-width: 1600px; margin: 0 auto;
    }

    .toplu-form-header {
      background: linear-gradient(135deg, var(--main-bg-color), var(--hover-bg-color));
      color: #fff; padding: 20px;
      display: flex; align-items: center; justify-content: space-between;
      flex-wrap: wrap; gap: 10px;
    }
    .toplu-form-header h2 {
      font-size: 20px; margin: 0;
      display: flex; align-items: center; gap: 10px;
    }
    .toplu-form-header svg { width: 28px; height: 28px; }
    .gun-count-badge {
      background: rgba(255,255,255,0.2); padding: 8px 16px;
      border-radius: 20px; font-size: 14px; font-weight: 600;
    }
    .type-badge-header {
      background: rgba(255,255,255,0.2); padding: 6px 14px;
      border-radius: 20px; font-size: 12px; font-weight: 600;
    }

    .yemek-tablo-wrapper {
      max-height: calc(100vh - 220px);
      overflow-y: auto; overflow-x: auto; padding: 15px;
    }
    .yemek-grid { display: grid; gap: 12px; }

    .gun-card {
      background: #fff; border: 2px solid #e0e0e0;
      border-radius: 8px; overflow: hidden; transition: all 0.3s ease;
    }
    .gun-card:hover {
      border-color: var(--main-bg-color);
      box-shadow: 0 4px 12px <% If menu_tipi = "diyet" Then %>rgba(76,175,80,0.2)<% Else %>rgba(69,184,195,0.2)<% End If %>;
    }

    .gun-card-header {
      background: linear-gradient(135deg, var(--main-bg-color), var(--hover-bg-color));
      color: #fff; padding: 10px 15px; font-weight: 600; font-size: 14px;
      display: flex; align-items: center; justify-content: space-between;
    }
    .gun-date { font-size: 15px; font-weight: 700; }
    .gun-name {
      font-size: 12px; opacity: 0.9;
      background: rgba(255,255,255,0.2); padding: 3px 10px; border-radius: 10px;
    }

    .gun-card-body {
      padding: 12px; display: grid;
      grid-template-columns: 1fr 1fr; gap: 12px;
    }

    .yemek-section {
      border: 1px solid #f0f0f0; border-radius: 6px;
      padding: 10px; background: #fafafa;
    }
    .yemek-section h4 {
      font-size: 12px; color: #666; margin-bottom: 8px;
      text-transform: uppercase; letter-spacing: 0.5px; font-weight: 700;
      display: flex; align-items: center; gap: 4px;
    }
    .yemek-section h4 svg { width: 14px; height: 14px; fill: #666; }

    .yemek-inputs { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; }

    .yemek-input {
      width: 100%; padding: 6px 8px;
      border: 1px solid #ddd; border-radius: 4px;
      font-size: 11px; transition: all 0.3s ease; outline: none;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .yemek-input:focus {
      border-color: var(--main-bg-color);
      box-shadow: 0 0 0 2px <% If menu_tipi = "diyet" Then %>rgba(76,175,80,0.1)<% Else %>rgba(69,184,195,0.1)<% End If %>;
      background: #fff;
    }
    .yemek-input::placeholder { color: #999; font-style: italic; font-size: 10px; }

    .sticky-kaydet {
      position: sticky; bottom: 0; background: #fff;
      padding: 15px 20px; border-top: 2px solid #e0e0e0;
      display: flex; gap: 10px;
      box-shadow: 0 -4px 12px rgba(0,0,0,0.1);
      z-index: 50; justify-content: flex-end;
    }

    .btn-kaydet {
      padding: 12px 30px;
      background: linear-gradient(135deg, var(--main-bg-color), var(--hover-bg-color));
      color: #fff; border: none; border-radius: 8px;
      font-size: 15px; font-weight: 600; cursor: pointer;
      transition: all 0.3s ease;
      display: flex; align-items: center; justify-content: center; gap: 8px;
      white-space: nowrap;
    }
    .btn-kaydet:hover {
      transform: translateY(-2px);
      box-shadow: 0 6px 15px <% If menu_tipi = "diyet" Then %>rgba(76,175,80,0.4)<% Else %>rgba(69,184,195,0.4)<% End If %>;
    }
    .btn-kaydet svg { width: 18px; height: 18px; }

    .btn-iptal {
      padding: 12px 24px; background: #fff; color: #666;
      border: 2px solid #ddd; border-radius: 8px;
      font-size: 15px; font-weight: 600; cursor: pointer;
      transition: all 0.3s ease; text-decoration: none;
      display: flex; align-items: center; gap: 6px; white-space: nowrap;
    }
    .btn-iptal:hover { border-color: #999; color: #333; }
    .btn-iptal svg { width: 18px; height: 18px; }

    .yemek-tablo-wrapper::-webkit-scrollbar { width: 8px; height: 8px; }
    .yemek-tablo-wrapper::-webkit-scrollbar-track { background: #f1f1f1; }
    .yemek-tablo-wrapper::-webkit-scrollbar-thumb { background: var(--main-bg-color); border-radius: 10px; }

    @media (max-width: 1200px) { .gun-card-body { grid-template-columns: 1fr; } }
    @media (max-width: 768px) {
      .yemek-inputs { grid-template-columns: 1fr; }
      .sticky-kaydet { flex-direction: column; }
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
    <a href="yemek_liste_detay.asp?yil=<%= secilen_yil %>&ay=<%= secilen_ay %><%= tip_param %>" class="nav-button" title="Detaya D&#246;n">
      <svg viewBox="0 0 24 24"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>
    </a>
  </div>
  <h1>
    <svg viewBox="0 0 24 24" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
    (<%= secilen_yil %>) <%= ay_adi %> Ay&#305; <% If menu_tipi = "diyet" Then %>Diyet <% End If %>Men&#252; D&#252;zenleme
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

    <form method="post" action="yemek_duzenle_ay.asp?yil=<%= secilen_yil %>&ay=<%= secilen_ay %><%= tip_param %>">
      <div class="toplu-form-container">
        <div class="toplu-form-header">
          <h2>
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
            <%= ay_adi %>
            <span style="opacity: 0.5; font-weight: 300;">|</span>
            <%= secilen_yil %>
          </h2>
          <div style="display:flex;gap:10px;align-items:center;">
            <% If menu_tipi = "diyet" Then %>
            <div class="type-badge-header">Diyet Men&#252;s&#252;</div>
            <% End If %>
            <div class="gun-count-badge"><%= gun_sayisi_hesap %> G&#252;n</div>
          </div>
        </div>

        <div class="yemek-tablo-wrapper">
          <div class="yemek-grid">
            <%
            rsYemek.MoveFirst
            Do While Not rsYemek.EOF
            %>
            <div class="gun-card">
              <div class="gun-card-header">
                <span class="gun-date"><%= Right("0" & Day(rsYemek("tarih")), 2) %>.<%= Right("0" & Month(rsYemek("tarih")), 2) %>.<%= Year(rsYemek("tarih")) %></span>
                <span class="gun-name"><%= rsYemek("gun_adi") %></span>
              </div>
              <input type="hidden" name="gun_adi_<%= rsYemek("id") %>" value="<%= rsYemek("gun_adi") %>">

              <div class="gun-card-body">
                <div class="yemek-section">
                  <h4>
                    <svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>
                    &#214;&#287;le Yeme&#287;i
                  </h4>
                  <div class="yemek-inputs">
                    <input type="text" name="ogle_corba_<%= rsYemek("id") %>" class="yemek-input" placeholder="&#199;orba" value="<%= rsYemek("ogle_corba") %>">
                    <input type="text" name="ogle_ana_<%= rsYemek("id") %>" class="yemek-input" placeholder="Ana Yemek" value="<%= rsYemek("ogle_ana_yemek") %>">
                    <input type="text" name="ogle_yan_<%= rsYemek("id") %>" class="yemek-input" placeholder="Yan &#220;r&#252;n" value="<%= rsYemek("ogle_yan_urun") %>">
                    <input type="text" name="ogle_tatli_<%= rsYemek("id") %>" class="yemek-input" placeholder="Tatl&#305;/Meyve" value="<%= rsYemek("ogle_tatli") %>">
                  </div>
                </div>

                <div class="yemek-section">
                  <h4>
                    <svg viewBox="0 0 24 24"><path d="M20 8.69V4h-4.69L12 .69 8.69 4H4v4.69L.69 12 4 15.31V20h4.69L12 23.31 15.31 20H20v-4.69L23.31 12 20 8.69zM12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6 6 2.69 6 6-2.69 6-6 6zm0-10c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4-1.79-4-4-4z"/></svg>
                    Ak&#351;am Yeme&#287;i
                  </h4>
                  <div class="yemek-inputs">
                    <input type="text" name="aksam_corba_<%= rsYemek("id") %>" class="yemek-input" placeholder="&#199;orba" value="<%= rsYemek("aksam_corba") %>">
                    <input type="text" name="aksam_ana_<%= rsYemek("id") %>" class="yemek-input" placeholder="Ana Yemek" value="<%= rsYemek("aksam_ana_yemek") %>">
                    <input type="text" name="aksam_yan_<%= rsYemek("id") %>" class="yemek-input" placeholder="Yan &#220;r&#252;n" value="<%= rsYemek("aksam_yan_urun") %>">
                    <input type="text" name="aksam_tatli_<%= rsYemek("id") %>" class="yemek-input" placeholder="Tatl&#305;/Meyve" value="<%= rsYemek("aksam_tatli") %>">
                  </div>
                </div>
              </div>
            </div>
            <%
            rsYemek.MoveNext
            Loop
            %>
          </div>
        </div>

        <div class="sticky-kaydet">
          <button type="submit" name="toplu_guncelle" value="1" class="btn-kaydet">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M17 3H5c-1.11 0-2 .9-2 2v14c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V7l-4-4zm-5 16c-1.66 0-3-1.34-3-3s1.34-3 3-3 3 1.34 3 3-1.34 3-3 3zm3-10H5V5h10v4z"/></svg>
            T&#252;m De&#287;i&#351;iklikleri Kaydet
          </button>
          <a href="yemek_liste_detay.asp?yil=<%= secilen_yil %>&ay=<%= secilen_ay %><%= tip_param %>" class="btn-iptal">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
            &#304;ptal
          </a>
        </div>
      </div>
    </form>

  </div>
</body>
</html>
<%
rsYemek.Close
Set rsYemek = Nothing
%>
