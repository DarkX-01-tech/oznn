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

Function GetDayName(dayNum)
    Select Case dayNum
        Case 1: GetDayName = "PAZAR"
        Case 2: GetDayName = "PAZARTES&#304;"
        Case 3: GetDayName = "SALI"
        Case 4: GetDayName = "&#199;AR&#350;AMBA"
        Case 5: GetDayName = "PER&#350;EMBE"
        Case 6: GetDayName = "CUMA"
        Case 7: GetDayName = "CUMARTES&#304;"
    End Select
End Function

Dim secilen_yil, secilen_ay, ay_adi
Dim gun_listesi()
Dim basarili_mesaj, hata_mesaj
Dim form_gosterilsin
form_gosterilsin = False

'====================================
' Menü Tipi Seçimi
'====================================
Dim menu_tipi, menu_tipi_secildi
menu_tipi_secildi = False

If Request.Form("menu_tipi_sec_btn") <> "" Then
    menu_tipi = Request.Form("menu_tipi")
    If menu_tipi <> "" Then
        menu_tipi_secildi = True
    End If
End If

If Request.Form("ay_sec_btn") <> "" Then
    menu_tipi = Request.Form("menu_tipi")
    menu_tipi_secildi = True
    secilen_yil = Request.Form("yil")
    secilen_ay = Request.Form("ay")
    ay_adi = GetMonthName(CInt(secilen_ay))

    Dim rsKontrol, sqlKontrol
    If menu_tipi = "diyet" Then
        sqlKontrol = "SELECT COUNT(*) AS adet FROM diyet_yemek_listesi WHERE YEAR(tarih) = " & secilen_yil & " AND MONTH(tarih) = " & secilen_ay
    Else
        sqlKontrol = "SELECT COUNT(*) AS adet FROM yemek_listesi WHERE YEAR(tarih) = " & secilen_yil & " AND MONTH(tarih) = " & secilen_ay
    End If
    Set rsKontrol = ConnYemek.Execute(sqlKontrol)
    If rsKontrol("adet") > 0 Then
        hata_mesaj = ay_adi & " " & secilen_yil & " &#304;&#231;in Zaten Kay&#305;t Mevcut!"
    Else
        form_gosterilsin = True
        Dim gun_sayisi, i
        gun_sayisi = Day(DateSerial(secilen_yil, secilen_ay + 1, 0))
        ReDim gun_listesi(gun_sayisi - 1)
        For i = 1 To gun_sayisi
            Dim tarih_obj, gun_no
            tarih_obj = DateSerial(secilen_yil, secilen_ay, i)
            gun_no = Weekday(tarih_obj)
            Set gun_listesi(i - 1) = Server.CreateObject("Scripting.Dictionary")
            gun_listesi(i - 1).Add "gun", i
            gun_listesi(i - 1).Add "tarih", tarih_obj
            gun_listesi(i - 1).Add "gun_adi", GetDayName(gun_no)
        Next
    End If
    rsKontrol.Close
    Set rsKontrol = Nothing
End If

If Request.Form("toplu_kaydet") <> "" Then
    secilen_yil = Request.Form("kayit_yil")
    secilen_ay = Request.Form("kayit_ay")
    menu_tipi = Request.Form("menu_tipi")
    ay_adi = GetMonthName(CInt(secilen_ay))
    Dim gun_sayisi_kayit, hata_sayisi, basarili_sayisi
    gun_sayisi_kayit = Day(DateSerial(secilen_yil, secilen_ay + 1, 0))
    hata_sayisi = 0
    basarili_sayisi = 0
    Dim j, tarih_str, gun_adi_str
    Dim ogle_corba, ogle_ana_yemek, ogle_yan_urun, ogle_tatli
    Dim aksam_corba, aksam_ana_yemek, aksam_yan_urun, aksam_tatli

    Dim hedef_tablo
    If menu_tipi = "diyet" Then
        hedef_tablo = "diyet_yemek_listesi"
    Else
        hedef_tablo = "yemek_listesi"
    End If

    On Error Resume Next
    For j = 1 To gun_sayisi_kayit
        tarih_str = Request.Form("tarih_" & j)
        gun_adi_str = Request.Form("gun_adi_" & j)
        ogle_corba = Trim(Request.Form("ogle_corba_" & j))
        ogle_ana_yemek = Trim(Request.Form("ogle_ana_" & j))
        ogle_yan_urun = Trim(Request.Form("ogle_yan_" & j))
        ogle_tatli = Trim(Request.Form("ogle_tatli_" & j))
        aksam_corba = Trim(Request.Form("aksam_corba_" & j))
        aksam_ana_yemek = Trim(Request.Form("aksam_ana_" & j))
        aksam_yan_urun = Trim(Request.Form("aksam_yan_" & j))
        aksam_tatli = Trim(Request.Form("aksam_tatli_" & j))
        gun_adi_str = Replace(gun_adi_str, "'", "''")
        ogle_corba = Replace(ogle_corba, "'", "''")
        ogle_ana_yemek = Replace(ogle_ana_yemek, "'", "''")
        ogle_yan_urun = Replace(ogle_yan_urun, "'", "''")
        ogle_tatli = Replace(ogle_tatli, "'", "''")
        aksam_corba = Replace(aksam_corba, "'", "''")
        aksam_ana_yemek = Replace(aksam_ana_yemek, "'", "''")
        aksam_yan_urun = Replace(aksam_yan_urun, "'", "''")
        aksam_tatli = Replace(aksam_tatli, "'", "''")
        Dim sqlInsert
        sqlInsert = "INSERT INTO " & hedef_tablo & " (tarih, gun_adi, ogle_corba, ogle_ana_yemek, ogle_yan_urun, ogle_tatli, aksam_corba, aksam_ana_yemek, aksam_yan_urun, aksam_tatli, aktif) VALUES (#" & tarih_str & "#, '" & gun_adi_str & "', '" & ogle_corba & "', '" & ogle_ana_yemek & "', '" & ogle_yan_urun & "', '" & ogle_tatli & "', '" & aksam_corba & "', '" & aksam_ana_yemek & "', '" & aksam_yan_urun & "', '" & aksam_tatli & "', True)"
        ConnYemek.Execute sqlInsert
        If Err.Number = 0 Then
            basarili_sayisi = basarili_sayisi + 1
        Else
            hata_sayisi = hata_sayisi + 1
        End If
        Err.Clear
    Next
    On Error GoTo 0
    If hata_sayisi = 0 Then
        If menu_tipi = "diyet" Then
            Response.Redirect "yemek_liste_detay.asp?yil=" & secilen_yil & "&ay=" & secilen_ay & "&tip=diyet&durum=eklendi"
        Else
            Response.Redirect "yemek_liste_detay.asp?yil=" & secilen_yil & "&ay=" & secilen_ay & "&durum=eklendi"
        End If
    Else
        hata_mesaj = basarili_sayisi & " Kay&#305;t Eklendi, " & hata_sayisi & " Eklenemedi!"
    End If
End If
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Yeni Yemek Listesi Ekle</title>
  <link rel="icon" href="../../images/hastane_portal_logo.png" type="image/png">
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
  <style>
    :root { --main: #45b8c3; --dark: #2e8b91; --txt: #fff; }
    * { margin: 0; padding: 0; box-sizing: border-box; -webkit-user-select: none; -moz-user-select: none; -ms-user-select: none; user-select: none; }
    input, textarea, select { -webkit-user-select: text; -moz-user-select: text; -ms-user-select: text; user-select: text; }
    html, body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      color: #333; min-height: 100vh;
      background: linear-gradient(135deg, #f0f9fa, #e6f4f1);
    }
    #header {
      background: linear-gradient(90deg, var(--main), var(--dark));
      padding: 8px 15px; color: var(--txt);
      display: flex; align-items: center; justify-content: space-between;
      box-shadow: 0 4px 10px rgba(0,0,0,0.15);
      position: sticky; top: 0; z-index: 100;
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
    .header-right { display: flex; align-items: center; gap: 8px; }
    .content { padding: 20px; max-width: 100%; margin: 0 auto; }
    .alert {
      padding: 12px 20px; border-radius: 10px; margin-bottom: 15px;
      font-size: 14px; display: flex; align-items: center; gap: 10px;
      max-width: 1200px; margin-left: auto; margin-right: auto;
      animation: slideDown 0.3s;
    }
    @keyframes slideDown { from { opacity: 0; transform: translateY(-10px); } to { opacity: 1; transform: translateY(0); } }
    .alert svg { width: 20px; height: 20px; flex-shrink: 0; }
    .alert-error { background: #ffebee; color: #c62828; border-left: 4px solid #f44336; }
    .alert-success { background: rgba(69,184,195,0.1); color: var(--dark); border-left: 4px solid var(--main); }

    /* Menu Tipi Secim */
    .menu-tipi-container {
      background: #fff; border-radius: 14px;
      box-shadow: 0 6px 20px rgba(0,0,0,0.06);
      padding: 35px; max-width: 700px; margin: 0 auto;
      overflow: hidden;
    }
    .menu-tipi-header { text-align: center; margin-bottom: 30px; }
    .menu-tipi-header-icon {
      width: 70px; height: 70px; border-radius: 50%; margin: 0 auto 18px;
      background: linear-gradient(135deg, rgba(69,184,195,0.1), rgba(46,139,145,0.05));
      display: flex; align-items: center; justify-content: center;
    }
    .menu-tipi-header-icon svg { width: 36px; height: 36px; fill: var(--main); }
    .menu-tipi-header h2 { font-size: 22px; color: #2c3e50; margin: 0 0 8px 0; }
    .menu-tipi-header p { color: #888; font-size: 13px; }
    .menu-tipi-options { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 25px; }
    .menu-tipi-option {
      border: 2px solid #e8e8e8; border-radius: 12px;
      padding: 24px 18px; text-align: center; cursor: pointer;
      transition: all 0.3s; position: relative;
    }
    .menu-tipi-option:hover { border-color: var(--main); background: rgba(69,184,195,0.04); }
    .menu-tipi-option.selected { border-color: var(--main); background: rgba(69,184,195,0.08); box-shadow: 0 4px 12px rgba(69,184,195,0.2); }
    .menu-tipi-option input[type="radio"] { position: absolute; opacity: 0; pointer-events: none; }
    .menu-tipi-option-icon {
      width: 56px; height: 56px; border-radius: 50%; margin: 0 auto 14px;
      display: flex; align-items: center; justify-content: center;
      transition: all 0.3s;
    }
    .menu-tipi-option-icon svg { width: 28px; height: 28px; }
    .menu-tipi-option.normal .menu-tipi-option-icon { background: rgba(69,184,195,0.1); }
    .menu-tipi-option.normal .menu-tipi-option-icon svg { fill: var(--main); }
    .menu-tipi-option.diyet .menu-tipi-option-icon { background: rgba(76,175,80,0.1); }
    .menu-tipi-option.diyet .menu-tipi-option-icon svg { fill: #4caf50; }
    .menu-tipi-option.selected.normal { border-color: var(--main); }
    .menu-tipi-option.selected.diyet { border-color: #4caf50; background: rgba(76,175,80,0.06); box-shadow: 0 4px 12px rgba(76,175,80,0.2); }
    .menu-tipi-option h3 { font-size: 15px; color: #2c3e50; margin: 0 0 6px 0; }
    .menu-tipi-option p { font-size: 12px; color: #999; margin: 0; }
    .menu-tipi-option .check-indicator {
      position: absolute; top: 10px; right: 10px;
      width: 22px; height: 22px; border-radius: 50%;
      border: 2px solid #e0e0e0; display: flex;
      align-items: center; justify-content: center; transition: all 0.3s;
    }
    .menu-tipi-option.selected .check-indicator {
      border-color: var(--main); background: var(--main);
    }
    .menu-tipi-option.selected.diyet .check-indicator {
      border-color: #4caf50; background: #4caf50;
    }
    .menu-tipi-option .check-indicator svg { width: 12px; height: 12px; fill: #fff; opacity: 0; transition: all 0.3s; }
    .menu-tipi-option.selected .check-indicator svg { opacity: 1; }

    .ay-secim-container {
      background: #fff; border-radius: 14px;
      box-shadow: 0 6px 20px rgba(0,0,0,0.06);
      padding: 35px; max-width: 700px; margin: 0 auto;
      overflow: hidden;
    }
    .ay-secim-header { text-align: center; margin-bottom: 30px; }
    .ay-secim-header-icon {
      width: 70px; height: 70px; border-radius: 50%; margin: 0 auto 18px;
      background: linear-gradient(135deg, rgba(69,184,195,0.1), rgba(46,139,145,0.05));
      display: flex; align-items: center; justify-content: center;
    }
    .ay-secim-header-icon svg { width: 36px; height: 36px; fill: var(--main); }
    .ay-secim-header h2 { font-size: 22px; color: #2c3e50; margin: 0 0 8px 0; }
    .ay-secim-header p { color: #888; font-size: 13px; }
    .selected-type-badge {
      display: inline-flex; align-items: center; gap: 6px;
      padding: 6px 14px; border-radius: 20px; font-size: 12px;
      font-weight: 600; margin-top: 10px;
    }
    .selected-type-badge.normal { background: rgba(69,184,195,0.1); color: var(--dark); }
    .selected-type-badge.diyet { background: rgba(76,175,80,0.1); color: #388e3c; }
    .selected-type-badge svg { width: 14px; height: 14px; }
    .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-bottom: 20px; }
    .form-group label {
      display: block; color: #555; font-size: 12px; font-weight: 600;
      margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.3px;
    }
    .form-group select {
      width: 100%; padding: 12px 14px; border: 2px solid #e8e8e8;
      border-radius: 10px; font-size: 14px; outline: none;
      transition: all 0.3s; background: #fff; color: #333;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .form-group select:focus { border-color: var(--main); box-shadow: 0 0 0 3px rgba(69,184,195,0.15); }
    .btn-devam {
      width: 100%; padding: 14px; border: none; border-radius: 10px;
      font-size: 15px; font-weight: 600; cursor: pointer;
      display: flex; align-items: center; justify-content: center; gap: 8px;
      background: linear-gradient(135deg, var(--main), var(--dark));
      color: #fff; margin-top: 15px; transition: all 0.3s;
    }
    .btn-devam:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(69,184,195,0.35); }
    .btn-devam svg { width: 18px; height: 18px; flex-shrink: 0; }
    .btn-devam:disabled { opacity: 0.5; cursor: not-allowed; transform: none; box-shadow: none; }
    .pdf-upload-section {
      margin-top: 25px; padding-top: 25px;
      border-top: 2px dashed #e8e8e8;
    }
    .pdf-upload-header { text-align: center; margin-bottom: 15px; }
    .pdf-upload-header h3 {
      font-size: 15px; color: var(--dark); margin-bottom: 5px;
      display: flex; align-items: center; justify-content: center; gap: 8px;
    }
    .pdf-upload-header h3 svg { width: 20px; height: 20px; fill: var(--main); }
    .pdf-upload-header p { font-size: 12px; color: #999; }
    .pdf-drop-zone {
      border: 2px dashed var(--main); border-radius: 12px;
      padding: 24px; text-align: center;
      background: rgba(69,184,195,0.04); cursor: pointer;
      position: relative; transition: all 0.3s;
    }
    .pdf-drop-zone:hover { border-color: var(--dark); background: rgba(69,184,195,0.08); }
    .pdf-drop-zone.dragover { border-color: var(--dark); background: rgba(69,184,195,0.12); }
    .pdf-drop-zone.file-selected { border-color: var(--main); background: rgba(69,184,195,0.08); }
    .pdf-drop-zone svg { width: 40px; height: 40px; fill: var(--main); margin-bottom: 8px; }
    .pdf-drop-zone p { font-size: 13px; color: #666; margin: 4px 0; }
    .pdf-drop-zone .file-name { font-weight: 700; color: var(--dark); font-size: 14px; }
    #pdfFileInput { position: absolute; top: 0; left: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; }
    .pdf-status { margin-top: 10px; padding: 10px 14px; border-radius: 8px; font-size: 12px; display: none; }
    .pdf-status.loading { display: block; background: rgba(69,184,195,0.1); color: var(--dark); border-left: 3px solid var(--main); }
    .pdf-status.success { display: block; background: rgba(69,184,195,0.1); color: var(--dark); border-left: 3px solid var(--main); }
    .pdf-status.error { display: block; background: #ffebee; color: #c62828; border-left: 3px solid #f44336; }
    .pdf-preview-wrapper { max-height: 250px; overflow-y: auto; border: 1px solid #e8e8e8; border-radius: 8px; margin-top: 10px; display: none; }
    .pdf-preview-wrapper::-webkit-scrollbar { width: 4px; }
    .pdf-preview-wrapper::-webkit-scrollbar-thumb { background: var(--main); border-radius: 10px; }
    .pdf-preview-table { width: 100%; border-collapse: collapse; font-size: 11px; display: none; }
    .pdf-preview-table th { background: linear-gradient(135deg, var(--main), var(--dark)); color: #fff; padding: 7px 8px; font-size: 10px; text-align: left; position: sticky; top: 0; }
    .pdf-preview-table td { padding: 5px 8px; border-bottom: 1px solid #f0f0f0; }
    .pdf-preview-table tr:nth-child(even) { background: rgba(69,184,195,0.03); }
    .pdf-preview-table .preview-date { font-weight: 700; color: var(--dark); white-space: nowrap; }
    .or-divider { text-align: center; margin: 20px 0 5px 0; position: relative; }
    .or-divider::before { content: ''; position: absolute; top: 50%; left: 0; right: 0; height: 1px; background: #e8e8e8; }
    .or-divider span { background: #fff; padding: 0 15px; color: #bbb; font-size: 12px; font-weight: 600; position: relative; }
    .toplu-form-container {
      background: #fff; border-radius: 14px;
      box-shadow: 0 6px 20px rgba(0,0,0,0.06);
      overflow: hidden; max-width: 1600px; margin: 0 auto;
    }
    .toplu-form-header {
      background: linear-gradient(135deg, var(--main), var(--dark));
      color: #fff; padding: 18px 22px;
      display: flex; align-items: center; justify-content: space-between;
      flex-wrap: wrap; gap: 10px;
    }
    .toplu-form-header h2 { font-size: 18px; margin: 0; display: flex; align-items: center; gap: 10px; }
    .toplu-form-header h2 svg { width: 22px; height: 22px; flex-shrink: 0; }
    .toplu-form-header.diyet-header { background: linear-gradient(135deg, #4caf50, #388e3c); }
    .gun-count-badge {
      background: rgba(255,255,255,0.2); padding: 6px 16px;
      border-radius: 20px; font-size: 13px; font-weight: 600; white-space: nowrap;
    }
    .type-badge-header {
      background: rgba(255,255,255,0.2); padding: 6px 16px;
      border-radius: 20px; font-size: 12px; font-weight: 600; white-space: nowrap;
    }
    .pdf-filled-badge {
      background: rgba(255,255,255,0.2); padding: 6px 16px;
      border-radius: 20px; font-size: 12px; font-weight: 600;
      display: none; align-items: center; gap: 6px; white-space: nowrap;
    }
    .yemek-tablo-wrapper { max-height: calc(100vh - 220px); overflow-y: auto; padding: 18px; }
    .yemek-tablo-wrapper::-webkit-scrollbar { width: 5px; }
    .yemek-tablo-wrapper::-webkit-scrollbar-thumb { background: var(--main); border-radius: 10px; }
    .yemek-grid { display: grid; gap: 12px; }
    .gun-card {
      background: #fff; border: 2px solid #e8e8e8; border-radius: 12px;
      overflow: hidden; transition: all 0.3s;
    }
    .gun-card:hover { border-color: var(--main); box-shadow: 0 6px 18px rgba(69,184,195,0.15); }
    .gun-card.pdf-filled { border-color: var(--main); background: rgba(69,184,195,0.02); }
    .gun-card-header {
      background: linear-gradient(135deg, rgba(69,184,195,0.08), rgba(46,139,145,0.04));
      color: #2c3e50; padding: 10px 16px; font-weight: 600; font-size: 14px;
      display: flex; align-items: center; justify-content: space-between;
      border-bottom: 1px solid #f0f0f0;
    }
    .gun-card.pdf-filled .gun-card-header {
      background: linear-gradient(135deg, rgba(69,184,195,0.15), rgba(46,139,145,0.08));
    }
    .gun-date { font-size: 14px; font-weight: 700; color: var(--dark); }
    .gun-name {
      font-size: 11px; color: var(--dark);
      background: rgba(69,184,195,0.12); padding: 3px 12px;
      border-radius: 10px; font-weight: 600;
    }
    .gun-card-body { padding: 14px 16px; display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
    .yemek-section {
      border: 1px solid #f0f0f0; border-radius: 10px;
      padding: 12px; background: #fafcfc; transition: all 0.2s;
    }
    .yemek-section:hover { background: #f5fffe; border-color: rgba(69,184,195,0.3); }
    .yemek-section h4 {
      font-size: 11px; color: var(--dark); margin-bottom: 10px;
      text-transform: uppercase; font-weight: 700; letter-spacing: 0.3px;
      display: flex; align-items: center; gap: 6px;
      padding-bottom: 8px; border-bottom: 1px solid #f0f0f0;
    }
    .yemek-section h4 svg { width: 14px; height: 14px; fill: var(--main); flex-shrink: 0; }
    .yemek-inputs { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
    .yemek-input {
      width: 100%; padding: 8px 10px; border: 2px solid #e8e8e8;
      border-radius: 8px; font-size: 12px; outline: none;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      transition: all 0.3s;
    }
    .yemek-input:focus { border-color: var(--main); background: #fff; box-shadow: 0 0 0 3px rgba(69,184,195,0.1); }
    .yemek-input::placeholder { color: #bbb; font-style: italic; font-size: 11px; }
    .yemek-input.pdf-filled-input { border-color: var(--main); background: rgba(69,184,195,0.05); }
    .sticky-kaydet {
      position: sticky; bottom: 0; background: #fff;
      padding: 14px 22px; border-top: 2px solid #f0f0f0;
      display: flex; gap: 10px;
      box-shadow: 0 -4px 16px rgba(0,0,0,0.06);
      z-index: 50; justify-content: flex-end;
    }
    .btn-kaydet {
      padding: 12px 28px;
      background: linear-gradient(135deg, var(--main), var(--dark));
      color: #fff; border: none; border-radius: 10px;
      font-size: 14px; font-weight: 600; cursor: pointer;
      display: flex; align-items: center; gap: 8px; transition: all 0.3s;
    }
    .btn-kaydet:hover { transform: translateY(-2px); box-shadow: 0 8px 20px rgba(69,184,195,0.35); }
    .btn-kaydet svg { width: 20px; height: 20px; flex-shrink: 0; }
    .btn-kaydet.diyet-btn { background: linear-gradient(135deg, #4caf50, #388e3c); }
    .btn-kaydet.diyet-btn:hover { box-shadow: 0 8px 20px rgba(76,175,80,0.35); }
    .btn-iptal {
      padding: 12px 22px; background: #fff; color: #666;
      border: 2px solid #e8e8e8; border-radius: 10px;
      font-size: 14px; font-weight: 600; cursor: pointer;
      text-decoration: none; display: flex; align-items: center; gap: 6px;
      transition: all 0.3s;
    }
    .btn-iptal:hover { border-color: var(--main); color: var(--dark); }
    .btn-iptal svg { width: 18px; height: 18px; flex-shrink: 0; }
    @media (max-width: 1200px) { .gun-card-body { grid-template-columns: 1fr; } }
    @media (max-width: 768px) {
      #header { flex-wrap: wrap; gap: 8px; padding: 10px; }
      #header h1 { font-size: 15px; }
      .yemek-inputs { grid-template-columns: 1fr; }
      .form-row { grid-template-columns: 1fr; }
      .sticky-kaydet { flex-direction: column; }
      .ay-secim-container { padding: 20px; }
      .menu-tipi-container { padding: 20px; }
      .menu-tipi-options { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <div id="header">
    <div id="nav-buttons">
      <a href="panel.asp" class="nav-button" title="Ana Sayfa">
        <svg viewBox="0 0 24 24"><path d="M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z"/></svg>
      </a>
      <a href="yemek_liste.asp" class="nav-button" title="Yemek Listesi">
        <svg viewBox="0 0 24 24"><path d="M3 13h2v-2H3v2zm0 4h2v-2H3v2zm0-8h2V7H3v2zm4 4h14v-2H7v2zm0 4h14v-2H7v2zM7 7v2h14V7H7z"/></svg>
      </a>
    </div>
    <h1>
      <svg viewBox="0 0 24 24" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
      Yeni Yemek Listesi Ekle
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

    <% If Not form_gosterilsin And Not menu_tipi_secildi Then %>
    <!-- ADIM 1: Menu Tipi Secimi -->
    <div class="menu-tipi-container">
      <div class="menu-tipi-header">
        <div class="menu-tipi-header-icon">
          <svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>
        </div>
        <h2>Men&#252; Tipi Se&#231;in</h2>
        <p>Eklemek &#304;stedi&#287;iniz Men&#252; T&#252;r&#252;n&#252; Se&#231;in</p>
      </div>

      <form method="post" action="yemek_ekle.asp" id="menuTipiForm">
        <div class="menu-tipi-options">
          <label class="menu-tipi-option normal" onclick="selectMenuType(this, 'normal')">
            <input type="radio" name="menu_tipi" value="normal">
            <div class="check-indicator">
              <svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>
            </div>
            <div class="menu-tipi-option-icon">
              <svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>
            </div>
            <h3>Ayl&#305;k Yemek Listesi</h3>
            <p>Normal men&#252; ekleme</p>
          </label>
          <label class="menu-tipi-option diyet" onclick="selectMenuType(this, 'diyet')">
            <input type="radio" name="menu_tipi" value="diyet">
            <div class="check-indicator">
              <svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>
            </div>
            <div class="menu-tipi-option-icon">
              <svg viewBox="0 0 24 24"><path d="M17.21 9l-4.38-6.56c-.19-.28-.51-.42-.83-.42-.32 0-.64.14-.83.43L6.79 9C6.3 9.71 6 10.57 6 11.5 6 14.53 8.47 17 11.5 17h1c3.03 0 5.5-2.47 5.5-5.5 0-.93-.3-1.79-.79-2.5zM12 15.5c-1.93 0-3.5-1.57-3.5-3.5 0-.53.2-1.33.57-1.88L12 5.97l2.93 4.15c.37.55.57 1.35.57 1.88 0 1.93-1.57 3.5-3.5 3.5zM5.33 20h13.34c.89 0 1.34-1.08.71-1.71L12 11l-7.38 7.29c-.63.63-.18 1.71.71 1.71z"/></svg>
            </div>
            <h3>Ayl&#305;k Diyet Yemek Listesi</h3>
            <p>Diyet men&#252;s&#252; ekleme</p>
          </label>
        </div>

        <button type="submit" name="menu_tipi_sec_btn" value="1" class="btn-devam" id="menuTipiBtn" disabled>
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 4l-1.41 1.41L16.17 11H4v2h12.17l-5.58 5.59L12 20l8-8z"/></svg>
          Devam Et
        </button>
      </form>
    </div>

    <script>
      function selectMenuType(el, type) {
        document.querySelectorAll('.menu-tipi-option').forEach(function(opt) { opt.classList.remove('selected'); });
        el.classList.add('selected');
        el.querySelector('input[type="radio"]').checked = true;
        document.getElementById('menuTipiBtn').disabled = false;
      }
    </script>

    <% ElseIf Not form_gosterilsin And menu_tipi_secildi Then %>
    <!-- ADIM 2: Ay Secimi -->
    <div class="ay-secim-container">
      <div class="ay-secim-header">
        <div class="ay-secim-header-icon">
          <svg viewBox="0 0 24 24"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10z"/></svg>
        </div>
        <h2>Hangi Ay &#304;&#231;in Men&#252; Eklenecek?</h2>
        <p>&#214;nce Ay Ve Y&#305;l Se&#231;in, Sistem Otomatik Olarak T&#252;m G&#252;nleri Olu&#351;turacak</p>
        <% If menu_tipi = "diyet" Then %>
        <span class="selected-type-badge diyet">
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M17.21 9l-4.38-6.56c-.19-.28-.51-.42-.83-.42-.32 0-.64.14-.83.43L6.79 9C6.3 9.71 6 10.57 6 11.5 6 14.53 8.47 17 11.5 17h1c3.03 0 5.5-2.47 5.5-5.5 0-.93-.3-1.79-.79-2.5zM12 15.5c-1.93 0-3.5-1.57-3.5-3.5 0-.53.2-1.33.57-1.88L12 5.97l2.93 4.15c.37.55.57 1.35.57 1.88 0 1.93-1.57 3.5-3.5 3.5z"/></svg>
          Diyet Men&#252;s&#252;
        </span>
        <% Else %>
        <span class="selected-type-badge normal">
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>
          Normal Men&#252;
        </span>
        <% End If %>
      </div>

      <form method="post" action="yemek_ekle.asp" id="aySecimForm">
        <input type="hidden" name="menu_tipi" value="<%= menu_tipi %>">
        <div class="form-row">
          <div class="form-group">
            <label>Ay Se&#231;in</label>
            <select name="ay" required>
              <option value="">Ay Se&#231;in...</option>
              <% For i = 1 To 12 %><option value="<%= i %>"><%= GetMonthName(i) %></option><% Next %>
            </select>
          </div>
          <div class="form-group">
            <label>Y&#305;l Se&#231;in</label>
            <select name="yil" required>
              <option value="">Y&#305;l Se&#231;in...</option>
              <% Dim baslangic_yil, bitis_yil, y: baslangic_yil = 2026: bitis_yil = Year(Now()) + 2
              For y = baslangic_yil To bitis_yil %><option value="<%= y %>"><%= y %></option><% Next %>
            </select>
          </div>
        </div>

        <div class="pdf-upload-section">
          <div class="pdf-upload-header">
            <h3>
              <svg viewBox="0 0 24 24"><path d="M20 2H8c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm-8.5 7.5c0 .83-.67 1.5-1.5 1.5H9v2H7.5V7H10c.83 0 1.5.67 1.5 1.5v1zm5 2c0 .83-.67 1.5-1.5 1.5h-2.5V7H15c.83 0 1.5.67 1.5 1.5v3zm4-3H19v1h1.5V11H19v2h-1.5V7h3v1.5zM9 9.5h1v-1H9v1zM4 6H2v14c0 1.1.9 2 2 2h14v-2H4V6zm10 5.5h1v-3h-1v3z"/></svg>
              PDF Otomatik Y&#252;kle (&#304;ste&#287;e Ba&#287;l&#305;)
            </h3>
            <p>Ayl&#305;k Yemek Listesi PDF Dosyan&#305;z&#305; Se&#231;in, Form Otomatik Doldurulacak</p>
          </div>

          <div class="pdf-drop-zone" id="pdfDropZone">
            <input type="file" id="pdfFileInput" accept=".pdf">
            <svg viewBox="0 0 24 24"><path d="M20 2H8c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm-8.5 7.5c0 .83-.67 1.5-1.5 1.5H9v2H7.5V7H10c.83 0 1.5.67 1.5 1.5v1zm5 2c0 .83-.67 1.5-1.5 1.5h-2.5V7H15c.83 0 1.5.67 1.5 1.5v3zm4-3H19v1h1.5V11H19v2h-1.5V7h3v1.5zM9 9.5h1v-1H9v1zM4 6H2v14c0 1.1.9 2 2 2h14v-2H4V6zm10 5.5h1v-3h-1v3z"/></svg>
            <p id="pdfDropText"><strong>PDF Dosyas&#305;n&#305; Buraya S&#252;r&#252;kleyin</strong> T&#305;klayarak Se&#231;in</p>
            <p id="pdfFileName" class="file-name" style="display:none;"></p>
          </div>

          <div class="pdf-status" id="pdfStatus"></div>
          <div class="pdf-preview-wrapper" id="pdfPreviewWrapper">
            <table class="pdf-preview-table" id="pdfPreviewTable">
              <thead><tr><th>Tarih</th><th>&#214;.&#199;orba</th><th>&#214;.Ana</th><th>&#214;.Yan</th><th>&#214;.Tatl&#305;</th><th>A.&#199;orba</th><th>A.Ana</th><th>A.Yan</th><th>A.Tatl&#305;</th></tr></thead>
              <tbody id="pdfPreviewBody"></tbody>
            </table>
          </div>
        </div>

        <div class="or-divider"><span>PDF Se&#231;meden Devam Edebilirsiniz</span></div>

        <button type="submit" name="ay_sec_btn" value="1" class="btn-devam">
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>
          Devam Et
        </button>
      </form>
    </div>

    <script>
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    var parsedPdfData = null;
    var dropZone = document.getElementById('pdfDropZone');
    var fileInput = document.getElementById('pdfFileInput');
    dropZone.addEventListener('dragover', function(e) { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', function(e) { e.preventDefault(); dropZone.classList.remove('dragover'); });
    dropZone.addEventListener('drop', function(e) { e.preventDefault(); dropZone.classList.remove('dragover'); if (e.dataTransfer.files.length > 0) { fileInput.files = e.dataTransfer.files; handleFileSelect(e.dataTransfer.files[0]); } });
    fileInput.addEventListener('change', function() { if (this.files.length > 0) handleFileSelect(this.files[0]); });

    function handleFileSelect(file) {
      if (file.type !== 'application/pdf') { showStatus('Sadece PDF!', 'error'); return; }
      document.getElementById('pdfDropText').style.display = 'none';
      document.getElementById('pdfFileName').style.display = 'block';
      document.getElementById('pdfFileName').textContent = file.name + ' (' + (file.size / 1024).toFixed(0) + ' KB)';
      dropZone.classList.add('file-selected');
      showStatus('PDF Okunuyor...', 'loading');
      parsePdfFile(file);
    }
    function showStatus(msg, type) { var el = document.getElementById('pdfStatus'); el.className = 'pdf-status ' + type; el.innerHTML = msg; }

    async function parsePdfFile(file) {
      try {
        var arrayBuffer = await file.arrayBuffer();
        var pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        var allItems = [];
        for (var p = 1; p <= pdf.numPages; p++) {
          var page = await pdf.getPage(p);
          var tc = await page.getTextContent();
          tc.items.forEach(function(item) {
            var text = item.str.trim();
            if (text.length > 0) {
              allItems.push({
                text: text,
                x: Math.round(item.transform[4]),
                y: Math.round(item.transform[5]),
                page: p
              });
            }
          });
        }
        var menuData = parseAllItems(allItems);
        if (menuData.length === 0) { showStatus('PDF Veri Okunamad\u0131.', 'error'); return; }
        parsedPdfData = menuData;
        showStatus(menuData.length + ' G\u00fcnl\u00fck Men\u00fc Okundu! \u00d6nizlemeyi Kontrol Edin.', 'success');
        showPreview(menuData);
      } catch (err) { showStatus('PDF Okuma Hatas\u0131: ' + err.message, 'error'); }
    }

    function parseAllItems(allItems) {
      var dateRegex = /^(\d{2})\.(\d{2})\.(\d{4})$/;
      var gunAdlari = ['PAZAR','PAZARTES\u0130','PAZARTESI','SALI','\u00c7AR\u015eAMBA','CARSAMBA','\u00c7ARSAMBA','PER\u015eEMBE','PERSEMBE','CUMA','CUMARTES\u0130','CUMARTESI'];
      var skipWords = ['TAR\u0130H','TARIH','\u00d6\u011eLE','OGLE','\u00d6GLE','AK\u015eAM','AKSAM','YEME\u011e\u0130','YEMEGI','YILI','AYI','NORMAL','L\u0130STES\u0130','LISTESI','AF\u0130YET','AFIYET','OLSUN','OCAK','\u015eUBAT','SUBAT','MART','N\u0130SAN','NISAN','MAYIS','HAZ\u0130RAN','HAZIRAN','TEMMUZ','A\u011eUSTOS','AGUSTOS','EYL\u00dcL','EYLUL','EK\u0130M','EKIM','KASIM','ARALIK','D\u0130YET','DIYET'];

      function isSkip(text) {
        var t = text.toUpperCase().trim();
        if (t.length === 0) return true;
        for (var i = 0; i < gunAdlari.length; i++) { if (t === gunAdlari[i]) return true; }
        for (var i = 0; i < skipWords.length; i++) { if (t.indexOf(skipWords[i]) >= 0) return true; }
        if (/^\d{4}$/.test(t)) return true;
        return false;
      }

      function isDate(text) { return dateRegex.test(text.trim()); }

      allItems.sort(function(a, b) {
        if (a.page !== b.page) return a.page - b.page;
        if (Math.abs(a.y - b.y) > 4) return b.y - a.y;
        return a.x - b.x;
      });

      var rows = [];
      if (allItems.length === 0) return [];
      var currentRow = [allItems[0]];
      for (var i = 1; i < allItems.length; i++) {
        var item = allItems[i];
        var prev = currentRow[0];
        if (item.page === prev.page && Math.abs(item.y - prev.y) <= 4) {
          currentRow.push(item);
        } else {
          rows.push(currentRow);
          currentRow = [item];
        }
      }
      rows.push(currentRow);

      var dateXValues = [];
      rows.forEach(function(row) {
        row.forEach(function(item) {
          if (isDate(item.text)) dateXValues.push(item.x);
        });
      });
      if (dateXValues.length === 0) return [];

      var dateXMin = Math.min.apply(null, dateXValues);
      var dateXMax = Math.max.apply(null, dateXValues);
      var dateColRight = dateXMax + 60;

      var foodXValues = [];
      rows.forEach(function(row) {
        row.forEach(function(item) {
          if (!isDate(item.text) && !isSkip(item.text) && item.x > dateColRight) {
            foodXValues.push(item.x);
          }
        });
      });

      if (foodXValues.length === 0) return [];
      foodXValues.sort(function(a, b) { return a - b; });

      var maxGap = 0, splitIdx = 0;
      for (var i = 1; i < foodXValues.length; i++) {
        var gap = foodXValues[i] - foodXValues[i - 1];
        if (gap > maxGap) { maxGap = gap; splitIdx = i; }
      }
      var splitX = Math.round((foodXValues[splitIdx - 1] + foodXValues[splitIdx]) / 2);

      var processedRows = [];

      rows.forEach(function(row) {
        row.sort(function(a, b) { return a.x - b.x; });

        var foundDate = null;
        for (var r = 0; r < row.length; r++) {
          if (isDate(row[r].text)) { foundDate = row[r].text; break; }
        }

        var ogleTexts = [];
        var aksamTexts = [];

        row.forEach(function(item) {
          if (isDate(item.text)) return;
          if (isSkip(item.text)) return;
          if (item.x <= dateColRight) return;

          if (item.x < splitX) {
            ogleTexts.push(item.text);
          } else {
            aksamTexts.push(item.text);
          }
        });

        var ogleCell = joinCellParts(ogleTexts);
        var aksamCell = joinCellParts(aksamTexts);

        if (foundDate) {
          processedRows.push({ type: 'date', tarih: foundDate, ogle: ogleCell, aksam: aksamCell });
        } else if (ogleCell || aksamCell) {
          processedRows.push({ type: 'food', ogle: ogleCell, aksam: aksamCell });
        }
      });

      var dayGroups = [];
      var currentDay = null;

      processedRows.forEach(function(row) {
        if (row.type === 'date') {
          currentDay = { tarih: row.tarih, ogleList: [], aksamList: [] };
          dayGroups.push(currentDay);
          if (row.ogle) currentDay.ogleList.push(row.ogle);
          if (row.aksam) currentDay.aksamList.push(row.aksam);
        } else if (row.type === 'food' && currentDay) {
          if (row.ogle) currentDay.ogleList.push(row.ogle);
          if (row.aksam) currentDay.aksamList.push(row.aksam);
        }
      });

      var menuData = dayGroups.map(function(day) {
        var o = day.ogleList;
        var a = day.aksamList;
        while (o.length < 4) o.push('');
        while (a.length < 4) a.push('');

        return {
          tarih: day.tarih,
          gun_adi: '',
          ogle_corba: o[0],
          ogle_ana_yemek: o[1],
          ogle_yan_urun: o[2],
          ogle_tatli: o[3],
          aksam_corba: a[0],
          aksam_ana_yemek: a[1],
          aksam_yan_urun: a[2],
          aksam_tatli: a[3]
        };
      });

      menuData.sort(function(a, b) {
        var pA = a.tarih.split('.'), pB = b.tarih.split('.');
        return new Date(pA[2], pA[1] - 1, pA[0]) - new Date(pB[2], pB[1] - 1, pB[0]);
      });

      return menuData;
    }

    function joinCellParts(parts) {
      if (parts.length === 0) return '';
      if (parts.length === 1) return parts[0].trim();
      return parts.join(' ').trim();
    }

    function showPreview(menuData) {
      var tbody = document.getElementById('pdfPreviewBody'); tbody.innerHTML = '';
      menuData.forEach(function(item) {
        var tr = document.createElement('tr');
        tr.innerHTML = '<td class="preview-date">' + item.tarih + '</td><td>' + item.ogle_corba + '</td><td>' + item.ogle_ana_yemek + '</td><td>' + item.ogle_yan_urun + '</td><td>' + item.ogle_tatli + '</td><td>' + item.aksam_corba + '</td><td>' + item.aksam_ana_yemek + '</td><td>' + item.aksam_yan_urun + '</td><td>' + item.aksam_tatli + '</td>';
        tbody.appendChild(tr);
      });
      document.getElementById('pdfPreviewWrapper').style.display = 'block';
      document.getElementById('pdfPreviewTable').style.display = 'table';
    }

    document.getElementById('aySecimForm').addEventListener('submit', function() {
      if (parsedPdfData && parsedPdfData.length > 0) sessionStorage.setItem('pdfMenuData', JSON.stringify(parsedPdfData));
    });
    </script>

    <% Else %>
    <!-- ADIM 3: Toplu Form -->
    <script>
      function validateForm() {
        var gs = <%= UBound(gun_listesi) + 1 %>, ba = [];
        for (var i = 1; i <= gs; i++) {
          var bl = [];
          if (!document.querySelector('input[name="ogle_corba_'+i+'"]').value.trim()) bl.push(1);
          if (!document.querySelector('input[name="ogle_ana_'+i+'"]').value.trim()) bl.push(2);
          if (!document.querySelector('input[name="ogle_yan_'+i+'"]').value.trim()) bl.push(3);
          if (!document.querySelector('input[name="ogle_tatli_'+i+'"]').value.trim()) bl.push(4);
          if (!document.querySelector('input[name="aksam_corba_'+i+'"]').value.trim()) bl.push(5);
          if (!document.querySelector('input[name="aksam_ana_'+i+'"]').value.trim()) bl.push(6);
          if (!document.querySelector('input[name="aksam_yan_'+i+'"]').value.trim()) bl.push(7);
          if (!document.querySelector('input[name="aksam_tatli_'+i+'"]').value.trim()) bl.push(8);
          if (bl.length > 0) ba.push(i);
        }
        if (ba.length > 0) { alert('T\u00fcm Alanlar\u0131 Eksiksiz Doldurun!'); var el = document.querySelector('input[name="ogle_corba_'+ba[0]+'"]'); if (el) { el.scrollIntoView({behavior:'smooth',block:'center'}); el.focus(); } return false; }
        return confirm(gs + ' G\u00fcnl\u00fck Men\u00fc Kaydedilecek. Devam Edilsin Mi?');
      }
      function fillFormFromPdf() {
        var s = sessionStorage.getItem('pdfMenuData'); if (!s) return;
        var md = JSON.parse(s); if (!md || !md.length) return;
        var sy='<%= secilen_yil %>', sa='<%= secilen_ay %>', gs=<%= UBound(gun_listesi) + 1 %>, dg=0;
        for (var g = 1; g <= gs; g++) {
          var ds = (g<10?'0':'')+g, ms = (parseInt(sa)<10?'0':'')+parseInt(sa), td = ds+'.'+ms+'.'+sy, f = null;
          for (var m = 0; m < md.length; m++) { if (md[m].tarih === td) { f = md[m]; break; } }
          if (f) {
            var fl = [{n:'ogle_corba_'+g,v:f.ogle_corba},{n:'ogle_ana_'+g,v:f.ogle_ana_yemek},{n:'ogle_yan_'+g,v:f.ogle_yan_urun},{n:'ogle_tatli_'+g,v:f.ogle_tatli},{n:'aksam_corba_'+g,v:f.aksam_corba},{n:'aksam_ana_'+g,v:f.aksam_ana_yemek},{n:'aksam_yan_'+g,v:f.aksam_yan_urun},{n:'aksam_tatli_'+g,v:f.aksam_tatli}];
            var gd = false;
            fl.forEach(function(ff) { var inp = document.querySelector('input[name="'+ff.n+'"]'); if (inp && ff.v) { inp.value = ff.v; inp.classList.add('pdf-filled-input'); gd = true; } });
            if (gd) { dg++; var ci = document.querySelector('input[name="ogle_corba_'+g+'"]'); if(ci){var gc=ci.closest('.gun-card');if(gc)gc.classList.add('pdf-filled');} }
          }
        }
        sessionStorage.removeItem('pdfMenuData');
        if (dg > 0) {
          var badge = document.getElementById('pdfFilledBadge');
          if (badge) { badge.style.display='flex'; badge.textContent='PDF '+dg+' G\u00fcn Dolduruldu'; }
          var ad = document.createElement('div'); ad.className='alert alert-success';
          ad.innerHTML='PDF <strong>'+dg+' G\u00fcn</strong> Otomatik Dolduruldu.';
          document.querySelector('.content').insertBefore(ad, document.querySelector('.content').children[0]);
        }
      }
      window.addEventListener('DOMContentLoaded', fillFormFromPdf);
    </script>

    <form method="post" action="yemek_ekle.asp" onsubmit="return validateForm()">
      <input type="hidden" name="kayit_yil" value="<%= secilen_yil %>">
      <input type="hidden" name="kayit_ay" value="<%= secilen_ay %>">
      <input type="hidden" name="menu_tipi" value="<%= menu_tipi %>">
      <div class="toplu-form-container">
        <div class="toplu-form-header<% If menu_tipi = "diyet" Then %> diyet-header<% End If %>">
          <h2>
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10z"/></svg>
            <%= ay_adi %> <%= secilen_yil %>
          </h2>
          <div style="display:flex;gap:10px;align-items:center;">
            <div class="pdf-filled-badge" id="pdfFilledBadge"></div>
            <% If menu_tipi = "diyet" Then %>
            <div class="type-badge-header">Diyet Men&#252;s&#252;</div>
            <% Else %>
            <div class="type-badge-header">Normal Men&#252;</div>
            <% End If %>
            <div class="gun-count-badge"><%= UBound(gun_listesi) + 1 %> G&#252;n</div>
          </div>
        </div>

        <div class="yemek-tablo-wrapper">
          <div class="yemek-grid">
            <% Dim k: For k = 0 To UBound(gun_listesi)
              Dim gun_obj: Set gun_obj = gun_listesi(k)
              Dim gun_index: gun_index = gun_obj("gun") %>
            <div class="gun-card">
              <div class="gun-card-header">
                <span class="gun-date"><%= Right("0" & gun_obj("gun"), 2) %>.<%= Right("0" & secilen_ay, 2) %>.<%= secilen_yil %></span>
                <span class="gun-name"><%= gun_obj("gun_adi") %></span>
              </div>
              <input type="hidden" name="tarih_<%= gun_index %>" value="<%= Month(gun_obj("tarih")) %>/<%= Day(gun_obj("tarih")) %>/<%= Year(gun_obj("tarih")) %>">
              <input type="hidden" name="gun_adi_<%= gun_index %>" value="<%= gun_obj("gun_adi") %>">
              <div class="gun-card-body">
                <div class="yemek-section">
                  <h4><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg> &#214;&#287;le Yeme&#287;i</h4>
                  <div class="yemek-inputs">
                    <input type="text" name="ogle_corba_<%= gun_index %>" class="yemek-input" placeholder="&#199;orba">
                    <input type="text" name="ogle_ana_<%= gun_index %>" class="yemek-input" placeholder="Ana Yemek">
                    <input type="text" name="ogle_yan_<%= gun_index %>" class="yemek-input" placeholder="Yan &#220;r&#252;n">
                    <input type="text" name="ogle_tatli_<%= gun_index %>" class="yemek-input" placeholder="Tatl&#305;/Meyve">
                  </div>
                </div>
                <div class="yemek-section">
                  <h4><svg viewBox="0 0 24 24"><path d="M20 8.69V4h-4.69L12 .69 8.69 4H4v4.69L.69 12 4 15.31V20h4.69L12 23.31 15.31 20H20v-4.69L23.31 12 20 8.69zM12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6 6 2.69 6 6-2.69 6-6 6zm0-10c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4-1.79-4-4-4z"/></svg> Ak&#351;am Yeme&#287;i</h4>
                  <div class="yemek-inputs">
                    <input type="text" name="aksam_corba_<%= gun_index %>" class="yemek-input" placeholder="&#199;orba">
                    <input type="text" name="aksam_ana_<%= gun_index %>" class="yemek-input" placeholder="Ana Yemek">
                    <input type="text" name="aksam_yan_<%= gun_index %>" class="yemek-input" placeholder="Yan &#220;r&#252;n">
                    <input type="text" name="aksam_tatli_<%= gun_index %>" class="yemek-input" placeholder="Tatl&#305;/Meyve">
                  </div>
                </div>
              </div>
            </div>
            <% Next %>
          </div>
        </div>

        <div class="sticky-kaydet">
          <button type="submit" name="toplu_kaydet" value="1" class="btn-kaydet<% If menu_tipi = "diyet" Then %> diyet-btn<% End If %>">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M17 3H5c-1.11 0-2 .9-2 2v14c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V7l-4-4zm-5 16c-1.66 0-3-1.34-3-3s1.34-3 3-3 3 1.34 3 3-1.34 3-3 3zm3-10H5V5h10v4z"/></svg>
            T&#252;m&#252;n&#252; Kaydet (<%= UBound(gun_listesi) + 1 %> G&#252;n)
          </button>
          <a href="yemek_ekle.asp" class="btn-iptal">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
            &#304;ptal
          </a>
        </div>
      </div>
    </form>
    <% End If %>
  </div>
</body>
</html>
