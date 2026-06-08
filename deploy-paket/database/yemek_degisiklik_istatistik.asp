<%
' Gerekli degiskenler: secilen_yil, secilen_ay, menu_tipi, ay_adi
Dim rsDegisiklik, degisiklikVar, degisiklikHata
degisiklikVar = False
degisiklikHata = False
Dim degisiklikSayisi, degisenMenuGun, ogleDegisim, aksamDegisim
degisiklikSayisi = 0
degisenMenuGun = 0
ogleDegisim = 0
aksamDegisim = 0

On Error Resume Next
Set rsDegisiklik = ConnYemek.Execute("SELECT * FROM yemek_degisiklik_log WHERE yil = " & secilen_yil & " AND ay = " & secilen_ay & " AND menu_tipi = '" & menu_tipi & "' ORDER BY degisiklik_tarihi DESC")
If Err.Number <> 0 Then
    degisiklikHata = True
Else
    degisiklikVar = True
    If Not rsDegisiklik.EOF Then
        Dim rsDegOzet
        Set rsDegOzet = ConnYemek.Execute("SELECT COUNT(*) AS toplam, COUNT(DISTINCT menu_tarih) AS gun FROM yemek_degisiklik_log WHERE yil = " & secilen_yil & " AND ay = " & secilen_ay & " AND menu_tipi = '" & menu_tipi & "'")
        If Not rsDegOzet.EOF Then
            degisiklikSayisi = rsDegOzet("toplam")
            degisenMenuGun = rsDegOzet("gun")
        End If
        rsDegOzet.Close: Set rsDegOzet = Nothing

        Dim rsOgleAksam
        Set rsOgleAksam = ConnYemek.Execute("SELECT ogun_tipi FROM yemek_degisiklik_log WHERE yil = " & secilen_yil & " AND ay = " & secilen_ay & " AND menu_tipi = '" & menu_tipi & "'")
        Do While Not rsOgleAksam.EOF
            If OgunTipiOgleMi(rsOgleAksam("ogun_tipi")) Then
                ogleDegisim = ogleDegisim + 1
            Else
                aksamDegisim = aksamDegisim + 1
            End If
            rsOgleAksam.MoveNext
        Loop
        rsOgleAksam.Close: Set rsOgleAksam = Nothing
    End If
End If
Err.Clear

Dim rsGuncelleme, guncellemeVar, guncellemeHata, guncellemeSayisi
guncellemeVar = False
guncellemeHata = False
guncellemeSayisi = 0
Dim guncellemeFiltre
If menu_tipi = "diyet" Then
    guncellemeFiltre = " AND (guncelleme_tipi LIKE '%Diyet%' OR aciklama LIKE '%Diyet%')"
Else
    guncellemeFiltre = " AND ((guncelleme_tipi NOT LIKE '%Diyet%' OR guncelleme_tipi IS NULL) AND (aciklama NOT LIKE '%Diyet%' OR aciklama IS NULL))"
End If

Set rsGuncelleme = ConnYemek.Execute("SELECT * FROM yemek_guncelleme_log WHERE yil = " & secilen_yil & " AND ay = " & secilen_ay & guncellemeFiltre & " ORDER BY guncelleme_tarihi DESC")
If Err.Number <> 0 Then
    guncellemeHata = True
Else
    guncellemeVar = True
    Dim rsGuncSay
    Set rsGuncSay = ConnYemek.Execute("SELECT COUNT(*) AS adet FROM yemek_guncelleme_log WHERE yil = " & secilen_yil & " AND ay = " & secilen_ay & guncellemeFiltre)
    If Not rsGuncSay.EOF Then guncellemeSayisi = rsGuncSay("adet")
    rsGuncSay.Close: Set rsGuncSay = Nothing
End If
Err.Clear
On Error GoTo 0
%>
<style>
  .degisiklik-section { margin-bottom: 25px; }
  .degisiklik-card { background: #fff; border-radius: 12px; box-shadow: 0 4px 14px rgba(0,0,0,0.05); overflow: hidden; }
  .degisiklik-card-header { background: linear-gradient(135deg, #ff9800, #e65100); color: #fff; padding: 14px 18px; display: flex; align-items: center; gap: 8px; }
  .degisiklik-card-header h3 { font-size: 14px; margin: 0; font-weight: 600; }
  .degisiklik-card-header svg { width: 18px; height: 18px; fill: currentColor; flex-shrink: 0; }
  .degisiklik-ozet { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; padding: 16px 18px; border-bottom: 1px solid #f0f0f0; }
  .degisiklik-ozet-item { background: #fff8f0; border: 1px solid #ffe0b2; border-radius: 10px; padding: 12px; text-align: center; }
  .degisiklik-ozet-item .do-sayi { font-size: 20px; font-weight: 700; color: #e65100; }
  .degisiklik-ozet-item .do-label { font-size: 10px; color: #999; text-transform: uppercase; margin-top: 4px; }
  .degisiklik-oturumlar { padding: 14px 18px; border-bottom: 1px solid #f0f0f0; }
  .degisiklik-oturumlar h4 { font-size: 12px; color: #666; margin: 0 0 10px 0; text-transform: uppercase; letter-spacing: 0.3px; }
  .oturum-list { display: flex; flex-wrap: wrap; gap: 8px; }
  .oturum-badge { background: rgba(255,152,0,0.12); color: #e65100; padding: 5px 12px; border-radius: 16px; font-size: 11px; font-weight: 600; }
  .degisiklik-tablo-wrap { max-height: 420px; overflow: auto; }
  .degisiklik-tablo-wrap::-webkit-scrollbar { width: 5px; }
  .degisiklik-tablo-wrap::-webkit-scrollbar-thumb { background: #ff9800; border-radius: 10px; }
  .degisiklik-tablo { width: 100%; border-collapse: collapse; font-size: 12px; }
  .degisiklik-tablo thead th { background: #fff3e0; color: #bf360c; padding: 10px 12px; text-align: left; font-size: 10px; text-transform: uppercase; letter-spacing: 0.3px; position: sticky; top: 0; z-index: 1; }
  .degisiklik-tablo tbody td { padding: 9px 12px; border-bottom: 1px solid #f5f5f5; vertical-align: top; }
  .degisiklik-tablo tbody tr:hover { background: #fffaf5; }
  .degisiklik-tablo .eski-val { color: #c62828; text-decoration: line-through; }
  .degisiklik-tablo .yeni-val { color: #2e7d32; font-weight: 600; }
  .degisiklik-empty { text-align: center; padding: 30px 18px; color: #bbb; font-size: 12px; font-style: italic; }
  .degisiklik-warn { text-align: center; padding: 24px 18px; color: #e65100; font-size: 12px; }
  .ogun-badge { display: inline-block; padding: 3px 10px; border-radius: 8px; font-size: 10px; font-weight: 600; margin-right: 6px; }
  .ogun-badge.ogle { background: rgba(69,184,195,0.15); color: #2e8b91; }
  .ogun-badge.aksam { background: rgba(156,39,176,0.12); color: #7b1fa2; }
  @media (max-width: 1000px) { .degisiklik-ozet { grid-template-columns: repeat(2, 1fr); } }
</style>

<div class="degisiklik-section">
  <div class="degisiklik-card">
    <div class="degisiklik-card-header">
      <svg viewBox="0 0 24 24"><path d="M13 3c-4.97 0-9 4.03-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42C8.27 19.99 10.51 21 13 21c4.97 0 9-4.03 9-9s-4.03-9-9-9zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>
      <h3><%= ay_adi %> <%= secilen_yil %> Men&#252; De&#287;i&#351;iklik &#304;statistikleri</h3>
    </div>

    <% If degisiklikHata And guncellemeHata Then %>
    <div class="degisiklik-warn">Log tablolar&#305;na eri&#351;ilemiyor. Access&#39;te <strong>yemek_degisiklik_log</strong> ve <strong>yemek_guncelleme_log</strong> tablolar&#305;n&#305; kontrol edin.</div>
    <% Else %>
    <div class="degisiklik-ozet">
      <div class="degisiklik-ozet-item"><div class="do-sayi"><%= guncellemeSayisi %></div><div class="do-label">Men&#252; G&#252;ncelleme</div></div>
      <div class="degisiklik-ozet-item"><div class="do-sayi"><%= degisiklikSayisi %></div><div class="do-label">&#214;&#287;&#252;n De&#287;i&#351;ikli&#287;i</div></div>
      <div class="degisiklik-ozet-item"><div class="do-sayi"><%= degisenMenuGun %></div><div class="do-label">De&#287;i&#351;en G&#252;n</div></div>
      <div class="degisiklik-ozet-item"><div class="do-sayi"><%= ogleDegisim %> / <%= aksamDegisim %></div><div class="do-label">&#214;&#287;le / Ak&#351;am</div></div>
    </div>

    <% If guncellemeVar And Not rsGuncelleme.EOF Then %>
    <div class="degisiklik-oturumlar">
      <h4>Men&#252;n&#252;n De&#287;i&#351;tirildi&#287;i Tarihler (<%= guncellemeSayisi %> kez)</h4>
      <div class="oturum-list">
        <% Do While Not rsGuncelleme.EOF %>
        <span class="oturum-badge"><%= FormatDateTime(rsGuncelleme("guncelleme_tarihi"), 0) %> &mdash; <%= GuncellemeTipiGoster(rsGuncelleme) %><% If LogSafeStr(rsGuncelleme("kullanici")) <> "" And LogSafeStr(rsGuncelleme("kullanici")) <> GuncellemeTipiGoster(rsGuncelleme) Then %> (<%= rsGuncelleme("kullanici") %>)<% End If %></span>
        <% rsGuncelleme.MoveNext
        Loop
        rsGuncelleme.MoveFirst %>
      </div>
    </div>
    <% End If %>

    <div class="degisiklik-tablo-wrap">
      <% If degisiklikVar And Not rsDegisiklik.EOF Then %>
      <table class="degisiklik-tablo">
        <thead>
          <tr>
            <th>De&#287;i&#351;iklik Tarihi/Saati</th>
            <th>Men&#252; G&#252;n&#252;</th>
            <th>&#214;&#287;&#252;n</th>
            <th>Eski &#214;&#287;&#252;n</th>
            <th>Yeni &#214;&#287;&#252;n</th>
            <th>Kullan&#305;c&#305;</th>
          </tr>
        </thead>
        <tbody>
          <% Do While Not rsDegisiklik.EOF %>
          <tr>
            <td><%= FormatDateTime(rsDegisiklik("degisiklik_tarihi"), 0) %></td>
            <td><%= FormatDateTime(rsDegisiklik("menu_tarih"), 2) %></td>
            <td>
              <% If OgunTipiOgleMi(rsDegisiklik("ogun_tipi")) Then %>
              <span class="ogun-badge ogle">&#214;&#287;le</span>
              <% Else %>
              <span class="ogun-badge aksam">Ak&#351;am</span>
              <% End If %>
              <%= GetOgunTipiEtiket(rsDegisiklik("ogun_tipi")) %>
            </td>
            <td class="eski-val"><% If LogSafeStr(rsDegisiklik("eski_deger")) <> "" Then %><%= rsDegisiklik("eski_deger") %><% Else %><em>(bo&#351;)</em><% End If %></td>
            <td class="yeni-val"><% If LogSafeStr(rsDegisiklik("yeni_deger")) <> "" Then %><%= rsDegisiklik("yeni_deger") %><% Else %><em>(bo&#351;)</em><% End If %></td>
            <td><%= rsDegisiklik("kullanici") %></td>
          </tr>
          <% rsDegisiklik.MoveNext
          Loop %>
        </tbody>
      </table>
      <% Else %>
      <div class="degisiklik-empty">Bu ay i&#231;in hen&#252;z kay&#305;tl&#305; &#246;&#287;&#252;n de&#287;i&#351;ikli&#287;i yok. Men&#252; d&#252;zenlendi&#287;inde de&#287;i&#351;iklikler burada g&#246;r&#252;necek.</div>
      <% End If %>
    </div>
    <% End If %>
  </div>
</div>
