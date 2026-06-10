<%
' degisiklik_bolum: "veri" | "buton" | "modal"
' degisiklik_kapsam: "ay" (varsayilan) | "yil"
' Gerekli: secilen_yil, menu_tipi | aylik icin ay_adi, secilen_ay

If degisiklik_kapsam = "" Then degisiklik_kapsam = "ay"

If degisiklikVeriHazir <> "evet" Then
    degisiklikVeriHazir = "evet"

    degisiklikVar = False
    degisiklikHata = False
    degisiklikSayisi = 0
    degisenMenuGun = 0
    ogleDegisim = 0
    aksamDegisim = 0
    degisiklikTabloHtml = ""
    degisiklikJson = "[]"
    guncellemeTabloHtml = ""
    guncellemeJson = "[]"

    degisiklikLogFiltre = "yil = " & secilen_yil & " AND menu_tipi = '" & menu_tipi & "'"
    If degisiklik_kapsam = "ay" Then
        degisiklikLogFiltre = degisiklikLogFiltre & " AND ay = " & secilen_ay
    End If

    On Error Resume Next
    Set rsDegisiklik = ConnYemek.Execute("SELECT * FROM yemek_degisiklik_log WHERE " & degisiklikLogFiltre & " ORDER BY degisiklik_tarihi DESC")
    If Err.Number <> 0 Then
        degisiklikHata = True
    Else
        degisiklikVar = True
        If Not rsDegisiklik.EOF Then
            Set rsDegOzet = ConnYemek.Execute("SELECT COUNT(DISTINCT menu_tarih) AS gun FROM yemek_degisiklik_log WHERE " & degisiklikLogFiltre)
            If Not rsDegOzet.EOF Then degisenMenuGun = rsDegOzet("gun")
            rsDegOzet.Close: Set rsDegOzet = Nothing

            Set rsOgleAksam = ConnYemek.Execute("SELECT ogun_tipi FROM yemek_degisiklik_log WHERE " & degisiklikLogFiltre)
            Do While Not rsOgleAksam.EOF
                If OgunTipiOgleMi(rsOgleAksam("ogun_tipi")) Then
                    ogleDegisim = ogleDegisim + 1
                Else
                    aksamDegisim = aksamDegisim + 1
                End If
                rsOgleAksam.MoveNext
            Loop
            rsOgleAksam.Close: Set rsOgleAksam = Nothing

            degisiklikJsonIlk = True
            degisiklikJsonBuf = ""

            Do While Not rsDegisiklik.EOF
                degisiklikSayisi = degisiklikSayisi + 1
                If OgunTipiOgleMi(rsDegisiklik("ogun_tipi")) Then
                    satirOgunKod = "ogle"
                    satirOgun = "<span class=""degisiklik-modal-ogun ogle"">&#214;&#287;le</span> " & GetOgunTipiEtiket(rsDegisiklik("ogun_tipi"))
                Else
                    satirOgunKod = "aksam"
                    satirOgun = "<span class=""degisiklik-modal-ogun aksam"">Ak&#351;am</span> " & GetOgunTipiEtiket(rsDegisiklik("ogun_tipi"))
                End If
                If LogSafeStr(rsDegisiklik("eski_deger")) <> "" Then
                    satirEski = LogHtmlSafe(rsDegisiklik("eski_deger"))
                    satirEskiJson = LogJsonStr(rsDegisiklik("eski_deger"))
                Else
                    satirEski = "<em>(bo&#351;)</em>"
                    satirEskiJson = ""
                End If
                If LogSafeStr(rsDegisiklik("yeni_deger")) <> "" Then
                    satirYeni = LogHtmlSafe(rsDegisiklik("yeni_deger"))
                    satirYeniJson = LogJsonStr(rsDegisiklik("yeni_deger"))
                Else
                    satirYeni = "<em>(bo&#351;)</em>"
                    satirYeniJson = ""
                End If
                satirKullanici = LogHtmlSafe(rsDegisiklik("kullanici"))
                satirDegTarih = FormatDateTime(rsDegisiklik("degisiklik_tarihi"), 0)
                satirMenuGun = FormatDateTime(rsDegisiklik("menu_tarih"), 2)

                If degisiklik_kapsam = "ay" Then
                    degisiklikTabloHtml = degisiklikTabloHtml & "<tr>"
                    degisiklikTabloHtml = degisiklikTabloHtml & "<td>" & satirDegTarih & "</td>"
                    degisiklikTabloHtml = degisiklikTabloHtml & "<td>" & satirMenuGun & "</td>"
                    degisiklikTabloHtml = degisiklikTabloHtml & "<td>" & satirOgun & "</td>"
                    degisiklikTabloHtml = degisiklikTabloHtml & "<td class=""eski-val"">" & satirEski & "</td>"
                    degisiklikTabloHtml = degisiklikTabloHtml & "<td class=""yeni-val"">" & satirYeni & "</td>"
                    degisiklikTabloHtml = degisiklikTabloHtml & "<td>" & satirKullanici & "</td>"
                    degisiklikTabloHtml = degisiklikTabloHtml & "</tr>"
                Else
                    If Not degisiklikJsonIlk Then degisiklikJsonBuf = degisiklikJsonBuf & ","
                    degisiklikJsonIlk = False
                    degisiklikJsonBuf = degisiklikJsonBuf & "{"
                    degisiklikJsonBuf = degisiklikJsonBuf & """ay"":" & CInt(rsDegisiklik("ay")) & ","
                    degisiklikJsonBuf = degisiklikJsonBuf & """degTarih"":""" & LogJsonStr(satirDegTarih) & ""","
                    degisiklikJsonBuf = degisiklikJsonBuf & """menuGun"":""" & LogJsonStr(satirMenuGun) & ""","
                    degisiklikJsonBuf = degisiklikJsonBuf & """ogunKod"":""" & satirOgunKod & ""","
                    degisiklikJsonBuf = degisiklikJsonBuf & """ogunTipi"":""" & LogJsonStr(rsDegisiklik("ogun_tipi")) & ""","
                    degisiklikJsonBuf = degisiklikJsonBuf & """eski"":""" & satirEskiJson & ""","
                    degisiklikJsonBuf = degisiklikJsonBuf & """yeni"":""" & satirYeniJson & ""","
                    degisiklikJsonBuf = degisiklikJsonBuf & """kullanici"":""" & LogJsonStr(rsDegisiklik("kullanici")) & """"
                    degisiklikJsonBuf = degisiklikJsonBuf & "}"
                End If
                rsDegisiklik.MoveNext
            Loop

            If degisiklik_kapsam = "yil" Then
                degisiklikJson = "[" & degisiklikJsonBuf & "]"
            End If
        End If
    End If
    Err.Clear

    guncellemeHata = False
    guncellemeSayisi = 0
    If menu_tipi = "diyet" Then
        guncellemeFiltre = " AND (guncelleme_tipi LIKE '%Diyet%' OR aciklama LIKE '%Diyet%')"
    Else
        guncellemeFiltre = " AND ((guncelleme_tipi NOT LIKE '%Diyet%' OR guncelleme_tipi IS NULL) AND (aciklama NOT LIKE '%Diyet%' OR aciklama IS NULL))"
    End If

    guncellemeLogFiltre = "yil = " & secilen_yil & guncellemeFiltre
    If degisiklik_kapsam = "ay" Then
        guncellemeLogFiltre = "yil = " & secilen_yil & " AND ay = " & secilen_ay & guncellemeFiltre
    End If

    Set rsGunc = ConnYemek.Execute("SELECT * FROM yemek_guncelleme_log WHERE " & guncellemeLogFiltre & " ORDER BY guncelleme_tarihi DESC")
    If Err.Number <> 0 Then
        guncellemeHata = True
    Else
        guncellemeJsonIlk = True
        guncellemeJsonBuf = ""
        Do While Not rsGunc.EOF
            guncellemeSayisi = guncellemeSayisi + 1
            satirGuncTarih = FormatDateTime(rsGunc("guncelleme_tarihi"), 0)
            satirGuncAy = CInt(rsGunc("ay"))
            satirGuncAyAdi = ""
            If degisiklik_kapsam = "ay" And ay_adi <> "" Then
                satirGuncAyAdi = ay_adi
            End If
            satirGuncTip = GuncellemeTipEtiket(rsGunc("guncelleme_tipi"))
            satirGuncKullanici = GuncellemeMetinGoster(rsGunc("kullanici"))
            satirGuncOzet = GuncellemeOzetMetni(rsGunc("guncelleme_tipi"), rsGunc("aciklama"), secilen_yil, satirGuncAy, satirGuncAyAdi)

            If InStr(LCase(EntityDecode(LogSafeStr(rsGunc("guncelleme_tipi")))), "toplu") > 0 Or InStr(LCase(EntityDecode(LogSafeStr(rsGunc("guncelleme_tipi")))), "ayl") > 0 Then
                satirGuncBadge = "aylik"
            Else
                satirGuncBadge = "gunluk"
            End If

            If degisiklik_kapsam = "ay" Then
                guncellemeTabloHtml = guncellemeTabloHtml & "<tr>"
                guncellemeTabloHtml = guncellemeTabloHtml & "<td>" & satirGuncTarih & "</td>"
                guncellemeTabloHtml = guncellemeTabloHtml & "<td><span class=""gunc-badge " & satirGuncBadge & """>" & satirGuncTip & "</span></td>"
                guncellemeTabloHtml = guncellemeTabloHtml & "<td class=""gunc-ozet"">" & satirGuncOzet & "</td>"
                guncellemeTabloHtml = guncellemeTabloHtml & "<td>" & satirGuncKullanici & "</td>"
                guncellemeTabloHtml = guncellemeTabloHtml & "</tr>"
            Else
                If Not guncellemeJsonIlk Then guncellemeJsonBuf = guncellemeJsonBuf & ","
                guncellemeJsonIlk = False
                guncellemeJsonBuf = guncellemeJsonBuf & "{"
                guncellemeJsonBuf = guncellemeJsonBuf & """ay"":" & satirGuncAy & ","
                guncellemeJsonBuf = guncellemeJsonBuf & """tarih"":""" & LogJsonStr(satirGuncTarih) & ""","
                guncellemeJsonBuf = guncellemeJsonBuf & """tip"":""" & LogJsonStr(rsGunc("guncelleme_tipi")) & ""","
                guncellemeJsonBuf = guncellemeJsonBuf & """aciklama"":""" & LogJsonStr(rsGunc("aciklama")) & ""","
                guncellemeJsonBuf = guncellemeJsonBuf & """kullanici"":""" & LogJsonStr(rsGunc("kullanici")) & """"
                guncellemeJsonBuf = guncellemeJsonBuf & "}"
            End If
            rsGunc.MoveNext
        Loop
        If degisiklik_kapsam = "yil" Then
            guncellemeJson = "[" & guncellemeJsonBuf & "]"
        End If
    End If
    If Not rsGunc Is Nothing Then rsGunc.Close: Set rsGunc = Nothing
    Err.Clear
    On Error GoTo 0
End If

If degisiklik_bolum = "buton" Then
%>
<style>
  .degisiklik-btn-card { background: #fff; border-radius: 12px; box-shadow: 0 4px 14px rgba(0,0,0,0.05); overflow: hidden; margin-top: 15px; }
  .degisiklik-btn-header { background: linear-gradient(135deg, #ff9800, #e65100); color: #fff; padding: 14px 16px; display: flex; align-items: center; gap: 8px; }
  .degisiklik-btn-header h3 { font-size: 13px; margin: 0; font-weight: 600; }
  .degisiklik-btn-header svg { width: 18px; height: 18px; fill: currentColor; flex-shrink: 0; }
  .degisiklik-btn-ozet { padding: 12px 16px; display: flex; gap: 10px; border-bottom: 1px solid #f0f0f0; }
  .degisiklik-btn-ozet-item { flex: 1; background: #fff8f0; border: 1px solid #ffe0b2; border-radius: 8px; padding: 8px; text-align: center; }
  .degisiklik-btn-ozet-item .dbo-sayi { font-size: 16px; font-weight: 700; color: #e65100; }
  .degisiklik-btn-ozet-item .dbo-label { font-size: 9px; color: #999; text-transform: uppercase; margin-top: 2px; }
  .degisiklik-btn-action { padding: 14px 16px; }
  .degisiklik-gor-btn { width: 100%; justify-content: center; padding: 10px 16px; font-size: 12px; gap: 6px; }
  .degisiklik-gor-btn:active { transform: translateY(0); }
  .degisiklik-btn-warn { padding: 16px; color: #e65100; font-size: 11px; text-align: center; }
</style>
<div class="degisiklik-btn-card">
  <div class="degisiklik-btn-header">
    <svg viewBox="0 0 24 24"><path d="M13 3c-4.97 0-9 4.03-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42C8.27 19.99 10.51 21 13 21c4.97 0 9-4.03 9-9s-4.03-9-9-9zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>
    <h3>Men&#252; De&#287;i&#351;iklikleri</h3>
  </div>
  <% If degisiklikHata And guncellemeHata Then %>
  <div class="degisiklik-btn-warn">Log tablolar&#305;na eri&#351;ilemiyor.</div>
  <% Else %>
  <div class="degisiklik-btn-ozet">
    <div class="degisiklik-btn-ozet-item"><div class="dbo-sayi"><%= guncellemeSayisi %></div><div class="dbo-label">G&#252;ncelleme</div></div>
    <div class="degisiklik-btn-ozet-item"><div class="dbo-sayi"><%= degisiklikSayisi %></div><div class="dbo-label">&#214;&#287;&#252;n De&#287;i&#351;ikli&#287;i</div></div>
  </div>
  <div class="degisiklik-btn-action">
    <button type="button" class="detay-btn degisiklik-gor-btn" onclick="degisiklikModalAc(); return false;">
      <svg viewBox="0 0 24 24"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>
      De&#287;i&#351;iklikleri G&#246;r&#252;nt&#252;le
    </button>
  </div>
  <% End If %>
</div>
<%
ElseIf degisiklik_bolum = "modal" Then
%>
<style>
  #degisiklikModal .modal-box { max-width: 960px; }
  #degisiklikModalIcerik { overflow: visible; }
  .degisiklik-modal-nav { margin-bottom: 12px; padding: 10px 0 6px; overflow: visible; }
  .degisiklik-tablo-scroll { max-height: 370px; overflow-x: auto; overflow-y: auto; }
  .degisiklik-tablo { width: 100%; border-collapse: collapse; font-size: 12px; }
  .degisiklik-tablo thead th { background: #fff3e0; color: #bf360c; padding: 10px 12px; text-align: left; font-size: 10px; text-transform: uppercase; }
  .degisiklik-tablo tbody td { padding: 9px 12px; border-bottom: 1px solid #f5f5f5; vertical-align: top; }
  .degisiklik-tablo .eski-val { color: #c62828; text-decoration: line-through; }
  .degisiklik-tablo .yeni-val { color: #2e7d32; font-weight: 600; }
  .degisiklik-empty { text-align: center; padding: 30px 18px; color: #bbb; font-size: 12px; font-style: italic; }
  .degisiklik-modal-ogun { display: inline-block; padding: 3px 10px; border-radius: 8px; font-size: 10px; font-weight: 600; margin-right: 6px; }
  .degisiklik-modal-ogun.ogle { background: rgba(69,184,195,0.15); color: #2e8b91; }
  .degisiklik-modal-ogun.aksam { background: rgba(156,39,176,0.12); color: #7b1fa2; }
  #degisiklikModal .modal-ozet-card { background: #fff8f0; border-color: #ffe0b2; }
  #degisiklikModal .modal-ozet-card .mo-sayi { color: #e65100; }
  #degisiklikModal .modal-tablo thead th { background: linear-gradient(135deg, #ff9800, #e65100); }
  #degisiklikModal .ay-link { color: #e65100; font-weight: 600; }
  .degisiklik-section { margin-bottom: 18px; }
  .degisiklik-section-title { font-size: 11px; font-weight: 700; color: #e65100; margin: 0 0 8px; text-transform: uppercase; letter-spacing: 0.4px; display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
  .degisiklik-section-title svg { width: 14px; height: 14px; fill: currentColor; flex-shrink: 0; }
  .degisiklik-section-etiket { font-weight: 600; text-transform: none; letter-spacing: 0; opacity: 0.9; }
  .gunc-badge { display: inline-block; padding: 3px 10px; border-radius: 8px; font-size: 10px; font-weight: 600; white-space: nowrap; }
  .gunc-badge.aylik { background: rgba(255,152,0,0.15); color: #e65100; }
  .gunc-badge.gunluk { background: rgba(69,184,195,0.15); color: #2e8b91; }
  .gunc-ozet { color: #333; font-weight: 500; }
</style>
<div class="modal-overlay" id="degisiklikModal" onclick="if(event.target===this) degisiklikModalKapat();">
  <div class="modal-box">
    <div class="modal-header" style="background:linear-gradient(90deg,#ff9800,#e65100);">
      <h2>
        <svg viewBox="0 0 24 24"><path d="M13 3c-4.97 0-9 4.03-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42C8.27 19.99 10.51 21 13 21c4.97 0 9-4.03 9-9s-4.03-9-9-9zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>
        <span id="degisiklikModalBaslik"><% If degisiklik_kapsam = "yil" Then %><%= secilen_yil %> Y&#305;l&#305; Men&#252; De&#287;i&#351;iklikleri<% Else %><%= ay_adi & " " & secilen_yil %> Men&#252; De&#287;i&#351;iklikleri<% End If %></span>
      </h2>
      <button type="button" class="modal-close" onclick="degisiklikModalKapat();">&times;</button>
    </div>
    <div class="modal-body">
      <div id="degisiklikModalIcerik">
        <% If degisiklik_kapsam = "ay" Then %>
        <% If guncellemeTabloHtml <> "" Then %>
        <div class="degisiklik-section">
          <div class="degisiklik-section-title">
            <svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>
            &#214;&#287;&#252;n De&#287;i&#351;iklikleri <span class="degisiklik-section-etiket">(Ayl&#305;k Liste)</span>
          </div>
          <div class="degisiklik-tablo-scroll gunc-scroll">
          <table class="degisiklik-tablo">
            <thead>
              <tr>
                <th>G&#252;ncelleme Tarihi/Saati</th>
                <th>&#304;&#351;lem Tipi</th>
                <th>A&#231;&#305;klama</th>
                <th>Kullan&#305;c&#305;</th>
              </tr>
            </thead>
            <tbody><%= guncellemeTabloHtml %></tbody>
          </table>
          </div>
        </div>
        <% End If %>
        <% If degisiklikTabloHtml <> "" Then %>
        <div class="degisiklik-section">
          <div class="degisiklik-section-title">
            <svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>
            &#214;&#287;&#252;n De&#287;i&#351;iklikleri
          </div>
          <div class="degisiklik-tablo-scroll ogun-scroll">
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
            <tbody><%= degisiklikTabloHtml %></tbody>
          </table>
          </div>
        </div>
        <% End If %>
        <% If guncellemeTabloHtml = "" And degisiklikTabloHtml = "" Then %>
        <div class="degisiklik-empty">Bu ay i&#231;in hen&#252;z kay&#305;tl&#305; g&#252;ncelleme veya &#246;&#287;&#252;n de&#287;i&#351;ikli&#287;i yok.</div>
        <% End If %>
        <% End If %>
      </div>
    </div>
  </div>
</div>
<script type="text/javascript">
function modalAnimAc(el) {
  if (!el) return;
  el.classList.remove('closing');
  el.classList.add('show');
}
function modalAnimKapat(el, done) {
  if (!el || !el.classList.contains('show')) { if (done) done(); return; }
  el.classList.add('closing');
  setTimeout(function() {
    el.classList.remove('show', 'closing');
    if (done) done();
  }, 280);
}
var degisiklikKapsam = '<%= degisiklik_kapsam %>';
var degisiklikYil = <%= secilen_yil %>;
<% If degisiklik_kapsam = "yil" Then %>
var degisiklikKayitlar = <%= degisiklikJson %>;
var guncellemeKayitlar = <%= guncellemeJson %>;
var degisiklikAyAdlari = ['','Ocak','\u015eubat','Mart','Nisan','May\u0131s','Haziran','Temmuz','A\u011fustos','Eyl\u00fcl','Ekim','Kas\u0131m','Aral\u0131k'];
var degisiklikOgunEtiket = {
  'ogle_corba':'\u00C7orba','ogle_ana':'Ana Yemek','ogle_yan':'Yan \u00DCr\u00FCn','ogle_tatli':'Tatl\u0131/Meyve',
  'aksam_corba':'\u00C7orba','aksam_ana':'Ana Yemek','aksam_yan':'Yan \u00DCr\u00FCn','aksam_tatli':'Tatl\u0131/Meyve'
};

function degisiklikOgunHtml(kayit) {
  var oa = kayit.ogunKod === 'ogle' ? '\u00D6\u011Fle' : 'Ak\u015Fam';
  var etiket = degisiklikOgunEtiket[kayit.ogunTipi] || kayit.ogunTipi;
  return '<span class="degisiklik-modal-ogun ' + kayit.ogunKod + '">' + oa + '</span> ' + oa + ' - ' + etiket;
}

function degisiklikDegerHtml(val, sinif) {
  if (!val) return '<em>(bo\u015f)</em>';
  return '<span class="' + sinif + '">' + val + '</span>';
}

function guncellemeTipEtiketJs(tip) {
  var t = (tip || '').toLowerCase();
  if (t.indexOf('toplu') >= 0 || t.indexOf('ayl') >= 0) {
    return t.indexOf('diyet') >= 0 ? 'Ayl\u0131k Liste G\u00FCncellenmesi (Diyet)' : 'Ayl\u0131k Liste G\u00FCncellenmesi';
  }
  if (t.indexOf('tek') >= 0 || t.indexOf('tekil') >= 0 || t.indexOf('g\u00fcnl') >= 0) {
    return t.indexOf('diyet') >= 0 ? 'G\u00FCnl\u00FCk Liste G\u00FCncellenmesi (Diyet)' : 'G\u00FCnl\u00FCk Liste G\u00FCncellenmesi';
  }
  return tip || '-';
}

function guncellemeBadgeSinif(tip) {
  var t = (tip || '').toLowerCase();
  if (t.indexOf('toplu') >= 0 || t.indexOf('ayl') >= 0) return 'aylik';
  return 'gunluk';
}

function guncellemeOzetMetni(k) {
  var tip = (k.tip || '').toLowerCase();
  if (tip.indexOf('toplu') >= 0 || tip.indexOf('ayl') >= 0) {
    var t = degisiklikAyAdlari[parseInt(k.ay)] + ' ' + degisiklikYil + ' ay\u0131n\u0131n men\u00FC listesi g\u00FCncellendi';
    if (k.aciklama) t += ' (' + k.aciklama + ')';
    return t;
  }
  return k.aciklama || guncellemeTipEtiketJs(k.tip) || '-';
}

function guncellemeTabloHtmlJs(liste) {
  if (!liste || liste.length === 0) return '';
  var h = '<div class="degisiklik-section"><div class="degisiklik-section-title"><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>\u00D6\u011F\u00FCn De\u011Fi\u015Fiklikleri <span class="degisiklik-section-etiket">(Ayl\u0131k Liste)</span></div>';
  h += '<div class="degisiklik-tablo-scroll gunc-scroll"><table class="degisiklik-tablo"><thead><tr><th>G\u00FCncelleme Tarihi/Saati</th><th>\u0130\u015Flem Tipi</th><th>A\u00E7\u0131klama</th><th>Kullan\u0131c\u0131</th></tr></thead><tbody>';
  for (var g = 0; g < liste.length; g++) {
    var kg = liste[g];
    h += '<tr><td>' + kg.tarih + '</td>';
    h += '<td><span class="gunc-badge ' + guncellemeBadgeSinif(kg.tip) + '">' + guncellemeTipEtiketJs(kg.tip) + '</span></td>';
    h += '<td class="gunc-ozet">' + guncellemeOzetMetni(kg) + '</td>';
    h += '<td>' + kg.kullanici + '</td></tr>';
  }
  h += '</tbody></table></div></div>';
  return h;
}

function degisiklikAyGoster() {
  document.getElementById('degisiklikModalBaslik').innerHTML = degisiklikYil + ' Y\u0131l\u0131 Men\u00FC De\u011Fi\u015Fiklikleri';

  var ayBazliOgun = {}, aySet = {}, toplamOgun = 0, guncellemeToplam = 0;
  if (degisiklikKayitlar && degisiklikKayitlar.length > 0) {
    toplamOgun = degisiklikKayitlar.length;
    for (var i = 0; i < degisiklikKayitlar.length; i++) {
      var a = degisiklikKayitlar[i].ay;
      aySet[a] = true;
      if (!ayBazliOgun[a]) ayBazliOgun[a] = 0;
      ayBazliOgun[a]++;
    }
  }
  if (guncellemeKayitlar && guncellemeKayitlar.length > 0) {
    guncellemeToplam = guncellemeKayitlar.length;
    for (var gi = 0; gi < guncellemeKayitlar.length; gi++) {
      aySet[guncellemeKayitlar[gi].ay] = true;
    }
  }

  var ayKeys = Object.keys(aySet).sort(function(a,b){return parseInt(a)-parseInt(b);});
  if (ayKeys.length === 0) {
    document.getElementById('degisiklikModalIcerik').innerHTML = '<div class="degisiklik-empty">Bu y\u0131l i\u00E7in hen\u00FCz kay\u0131tl\u0131 g\u00FCncelleme veya \u00F6\u011F\u00FCn de\u011Fi\u015Fikli\u011Fi yok.</div>';
    return;
  }

  var hAy = '<div class="modal-ozet">';
  hAy += '<div class="modal-ozet-card"><div class="mo-sayi">' + toplamOgun + '</div><div class="mo-baslik">Toplam</div></div>';
  hAy += '<div class="modal-ozet-card"><div class="mo-sayi">' + ayKeys.length + '</div><div class="mo-baslik">Farkl\u0131 Ay</div></div>';
  if (guncellemeToplam > 0) {
    hAy += '<div class="modal-ozet-card"><div class="mo-sayi">' + guncellemeToplam + '</div><div class="mo-baslik">Liste G\u00FCncelleme</div></div>';
  }
  hAy += '</div>';
  hAy += '<div class="degisiklik-section"><div class="degisiklik-section-title"><svg viewBox="0 0 24 24"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V9h14v10z"/></svg>Ay Baz\u0131 \u00D6zet</div>';
  hAy += '<div class="degisiklik-tablo-scroll gunc-scroll"><div class="modal-tablo-wrap"><table class="modal-tablo"><thead><tr><th>#</th><th>Ay</th><th>De\u011Fi\u015Fiklik</th><th></th></tr></thead><tbody>';
  var sira = 0;
  for (var j = 0; j < ayKeys.length; j++) {
    sira++;
    var ayK = parseInt(ayKeys[j]);
    var ogunSay = ayBazliOgun[ayK] || 0;
    var degisiklikMetin = ogunSay > 0 ? (ogunSay + ' Kez') : '<em style="color:#999;">Liste g\u00FCncellemesi</em>';
    hAy += '<tr onclick="degisiklikGunGoster(' + ayK + ')" style="cursor:pointer;">';
    hAy += '<td>' + sira + '</td>';
    hAy += '<td><span class="ay-link">' + degisiklikAyAdlari[ayK] + '</span></td>';
    hAy += '<td><strong>' + degisiklikMetin + '</strong></td>';
    hAy += '<td><button type="button" class="detay-btn" onclick="event.stopPropagation();degisiklikGunGoster(' + ayK + ')"><svg viewBox="0 0 24 24"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>G\u00fcn Baz</button></td>';
    hAy += '</tr>';
  }
  hAy += '</tbody></table></div></div></div>';
  document.getElementById('degisiklikModalIcerik').innerHTML = hAy;
}

function degisiklikGunGoster(ayNo) {
  document.getElementById('degisiklikModalBaslik').innerHTML = degisiklikAyAdlari[ayNo] + ' ' + degisiklikYil + ' - G\u00FCn Detay\u0131';
  var kayitlar = [];
  for (var i = 0; i < degisiklikKayitlar.length; i++) {
    if (parseInt(degisiklikKayitlar[i].ay) === parseInt(ayNo)) kayitlar.push(degisiklikKayitlar[i]);
  }
  var guncAy = [];
  if (guncellemeKayitlar) {
    for (var gx = 0; gx < guncellemeKayitlar.length; gx++) {
      if (parseInt(guncellemeKayitlar[gx].ay) === parseInt(ayNo)) guncAy.push(guncellemeKayitlar[gx]);
    }
  }
  var h = '<div class="degisiklik-modal-nav"><button type="button" class="detay-btn" onclick="degisiklikAyGoster();return false;"><svg viewBox="0 0 24 24"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>Ay Baz\u0131na D\u00F6n</button></div>';
  h += guncellemeTabloHtmlJs(guncAy);
  if (kayitlar.length === 0) {
    if (guncAy.length === 0) {
      h += '<div class="degisiklik-empty">Bu ay i\u00E7in kay\u0131t yok.</div>';
    } else {
      h += '<div class="degisiklik-empty">Bu ay i\u00E7in \u00F6\u011F\u00FCn de\u011Fi\u015Fikli\u011Fi kayd\u0131 yok.</div>';
    }
  } else {
    h += '<div class="degisiklik-section"><div class="degisiklik-section-title"><svg viewBox="0 0 24 24"><path d="M11 9H9V2H7v7H5V2H3v7c0 2.12 1.66 3.84 3.75 3.97V22h2.5v-9.03C11.34 12.84 13 11.12 13 9V2h-2v7zm5-3v8h2.5v8H21V2c-2.76 0-5 2.24-5 4z"/></svg>\u00D6\u011F\u00FCn De\u011Fi\u015Fiklikleri</div>';
    h += '<div class="degisiklik-tablo-scroll ogun-scroll"><table class="degisiklik-tablo"><thead><tr>';
    h += '<th>De\u011Fi\u015Fiklik Tarihi/Saati</th><th>Men\u00FC G\u00FCn\u00FC</th><th>\u00D6\u011F\u00FCn</th><th>Eski \u00D6\u011F\u00FCn</th><th>Yeni \u00D6\u011F\u00FCn</th><th>Kullan\u0131c\u0131</th>';
    h += '</tr></thead><tbody>';
    for (var p = 0; p < kayitlar.length; p++) {
      var k = kayitlar[p];
      h += '<tr>';
      h += '<td>' + k.degTarih + '</td>';
      h += '<td>' + k.menuGun + '</td>';
      h += '<td>' + degisiklikOgunHtml(k) + '</td>';
      h += '<td class="eski-val">' + (k.eski || '<em>(bo\u015f)</em>') + '</td>';
      h += '<td class="yeni-val">' + (k.yeni || '<em>(bo\u015f)</em>') + '</td>';
      h += '<td>' + k.kullanici + '</td>';
      h += '</tr>';
    }
    h += '</tbody></table></div></div>';
  }
  document.getElementById('degisiklikModalIcerik').innerHTML = h;
}
<% End If %>

function degisiklikModalAc() {
  var modal = document.getElementById('degisiklikModal');
  if (!modal) return;
  if (degisiklikKapsam === 'yil' && typeof degisiklikAyGoster === 'function') degisiklikAyGoster();
  modalAnimAc(modal);
  document.body.style.overflow = 'hidden';
}
function degisiklikModalKapat() {
  var modal = document.getElementById('degisiklikModal');
  if (!modal) return;
  modalAnimKapat(modal, function() {
    var detayModal = document.getElementById('detayModal');
    if (!detayModal || !detayModal.classList.contains('show')) {
      document.body.style.overflow = '';
    }
  });
}
</script>
<%
End If
%>
