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

    Set rsGuncSay = ConnYemek.Execute("SELECT COUNT(*) AS adet FROM yemek_guncelleme_log WHERE " & guncellemeLogFiltre)
    If Err.Number <> 0 Then
        guncellemeHata = True
    Else
        If Not rsGuncSay.EOF Then
            If Not IsNull(rsGuncSay("adet")) Then guncellemeSayisi = rsGuncSay("adet")
        End If
    End If
    rsGuncSay.Close: Set rsGuncSay = Nothing
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
  .degisiklik-ac-btn { width: 100%; padding: 11px 16px; border: none; border-radius: 10px; background: linear-gradient(135deg, #ff9800, #e65100); color: #fff; font-size: 12px; font-weight: 600; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 8px; }
  .degisiklik-ac-btn svg { width: 16px; height: 16px; fill: currentColor; pointer-events: none; }
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
    <button type="button" class="degisiklik-ac-btn" onclick="degisiklikModalAc(); return false;">
      <svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg>
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
  #degisiklikModal .degisiklik-tablo-wrap { max-height: 60vh; overflow: auto; }
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
      <div class="degisiklik-tablo-wrap" id="degisiklikModalIcerik">
        <% If degisiklik_kapsam = "ay" Then %>
        <% If degisiklikTabloHtml <> "" Then %>
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
        <% Else %>
        <div class="degisiklik-empty">Bu ay i&#231;in hen&#252;z kay&#305;tl&#305; &#246;&#287;&#252;n de&#287;i&#351;ikli&#287;i yok.</div>
        <% End If %>
        <% End If %>
      </div>
    </div>
  </div>
</div>
<script type="text/javascript">
var degisiklikKapsam = '<%= degisiklik_kapsam %>';
var degisiklikYil = <%= secilen_yil %>;
<% If degisiklik_kapsam = "yil" Then %>
var degisiklikKayitlar = <%= degisiklikJson %>;
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

function degisiklikAyGoster() {
  document.getElementById('degisiklikModalBaslik').innerHTML = degisiklikYil + ' Y\u0131l\u0131 Men\u00FC De\u011Fi\u015Fiklikleri';
  if (!degisiklikKayitlar || degisiklikKayitlar.length === 0) {
    document.getElementById('degisiklikModalIcerik').innerHTML = '<div class="degisiklik-empty">Bu y\u0131l i\u00E7in hen\u00FCz kay\u0131tl\u0131 \u00F6\u011F\u00FCn de\u011Fi\u015Fikli\u011Fi yok.</div>';
    return;
  }
  var ayBazli = {}, toplam = degisiklikKayitlar.length;
  for (var i = 0; i < degisiklikKayitlar.length; i++) {
    var a = degisiklikKayitlar[i].ay;
    if (!ayBazli[a]) ayBazli[a] = 0;
    ayBazli[a]++;
  }
  var aySayisi = Object.keys(ayBazli).length;
  var h = '<div class="modal-ozet">';
  h += '<div class="modal-ozet-card"><div class="mo-sayi">' + toplam + '</div><div class="mo-baslik">Toplam</div></div>';
  h += '<div class="modal-ozet-card"><div class="mo-sayi">' + aySayisi + '</div><div class="mo-baslik">Farkl\u0131 Ay</div></div>';
  h += '</div>';
  h += '<div class="modal-tablo-wrap"><table class="modal-tablo"><thead><tr><th>#</th><th>Ay</th><th>De\u011Fi\u015Fiklik</th><th></th></tr></thead><tbody>';
  var sira = 0;
  var ayKeys = Object.keys(ayBazli).sort(function(a,b){return parseInt(a)-parseInt(b);});
  for (var j = 0; j < ayKeys.length; j++) {
    sira++;
    var ayK = parseInt(ayKeys[j]);
    h += '<tr onclick="degisiklikGunGoster(' + ayK + ')" style="cursor:pointer;">';
    h += '<td>' + sira + '</td>';
    h += '<td><span class="ay-link">' + degisiklikAyAdlari[ayK] + '</span></td>';
    h += '<td><strong>' + ayBazli[ayK] + ' Kez</strong></td>';
    h += '<td><button type="button" class="detay-btn" onclick="event.stopPropagation();degisiklikGunGoster(' + ayK + ')"><svg viewBox="0 0 24 24"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>G\u00fcn Baz</button></td>';
    h += '</tr>';
  }
  h += '</tbody></table></div>';
  document.getElementById('degisiklikModalIcerik').innerHTML = h;
}

function degisiklikGunGoster(ayNo) {
  document.getElementById('degisiklikModalBaslik').innerHTML = degisiklikAyAdlari[ayNo] + ' ' + degisiklikYil + ' - G\u00FCn Detay\u0131';
  var kayitlar = [];
  for (var i = 0; i < degisiklikKayitlar.length; i++) {
    if (parseInt(degisiklikKayitlar[i].ay) === parseInt(ayNo)) kayitlar.push(degisiklikKayitlar[i]);
  }
  var h = '<div style="margin-bottom:10px;"><button type="button" class="detay-btn" onclick="degisiklikAyGoster();return false;"><svg viewBox="0 0 24 24"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>Ay Baz\u0131na D\u00F6n</button></div>';
  if (kayitlar.length === 0) {
    h += '<div class="degisiklik-empty">Bu ay i\u00E7in kay\u0131t yok.</div>';
  } else {
    h += '<table class="degisiklik-tablo"><thead><tr>';
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
    h += '</tbody></table>';
  }
  document.getElementById('degisiklikModalIcerik').innerHTML = h;
}
<% End If %>

function degisiklikModalAc() {
  var modal = document.getElementById('degisiklikModal');
  if (!modal) return;
  if (degisiklikKapsam === 'yil' && typeof degisiklikAyGoster === 'function') degisiklikAyGoster();
  modal.className = 'modal-overlay show';
  document.body.style.overflow = 'hidden';
}
function degisiklikModalKapat() {
  var modal = document.getElementById('degisiklikModal');
  if (!modal) return;
  modal.className = 'modal-overlay';
  var detayModal = document.getElementById('detayModal');
  if (!detayModal || detayModal.className.indexOf('show') === -1) {
    document.body.style.overflow = '';
  }
}
</script>
<%
End If
%>
