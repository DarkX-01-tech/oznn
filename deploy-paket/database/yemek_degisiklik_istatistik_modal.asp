<style>
  #degisiklikModal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.45); z-index: 1100; justify-content: center; align-items: flex-start; padding: 40px 15px; overflow-y: auto; }
  #degisiklikModal.show { display: flex; }
  #degisiklikModal .modal-box { max-width: 960px; background: linear-gradient(135deg, #f0f9fa, #e6f4f1); border-radius: 16px; width: 100%; box-shadow: 0 20px 60px rgba(0,0,0,0.2); overflow: hidden; }
  #degisiklikModal .modal-header { background: linear-gradient(90deg, #ff9800, #e65100); color: #fff; padding: 14px 20px; display: flex; align-items: center; justify-content: space-between; }
  #degisiklikModal .modal-header h2 { font-size: 15px; display: flex; align-items: center; gap: 8px; margin: 0; font-weight: 600; }
  #degisiklikModal .modal-header h2 svg { width: 18px; height: 18px; fill: currentColor; }
  #degisiklikModal .modal-close { background: rgba(255,255,255,0.2); border: 2px solid rgba(255,255,255,0.4); color: #fff; width: 32px; height: 32px; border-radius: 50%; font-size: 16px; cursor: pointer; }
  #degisiklikModal .modal-body { padding: 18px; max-height: 65vh; overflow-y: auto; background: #fff; }
  .degisiklik-tablo-wrap { max-height: 60vh; overflow: auto; }
  .degisiklik-tablo { width: 100%; border-collapse: collapse; font-size: 12px; }
  .degisiklik-tablo thead th { background: #fff3e0; color: #bf360c; padding: 10px 12px; text-align: left; font-size: 10px; text-transform: uppercase; }
  .degisiklik-tablo tbody td { padding: 9px 12px; border-bottom: 1px solid #f5f5f5; vertical-align: top; }
  .degisiklik-tablo .eski-val { color: #c62828; text-decoration: line-through; }
  .degisiklik-tablo .yeni-val { color: #2e7d32; font-weight: 600; }
  .degisiklik-empty { text-align: center; padding: 30px 18px; color: #bbb; font-size: 12px; font-style: italic; }
  .degisiklik-modal-ogun { display: inline-block; padding: 3px 10px; border-radius: 8px; font-size: 10px; font-weight: 600; margin-right: 6px; }
  .degisiklik-modal-ogun.ogle { background: rgba(69,184,195,0.15); color: #2e8b91; }
  .degisiklik-modal-ogun.aksam { background: rgba(156,39,176,0.12); color: #7b1fa2; }
</style>

<div id="degisiklikModal" onclick="if(event.target&&event.target.id==='degisiklikModal') degisiklikModalKapat();">
  <div class="modal-box">
    <div class="modal-header">
      <h2>
        <svg viewBox="0 0 24 24"><path d="M13 3c-4.97 0-9 4.03-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42C8.27 19.99 10.51 21 13 21c4.97 0 9-4.03 9-9s-4.03-9-9-9zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>
        <span><%= ay_adi %> <%= secilen_yil %> Men&#252; De&#287;i&#351;iklikleri</span>
      </h2>
      <button type="button" class="modal-close" onclick="degisiklikModalKapat();">&times;</button>
    </div>
    <div class="modal-body">
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
                <span class="degisiklik-modal-ogun ogle">&#214;&#287;le</span>
                <% Else %>
                <span class="degisiklik-modal-ogun aksam">Ak&#351;am</span>
                <% End If %>
                <%= GetOgunTipiEtiket(rsDegisiklik("ogun_tipi")) %>
              </td>
              <td class="eski-val"><% If LogSafeStr(rsDegisiklik("eski_deger")) <> "" Then %><%= Server.HTMLEncode(rsDegisiklik("eski_deger")) %><% Else %><em>(bo&#351;)</em><% End If %></td>
              <td class="yeni-val"><% If LogSafeStr(rsDegisiklik("yeni_deger")) <> "" Then %><%= Server.HTMLEncode(rsDegisiklik("yeni_deger")) %><% Else %><em>(bo&#351;)</em><% End If %></td>
              <td><%= Server.HTMLEncode(rsDegisiklik("kullanici")) %></td>
            </tr>
            <% rsDegisiklik.MoveNext
            Loop %>
          </tbody>
        </table>
        <% Else %>
        <div class="degisiklik-empty">Bu ay i&#231;in hen&#252;z kay&#305;tl&#305; &#246;&#287;&#252;n de&#287;i&#351;ikli&#287;i yok.</div>
        <% End If %>
      </div>
    </div>
  </div>
</div>

<script type="text/javascript">
function degisiklikModalAc() {
  var modal = document.getElementById('degisiklikModal');
  if (!modal) { alert('Modal yuklenemedi. database\\yemek_degisiklik_istatistik_modal.asp dosyasini kontrol edin.'); return; }
  modal.className = 'show';
  modal.style.display = 'flex';
  document.body.style.overflow = 'hidden';
}
function degisiklikModalKapat() {
  var modal = document.getElementById('degisiklikModal');
  if (!modal) return;
  modal.className = '';
  modal.style.display = 'none';
  var detayModal = document.getElementById('detayModal');
  if (!detayModal || detayModal.className.indexOf('show') === -1) {
    document.body.style.overflow = '';
  }
}
</script>
