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
  .degisiklik-ac-btn { width: 100%; padding: 11px 16px; border: none; border-radius: 10px; background: linear-gradient(135deg, #ff9800, #e65100); color: #fff; font-size: 12px; font-weight: 600; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 8px; transition: all 0.3s; }
  .degisiklik-ac-btn:hover { transform: translateY(-2px); box-shadow: 0 6px 18px rgba(230,81,0,0.3); }
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
    <button type="button" class="degisiklik-ac-btn" id="degisiklikAcBtn">
      <svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg>
      De&#287;i&#351;iklikleri G&#246;r&#252;nt&#252;le
    </button>
  </div>
  <% End If %>
</div>
