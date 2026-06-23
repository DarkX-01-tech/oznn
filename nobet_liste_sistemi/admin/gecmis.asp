<%@ Language=VBScript CodePage=65001 %>
<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../database/connection.asp" -->
<!-- #include file="../lib/config.asp" -->
<!-- #include file="../lib/functions.asp" -->
<!-- #include file="../lib/auth.asp" -->
<!-- #include file="../lib/ui.asp" -->
<%
AdminGirisGerekli
DonemFormIsle

Dim seciliYil, seciliAy, i
seciliYil = GetSeciliYil()
seciliAy = GetSeciliAyKlasor()
%>
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="utf-8">
  <meta http-equiv="Content-Language" content="tr">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Geçmiş Dönem</title>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body class="admin-body">
  <div class="page-shell">
    <div class="admin-wrap">
      <div class="admin-topbar">
        <div>
          <p class="eyebrow">Arşiv</p>
          <h1>Geçmiş Dönem Listeleri</h1>
        </div>
        <% AdminNavGoster "gecmis" %>
      </div>

      <div class="dashboard-card">
        <h2>Dönem Seçin</h2>
        <form method="post" action="gecmis.asp" class="period-form">
          <div class="form-row">
            <div class="form-group">
              <label for="yil">Yıl</label>
              <select id="yil" name="yil" required>
                <% For i = GuncelYil() + 1 To GuncelYil() - 5 Step -1 %>
                  <option value="<%= i %>"<% If i = seciliYil Then %> selected<% End If %>><%= i %></option>
                <% Next %>
              </select>
            </div>
            <div class="form-group">
              <label for="ay">Ay</label>
              <select id="ay" name="ay" required>
                <% For i = 1 To 12 %>
                  <option value="<%= AyKlasorFromNumara(i) %>"<% If AyKlasorFromNumara(i) = seciliAy Then %> selected<% End If %>><%= TurkceAyAdi(i) %></option>
                <% Next %>
              </select>
            </div>
          </div>
          <button type="submit" class="btn btn-primary">Listeleri Göster</button>
        </form>
      </div>

      <% DonemOzetKartlariGoster %>

      <div class="content-panel gecmis-panel">
        <div class="section-title">Pendik E.A.H. Nöbet Listeleri</div>
        <div class="ay-baslik"><%= GetSeciliDonemBaslik() %></div>
        <table class="liste-tablo">
          <% RenderNobetListeTablosu BINA_PENDIK, pendikNobetListeleri, seciliYil, seciliAy %>
        </table>

        <div class="section-title">Prof. Dr. Asaf Ataseven Ek Hizmet Binası</div>
        <div class="ay-baslik"><%= GetSeciliDonemBaslik() %></div>
        <table class="liste-tablo">
          <% RenderNobetListeTablosu BINA_BASIBUYUK, basibuyukNobetListeleri, seciliYil, seciliAy %>
        </table>
      </div>
    </div>
  </div>
</body>
</html>
