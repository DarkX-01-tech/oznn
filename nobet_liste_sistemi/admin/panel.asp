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

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Trim(Request.Form("git")) = "kaydet" Then
        Response.Redirect "panel.asp?mesaj=donem"
        Response.End
    End If
    If Trim(Request.Form("git")) = "listeler" Then
        Response.Redirect "listeler.asp"
        Response.End
    End If
    If Trim(Request.Form("git")) = "gecmis" Then
        Response.Redirect "gecmis.asp"
        Response.End
    End If
End If

Dim yilSecenek, i
%>
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="utf-8">
  <meta http-equiv="Content-Language" content="tr">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Yönetim Paneli</title>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body class="admin-body">
  <div class="page-shell">
    <div class="admin-wrap">
      <div class="admin-topbar">
        <div>
          <p class="eyebrow">Ana Yönetim</p>
          <h1>Nöbet Liste Yönetimi</h1>
        </div>
        <% AdminNavGoster "panel" %>
      </div>

      <% FlashMesajGoster %>
      <% DonemOzetKartlariGoster %>

      <div class="dashboard-grid">
        <div class="dashboard-card">
          <h2>Dönem Seçimi</h2>
          <p>Yükleme ve liste yönetimi için yıl ve ay seçin.</p>
          <form method="post" action="panel.asp" class="period-form">
            <div class="form-row">
              <div class="form-group">
                <label for="yil">Yıl</label>
                <select id="yil" name="yil" required>
                  <% For i = GuncelYil() + 1 To GuncelYil() - 5 Step -1 %>
                    <option value="<%= i %>"<% If i = GetSeciliYil() Then %> selected<% End If %>><%= i %></option>
                  <% Next %>
                </select>
              </div>
              <div class="form-group">
                <label for="ay">Ay</label>
                <select id="ay" name="ay" required>
                  <% For i = 1 To 12 %>
                    <option value="<%= AyKlasorFromNumara(i) %>"<% If AyKlasorFromNumara(i) = GetSeciliAyKlasor() Then %> selected<% End If %>><%= TurkceAyAdi(i) %></option>
                  <% Next %>
                </select>
              </div>
            </div>
            <div class="form-actions">
              <button type="submit" name="git" value="kaydet" class="btn btn-secondary">Dönemi Kaydet</button>
              <button type="submit" name="git" value="listeler" class="btn btn-primary">Listeleri Yönet</button>
              <button type="submit" name="git" value="gecmis" class="btn btn-secondary">Geçmiş Dönemi Gör</button>
            </div>
          </form>
        </div>

        <div class="dashboard-card">
          <h2>İstatistikler</h2>
          <p>Yıllık ve aylık yükleme, güncelleme kayıtlarını inceleyin.</p>
          <a href="istatistikler.asp" class="btn btn-primary">İstatistiklere Git</a>
        </div>
      </div>
    </div>
  </div>
</body>
</html>
