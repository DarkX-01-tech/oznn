<%@ Language=VBScript CodePage=65001 %>
<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../database/connection.asp" -->
<!-- #include file="../lib/config.asp" -->
<!-- #include file="../lib/functions.asp" -->
<!-- #include file="../lib/auth.asp" -->
<!-- #include file="../lib/ui.asp" -->
<%
AdminGirisGerekli

Dim i, seciliYil, minAy

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim yil, ayKlasor
    yil = Trim(Request.Form("yil"))
    ayKlasor = LCase(Trim(Request.Form("ay")))
    If DonemSecimGecerliMi(yil, ayKlasor) Then
        DonemKaydet yil, ayKlasor
        Response.Redirect "listeler.asp"
        Response.End
    End If
End If

seciliYil = GetSeciliYil()
minAy = ProjeMinAyForYil(seciliYil)
%>
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="utf-8">
  <meta http-equiv="Content-Language" content="tr">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Dönem Seçimi</title>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body class="admin-body">
  <div class="page-shell">
    <div class="admin-wrap">
      <div class="admin-topbar">
        <div>
          <p class="eyebrow">Yükleme / Güncelleme</p>
          <h1>Dönem Seçimi</h1>
        </div>
        <% AdminUstLinkler %>
      </div>

      <div class="dashboard-card">
        <p>Listeleri yönetmek için yıl ve ay seçin. Proje dönemi: Haziran 2026 — 2040.</p>
        <form method="post" action="donem_sec.asp" class="period-form">
          <div class="form-row">
            <div class="form-group">
              <label for="yil">Yıl</label>
              <select id="yil" name="yil" required>
                <% For i = PROJE_BITIS_YIL To PROJE_BASLANGIC_YIL Step -1 %>
                  <option value="<%= i %>"<% If i = seciliYil Then %> selected<% End If %>><%= i %></option>
                <% Next %>
              </select>
            </div>
            <div class="form-group">
              <label for="ay">Ay</label>
              <select id="ay" name="ay" required>
                <% For i = minAy To 12 %>
                  <option value="<%= AyKlasorFromNumara(i) %>"<% If AyKlasorFromNumara(i) = GetSeciliAyKlasor() Then %> selected<% End If %>><%= TurkceAyAdi(i) %></option>
                <% Next %>
              </select>
            </div>
          </div>
          <div class="form-actions">
            <button type="submit" class="btn btn-primary">Listeleri Yönet</button>
            <a href="panel.asp" class="btn btn-secondary">Geri</a>
          </div>
        </form>
      </div>
    </div>
  </div>
  <script>
    (function() {
      var ayAdlari = ["Ocak","Şubat","Mart","Nisan","Mayıs","Haziran","Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"];
      var ayKlasorleri = ["ocak","subat","mart","nisan","mayis","haziran","temmuz","agustos","eylul","ekim","kasim","aralik"];
      var baslangicYil = <%= PROJE_BASLANGIC_YIL %>;
      var baslangicAy = <%= PROJE_BASLANGIC_AY %>;
      var yilSelect = document.getElementById("yil");
      var aySelect = document.getElementById("ay");

      function guncelleAylar() {
        var yil = parseInt(yilSelect.value, 10);
        var minAy = (yil === baslangicYil) ? baslangicAy : 1;
        var secili = aySelect.value;
        aySelect.innerHTML = "";
        for (var i = minAy; i <= 12; i++) {
          var opt = document.createElement("option");
          opt.value = ayKlasorleri[i - 1];
          opt.textContent = ayAdlari[i - 1];
          aySelect.appendChild(opt);
        }
        for (var j = 0; j < aySelect.options.length; j++) {
          if (aySelect.options[j].value === secili) {
            aySelect.selectedIndex = j;
            return;
          }
        }
        aySelect.selectedIndex = 0;
      }

      yilSelect.addEventListener("change", guncelleAylar);
    })();
  </script>
</body>
</html>
