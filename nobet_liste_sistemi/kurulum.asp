<%@ Language=VBScript CodePage=65001 %>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType = "text/html; charset=utf-8"
%>
<!-- #include file="ayarlar.asp" -->
<!-- #include file="lib/functions.asp" -->
<!-- #include file="database/connection.asp" -->
<%
Dim dbPath, fso, mesajlar
Set fso = Server.CreateObject("Scripting.FileSystemObject")
dbPath = AccessDbFizikselYol()
mesajlar = ""

EnsureAyKlasoru BINA_PENDIK, GuncelYil(), GuncelAyKlasor()
EnsureAyKlasoru BINA_BASIBUYUK, GuncelYil(), GuncelAyKlasor()

If VeritabaniHazirMi() And fso.FileExists(dbPath) Then
    mesajlar = mesajlar & "<li>Access veritabanı hazır: <code>database/nobet_liste.mdb</code></li>"
Else
    mesajlar = mesajlar & "<li><strong>Hata:</strong> Veritabanı oluşturulamadı. IIS yazma iznini kontrol edin."
    If connHata <> "" Then
        mesajlar = mesajlar & "<br><small>" & Server.HTMLEncode(connHata) & "</small>"
    End If
    mesajlar = mesajlar & "</li>"
End If

If AyKlasoruMevcut(BINA_PENDIK, GuncelYil(), GuncelAyKlasor()) Then
    mesajlar = mesajlar & "<li>Liste klasörleri hazır: <code>listeler/" & GuncelYil() & "/...</code></li>"
Else
    mesajlar = mesajlar & "<li><strong>Uyarı:</strong> <code>listeler</code> klasörü oluşturulamadı. IIS yazma iznini kontrol edin.</li>"
End If

If AccessTabloVarMi("NobetListeDosyalar") Then
    mesajlar = mesajlar & "<li><code>NobetListeDosyalar</code> tablosu oluşturuldu</li>"
End If

If AccessTabloVarMi("NobetAdminKullanicilar") Then
    mesajlar = mesajlar & "<li><code>NobetAdminKullanicilar</code> tablosu oluşturuldu</li>"
    mesajlar = mesajlar & "<li>Varsayılan giriş: <strong>admin</strong> / <strong>admin123</strong></li>"
End If

Set fso = Nothing
%>
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta http-equiv="Content-Language" content="tr">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Nöbet Liste Kurulum</title>
  <link rel="stylesheet" href="assets/style.css">
</head>
<body class="admin-body">
  <div class="page-shell">
    <div class="admin-wrap">
      <div class="admin-topbar">
        <div>
          <p class="eyebrow">Kurulum</p>
          <h1>Kurulum Tamamlandı</h1>
        </div>
      </div>

      <div class="alert alert-success">Sistem başarıyla hazırlandı. Aşağıdaki kontrolleri inceleyebilirsiniz.</div>

      <ul class="kurulum-list"><%= mesajlar %></ul>

      <div class="kurulum-actions">
        <a class="btn btn-primary" href="admin/login.asp">Yönetici Girişi</a>
        <a class="btn btn-secondary" href="index.asp">Liste Sayfası</a>
      </div>
    </div>
  </div>
</body>
</html>
