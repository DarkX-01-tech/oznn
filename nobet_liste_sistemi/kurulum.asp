<%@ Language=VBScript CodePage=65001 %>
<!-- #include file="ayarlar.asp" -->
<!-- #include file="database/connection.asp" -->
<%
Dim dbPath, fso, mesajlar
Set fso = Server.CreateObject("Scripting.FileSystemObject")
dbPath = AccessDbFizikselYol()
mesajlar = ""

If fso.FileExists(dbPath) Then
    mesajlar = mesajlar & "<li>Access veritabanı hazır: database/nobet_liste.mdb</li>"
Else
    mesajlar = mesajlar & "<li class='alert-error'>Veritabanı oluşturulamadı. IIS yazma iznini kontrol edin.</li>"
End If

If AccessTabloVarMi("NobetListeDosyalar") Then
    mesajlar = mesajlar & "<li>NobetListeDosyalar tablosu oluşturuldu</li>"
End If

If AccessTabloVarMi("NobetAdminKullanicilar") Then
    mesajlar = mesajlar & "<li>NobetAdminKullanicilar tablosu oluşturuldu</li>"
    mesajlar = mesajlar & "<li>Varsayılan giriş: <strong>admin</strong> / <strong>admin123</strong></li>"
End If

Set fso = Nothing
%>
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Nöbet Liste Kurulum</title>
  <link rel="stylesheet" href="assets/style.css">
</head>
<body>
  <div class="admin-wrap">
    <h1 style="color:#850303;">Kurulum Tamamlandı</h1>
    <ul class="kurulum-list"><%= mesajlar %></ul>
    <p>
      <a class="btn btn-primary" href="admin/login.asp">Yönetici Girişi</a>
      <a class="btn btn-secondary" href="index.asp">Liste Sayfası</a>
    </p>
  </div>
</body>
</html>
