<!-- #include file="ayarlar.asp" -->
<!-- #include file="database/connection.asp" -->
<%
Dim dbPath, fso, mesajlar
Set fso = Server.CreateObject("Scripting.FileSystemObject")
dbPath = AccessDbFizikselYol()

mesajlar = ""
If fso.FileExists(dbPath) Then
    mesajlar = mesajlar & "<li>Access veritabanı oluşturuldu: database/nobet_liste.mdb</li>"
Else
    mesajlar = mesajlar & "<li style='color:red;'>Veritabanı dosyası oluşturulamadı.</li>"
End If

If AccessTabloVarMi("NobetListeDosyalar") Then
    mesajlar = mesajlar & "<li>NobetListeDosyalar tablosu hazır</li>"
End If

If AccessTabloVarMi("NobetAdminKullanicilar") Then
    mesajlar = mesajlar & "<li>NobetAdminKullanicilar tablosu hazır</li>"
    mesajlar = mesajlar & "<li>Varsayılan giriş: admin / admin123</li>"
End If

Set fso = Nothing
%>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <title>Nöbet Liste Kurulum</title>
  <link rel="stylesheet" href="assets/style.css">
</head>
<body>
  <div class="admin-wrap">
    <h1 style="color:#850303;">Kurulum Tamamlandı</h1>
    <ul><%= mesajlar %></ul>
    <p>
      <a class="btn btn-primary" href="admin/login.asp">Yönetici Girişi</a>
      <a class="btn btn-secondary" href="index.asp">Liste Sayfası</a>
    </p>
  </div>
</body>
</html>
