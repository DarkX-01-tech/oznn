<!-- #include file="../ayarlar.asp" -->
<!-- #include file="../database/connection.asp" -->
<!-- #include file="../lib/auth.asp" -->
<%
If Session(SESSION_ADMIN_KEY) = "1" Then
    Response.Redirect "panel.asp"
    Response.End
End If

Dim hata, kullanici, sifre
hata = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    kullanici = Trim(Request.Form("kullanici"))
    sifre = Trim(Request.Form("sifre"))

    If kullanici = "" Or sifre = "" Then
        hata = "Kullanıcı adı ve şifre zorunludur."
    ElseIf AdminGirisYap(kullanici, sifre) Then
        Response.Redirect "panel.asp"
        Response.End
    Else
        hata = "Geçersiz kullanıcı adı veya şifre."
    End If
End If
%>
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
  <title>Nöbet Liste Yönetimi - Giriş</title>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body>
  <div class="login-box">
    <h1>Yönetici Girişi</h1>
    <% If hata <> "" Then %>
      <div class="alert alert-error"><%= Server.HTMLEncode(hata) %></div>
    <% End If %>
    <form method="post" action="login.asp">
      <div class="form-group">
        <label for="kullanici">Kullanıcı Adı</label>
        <input type="text" id="kullanici" name="kullanici" value="<%= Server.HTMLEncode(kullanici) %>" required>
      </div>
      <div class="form-group">
        <label for="sifre">Şifre</label>
        <input type="password" id="sifre" name="sifre" required>
      </div>
      <button type="submit" class="btn btn-primary" style="width:100%;">Giriş Yap</button>
    </form>
    <p style="text-align:center;margin-top:16px;">
      <a href="../index.asp">Liste Sayfasına Dön</a>
    </p>
  </div>
</body>
</html>
