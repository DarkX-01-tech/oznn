<%
' Development login stub. Replace with production auth on IIS deployments.
If Request.Form("kullanici") <> "" Then
    Session("yemek_admin_giris") = "OK"
    Session("yemek_admin_kullanici") = Request.Form("kullanici")
    Response.Redirect Request.QueryString("return") 
    If Request.QueryString("return") = "" Then Response.Redirect "panel.asp"
End If
%>
<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="utf-8">
  <title>Yemek Paneli - Giriş (Dev)</title>
  <style>
    body { font-family: Segoe UI, sans-serif; background: #f0f9fa; display:flex; align-items:center; justify-content:center; min-height:100vh; margin:0; }
    form { background:#fff; padding:2rem; border-radius:8px; box-shadow:0 4px 16px rgba(0,0,0,.08); width:320px; }
    h1 { margin:0 0 1rem; font-size:1.25rem; color:#2e8b91; }
    label { display:block; margin-bottom:.5rem; }
    input { width:100%; padding:.5rem; margin-bottom:1rem; box-sizing:border-box; }
    button { width:100%; padding:.6rem; background:#45b8c3; color:#fff; border:0; border-radius:4px; cursor:pointer; }
    p.note { font-size:.85rem; color:#666; margin-top:1rem; }
  </style>
</head>
<body>
  <form method="post">
    <h1>Dev Giriş</h1>
    <label for="kullanici">Kullanıcı adı</label>
    <input id="kullanici" name="kullanici" value="admin" required>
    <button type="submit">Giriş yap</button>
    <p class="note">Geliştirme ortamı: herhangi bir kullanıcı adı kabul edilir.</p>
  </form>
</body>
</html>
