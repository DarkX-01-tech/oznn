<!-- #include file="database/yemek_auth.asp" -->
<%
'====================================
' Session Kontrolü
'====================================
' Oturum zaten varsa panele yönlendir
If Session("yemek_admin_giris") = "OK" Then
    Response.Redirect "panel.asp"
    Response.End
End If

'====================================
' Form İşlemleri
'====================================
Dim hata_mesaji
hata_mesaji = ""

' Form gönderildiyse
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim kullanici_adi, sifre, gorunen_ad, kullanici_rol
    kullanici_adi = Trim(Request.Form("username"))
    sifre = Trim(Request.Form("password"))

    If YemekAuthDogrula(kullanici_adi, sifre, gorunen_ad, kullanici_rol) Then
        Session("yemek_admin_giris") = "OK"
        Session("yemek_admin_kullanici") = gorunen_ad
        Session("yemek_admin_rol") = kullanici_rol
        Session.Timeout = 120

        Response.Redirect "panel.asp"
        Response.End
    Else
        hata_mesaji = "Kullanıcı adı veya şifre hatalı!"
    End If
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Language" content="tr">
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1254" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css" />
    <title>Yönetici Girişi - Yemek Sistemi</title>
    <link rel="icon" href="../../images/hastane_portal_logo.png" type="image/png">
    <style>
        :root {
            --main-bg-color: #45b8c3;
            --hover-bg-color: #2e8b91;
            --main-text-color: #ffffff;
            --reset-bg-color: #ff4d4d;
            --reset-hover-bg-color: #d43d3d;
            --button-radius: 8px;
            --button-padding: 12px;
            --button-font-size: 18px;
            --transition-speed: 0.3s;
        }

        body {
            margin: 0;
            padding: 0;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: url('../../images/hospital_background.jpg') no-repeat center center fixed;
            background-size: cover;
            height: 100vh;
            overflow: hidden;
            display: flex;
            justify-content: center;
            align-items: center;
        }

        .login-container {
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            backdrop-filter: blur(8px);
        }

        .login-box {
            background-color: rgba(255, 255, 255, 0.96);
            padding: 24px 18px 28px 18px;
            border-radius: 12px;
            box-shadow: 0 8px 22px rgba(0, 0, 0, 0.15);
            width: 340px;
            text-align: center;
            position: relative;
            max-width: 95%;
            transition: box-shadow 0.3s ease;
        }

        .login-box:hover {
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.3);
        }

        .login-box img {
            width: auto;
            height: 110px;
            margin-bottom: 10px;
            animation: rotate-logo 5s ease-in-out infinite;
            user-select: none;
            pointer-events: none;
        }

        @keyframes rotate-logo {
            0% { transform: rotateY(0deg); }
            50% { transform: rotateY(180deg); }
            100% { transform: rotateY(360deg); }
        }

        .hospital-name,
        .login-header {
            user-select: none;
        }

        .error-message {
            background: #fee;
            color: #c33;
            padding: 12px;
            border-radius: 8px;
            margin-bottom: 18px;
            text-align: center;
            border: 1px solid #fcc;
            animation: shake 0.5s;
        }

        @keyframes shake {
            0%, 100% { transform: translateX(0); }
            25% { transform: translateX(-10px); }
            75% { transform: translateX(10px); }
        }

        input[type="text"],
        input[type="password"] {
            width: calc(100% - 32px);
            padding: 15px;
            margin: 10px 0;
            border: 1px solid #ddd;
            border-radius: 8px;
            box-shadow: inset 0 1px 2px rgba(0, 0, 0, 0.1);
            transition: border 0.3s ease;
            font-size: 16px;
            background-color: #f9f9f9;
        }

        input[type="text"]:focus,
        input[type="password"]:focus {
            border-color: var(--main-bg-color);
            outline: none;
            box-shadow: 0 0 8px rgba(69, 184, 195, 0.3);
        }

        .buttons {
            display: flex;
            justify-content: space-between;
        }

        button[type="submit"],
        button[type="reset"],
        button[type="button"] {
            width: 30%;
            padding: var(--button-padding);
            background-color: var(--main-bg-color);
            color: var(--main-text-color);
            border: none;
            border-radius: var(--button-radius);
            cursor: pointer;
            margin: 10px 5px;
            font-size: var(--button-font-size);
            transition: background-color var(--transition-speed), transform var(--transition-speed), box-shadow var(--transition-speed);
            display: flex;
            justify-content: center;
            align-items: center;
            text-transform: uppercase;
            letter-spacing: 1px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }

        button[type="submit"]:hover,
        button[type="reset"]:hover,
        button[type="button"]:hover {
            background-color: var(--hover-bg-color);
            transform: translateY(-3px);
            box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2);
        }

        button[type="reset"] {
            background-color: var(--reset-bg-color);
        }

        button[type="reset"]:hover {
            background-color: var(--reset-hover-bg-color);
        }

        button[type="button"] i,
        button[type="submit"] i,
        button[type="reset"] i {
            font-size: 24px;
            transition: none;
        }

        .hospital-name {
            font-size: 18px;
            color: var(--main-bg-color);
            margin-bottom: 15px;
            font-weight: bold;
            text-transform: uppercase;
            position: relative;
            padding-bottom: 8px;
        }

        .login-header {
            font-size: 16px;
            color: var(--main-bg-color);
            margin-bottom: 24px;
            text-transform: uppercase;
            font-weight: 600;
        }

        .hospital-name::after {
            content: '';
            position: absolute;
            width: 0;
            height: 4px;
            background-color: var(--main-bg-color);
            left: 50%;
            bottom: -5px;
            transition: width 0.5s ease-in-out, left 0.5s ease-in-out;
        }

        .hospital-name:hover::after {
            width: 100%;
            left: 0;
        }

        @media (max-width: 768px) {
            .login-box {
                width: 90%;
                padding: 20px;
            }

            .buttons {
                flex-direction: column;
            }

            button[type="submit"],
            button[type="reset"],
            button[type="button"] {
                width: 100%;
                margin-top: 15px;
            }
        }
    </style>
</head>
<body>
<div class="login-container">
    <div class="login-box">
        <img src="../../images/hastane_yeni_logo.png" alt="Hastane Logo" />
        <div class="hospital-name">Marmara Üniversitesi <br /> Pendik Eğitim ve Araştırma Hastanesi</div>
        <div class="login-header">Yemek Sistemi <br>Yönetici Girişi</div>

        <% If hata_mesaji <> "" Then %>
            <div class="error-message"><%= hata_mesaji %></div>
        <% End If %>

        <form method="post" action="giris.asp">
            <input type="text" name="username" placeholder="Kullanıcı Adı" required autofocus>
            <input type="password" name="password" placeholder="Parola" required>
            <div class="buttons">
                <button type="submit" title="Giriş Yap"><i class="fas fa-sign-in-alt"></i></button>
                <button type="reset" title="Temizle"><i class="fas fa-undo"></i></button>
                <button type="button" title="Ana Sayfa" onclick="window.location.href='/anasayfa_new.asp'"><i class="fas fa-home"></i></button>
            </div>
        </form>
    </div>
</div>
</body>
</html>
