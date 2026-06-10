<%
' Yemek Sistemi yonetici giris bilgileri
' Sunucuya deploy sonrasi bu dosyadaki sifreleri degistirin.

Function YemekAuthDogrula(kullanici_adi, sifre, ByRef gorunen_ad, ByRef rol)
    YemekAuthDogrula = False
    gorunen_ad = ""
    rol = ""

    If kullanici_adi = "ys_yonetici_mk2026" And sifre = "Yk@7mP#x9Qz2!vL4" Then
        gorunen_ad = "Direk Yönetici"
        rol = "yonetici"
        YemekAuthDogrula = True
    ElseIf kullanici_adi = "diyetisyen" And sifre = "Dt#5nK@w8Rf3!mP1" Then
        gorunen_ad = "Diyetisyen"
        rol = "diyetisyen"
        YemekAuthDogrula = True
    End If
End Function
%>
