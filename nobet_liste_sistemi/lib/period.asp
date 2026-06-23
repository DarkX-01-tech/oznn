<!-- #include file="cookies.asp" -->
<%
Function TurkceAyAdi(ayNumarasi)
    Dim aylar(12)
    aylar(1)  = "Ocak"
    aylar(2)  = "Şubat"
    aylar(3)  = "Mart"
    aylar(4)  = "Nisan"
    aylar(5)  = "Mayıs"
    aylar(6)  = "Haziran"
    aylar(7)  = "Temmuz"
    aylar(8)  = "Ağustos"
    aylar(9)  = "Eylül"
    aylar(10) = "Ekim"
    aylar(11) = "Kasım"
    aylar(12) = "Aralık"
    TurkceAyAdi = aylar(ayNumarasi)
End Function

Function GuncelYil()
    GuncelYil = Year(Now())
End Function

Function GuncelAyKlasor()
    GuncelAyKlasor = LCase(TurkceAyAdi(Month(Now())))
End Function

Function AyKlasorFromNumara(ayNumarasi)
    AyKlasorFromNumara = LCase(TurkceAyAdi(ayNumarasi))
End Function

Function AyBaslikFromKlasor(ayKlasor)
    Dim i
    AyBaslikFromKlasor = ayKlasor
    For i = 1 To 12
        If LCase(AyKlasorFromNumara(i)) = LCase(ayKlasor) Then
            AyBaslikFromKlasor = TurkceAyAdi(i)
            Exit Function
        End If
    Next
End Function

Function AyNumarasiFromKlasor(ayKlasor)
    Dim i
    AyNumarasiFromKlasor = 0
    For i = 1 To 12
        If LCase(AyKlasorFromNumara(i)) = LCase(ayKlasor) Then
            AyNumarasiFromKlasor = i
            Exit Function
        End If
    Next
End Function

Function AyKlasorGecerliMi(ayKlasor)
    AyKlasorGecerliMi = (AyNumarasiFromKlasor(ayKlasor) > 0)
End Function

Function ProjeMinAyForYil(yil)
    If CInt(yil) = PROJE_BASLANGIC_YIL Then
        ProjeMinAyForYil = PROJE_BASLANGIC_AY
    Else
        ProjeMinAyForYil = 1
    End If
End Function

Function DonemSecimGecerliMi(yil, ayKlasor)
    Dim y, ayNo
    If Not IsNumeric(yil) Or Not AyKlasorGecerliMi(ayKlasor) Then
        DonemSecimGecerliMi = False
        Exit Function
    End If
    y = CInt(yil)
    ayNo = AyNumarasiFromKlasor(ayKlasor)
    If y < PROJE_BASLANGIC_YIL Or y > PROJE_BITIS_YIL Then
        DonemSecimGecerliMi = False
        Exit Function
    End If
    If ayNo < ProjeMinAyForYil(y) Then
        DonemSecimGecerliMi = False
        Exit Function
    End If
    DonemSecimGecerliMi = True
End Function

Function DonemYilSinirla(yil)
    Dim y
    y = CInt(yil)
    If y < PROJE_BASLANGIC_YIL Then y = PROJE_BASLANGIC_YIL
    If y > PROJE_BITIS_YIL Then y = PROJE_BITIS_YIL
    DonemYilSinirla = y
End Function

Function DonemAySinirla(yil, ayKlasor)
    Dim ayNo, minAy
    ayNo = AyNumarasiFromKlasor(ayKlasor)
    minAy = ProjeMinAyForYil(yil)
    If ayNo < minAy Then
        DonemAySinirla = AyKlasorFromNumara(minAy)
    Else
        DonemAySinirla = LCase(ayKlasor)
    End If
End Function

Sub DonemKaydet(yil, ayKlasor)
    Dim y, ay
    y = DonemYilSinirla(yil)
    ay = DonemAySinirla(y, ayKlasor)
    NobetCookieYaz SESSION_SECILI_YIL, CStr(y)
    NobetCookieYaz SESSION_SECILI_AY, ay
End Sub

Sub DonemFormIsle()
    Dim yil, ayKlasor
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
        yil = Trim(Request.Form("yil"))
        ayKlasor = LCase(Trim(Request.Form("ay")))
        If DonemSecimGecerliMi(yil, ayKlasor) Then
            DonemKaydet yil, ayKlasor
        End If
    End If
End Sub

Function GetSeciliYil()
    Dim y
    If NobetCookieOku(SESSION_SECILI_YIL) <> "" Then
        y = CInt(NobetCookieOku(SESSION_SECILI_YIL))
    Else
        y = GuncelYil()
    End If
    GetSeciliYil = DonemYilSinirla(y)
End Function

Function GetSeciliAyKlasor()
    Dim ayKlasor, yil
    yil = GetSeciliYil()
    If NobetCookieOku(SESSION_SECILI_AY) <> "" And AyKlasorGecerliMi(NobetCookieOku(SESSION_SECILI_AY)) Then
        ayKlasor = LCase(NobetCookieOku(SESSION_SECILI_AY))
    Else
        ayKlasor = GuncelAyKlasor()
    End If
    GetSeciliAyKlasor = DonemAySinirla(yil, ayKlasor)
End Function

Function GetSeciliAyBaslik()
    GetSeciliAyBaslik = AyBaslikFromKlasor(GetSeciliAyKlasor())
End Function

Function GetSeciliDonemBaslik()
    GetSeciliDonemBaslik = GetSeciliAyBaslik() & " " & GetSeciliYil()
End Function

Function DonemSorguAl(yilParam, ayParam, ByRef yil, ByRef ayKlasor)
    yil = GetSeciliYil()
    ayKlasor = GetSeciliAyKlasor()

    If yilParam <> "" And IsNumeric(yilParam) Then
        yil = CInt(yilParam)
    End If
    If ayParam <> "" And AyKlasorGecerliMi(ayParam) Then
        ayKlasor = LCase(ayParam)
    End If
End Function

Function GetYil()
    GetYil = GuncelYil()
End Function

Function GetAyKlasorAdi()
    GetAyKlasorAdi = GuncelAyKlasor()
End Function

Function GetAyBaslikMetni()
    GetAyBaslikMetni = TurkceAyAdi(Month(Now())) & " " & GuncelYil()
End Function
%>
