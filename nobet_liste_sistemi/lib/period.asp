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

Sub DonemKaydet(yil, ayKlasor)
    Session(SESSION_SECILI_YIL) = CInt(yil)
    Session(SESSION_SECILI_AY) = LCase(ayKlasor)
End Sub

Sub DonemFormIsle()
    Dim yil, ayKlasor
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
        yil = Trim(Request.Form("yil"))
        ayKlasor = LCase(Trim(Request.Form("ay")))
        If yil <> "" And AyKlasorGecerliMi(ayKlasor) Then
            DonemKaydet yil, ayKlasor
        End If
    End If
End Sub

Function GetSeciliYil()
    If Session(SESSION_SECILI_YIL) <> "" Then
        GetSeciliYil = CInt(Session(SESSION_SECILI_YIL))
    Else
        GetSeciliYil = GuncelYil()
    End If
End Function

Function GetSeciliAyKlasor()
    If Session(SESSION_SECILI_AY) <> "" And AyKlasorGecerliMi(Session(SESSION_SECILI_AY)) Then
        GetSeciliAyKlasor = LCase(Session(SESSION_SECILI_AY))
    Else
        GetSeciliAyKlasor = GuncelAyKlasor()
    End If
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
