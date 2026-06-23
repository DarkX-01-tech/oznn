<!-- #include file="period.asp" -->
<%
Function BosDizi()
    Dim d(0)
    d(0) = vbNullString
    BosDizi = d
End Function

Function DiziDoluMu(dizi)
    If Not IsArray(dizi) Then
        DiziDoluMu = False
    ElseIf UBound(dizi) = 0 And dizi(0) = vbNullString Then
        DiziDoluMu = False
    Else
        DiziDoluMu = True
    End If
End Function

Function AyKlasordeDosyaVarMi(yil, ayKlasor)
    Dim fso, binalar, i, klasorYolu, dosya
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    binalar = Array(BINA_PENDIK, BINA_BASIBUYUK)
    AyKlasordeDosyaVarMi = False

    For i = 0 To 1
        klasorYolu = AyKlasorFizikselYolu(binalar(i), yil, ayKlasor)
        If fso.FolderExists(klasorYolu) Then
            For Each dosya In fso.GetFolder(klasorYolu).Files
                AyKlasordeDosyaVarMi = True
                Exit Function
            Next
        End If
    Next
End Function

Function AyGecmisDonemMi(yil, ayKlasor)
    Dim ayNo
    ayNo = AyNumarasiFromKlasor(ayKlasor)
    If ayNo = 0 Then
        AyGecmisDonemMi = False
        Exit Function
    End If
    If CInt(yil) > GuncelYil() Then
        AyGecmisDonemMi = False
        Exit Function
    End If
    If CInt(yil) < GuncelYil() Then
        AyGecmisDonemMi = True
        Exit Function
    End If
    AyGecmisDonemMi = (ayNo < Month(Now()))
End Function

Function IstatistikYillariDizisi()
    Dim yillar, rs, sql, yilDeger, sonuc(), i, fso, yilKlasor
    Set yillar = Server.CreateObject("Scripting.Dictionary")

    On Error Resume Next
    If IsObject(conn) Then
        sql = "SELECT DISTINCT yil FROM NobetListeIslemLog ORDER BY yil DESC"
        Set rs = conn.Execute(sql)
        Do While Not rs.EOF
            yilDeger = CStr(rs("yil"))
            If Not yillar.Exists(yilDeger) Then yillar.Add yilDeger, True
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(ListelerKokYolu()) Then
        For Each yilKlasor In fso.GetFolder(ListelerKokYolu()).SubFolders
            yilDeger = yilKlasor.Name
            If IsNumeric(yilDeger) Then
                If Not yillar.Exists(yilDeger) Then yillar.Add yilDeger, True
            End If
        Next
    End If

    If yillar.Count = 0 Then
        ReDim sonuc(0)
        sonuc(0) = GuncelYil()
    Else
        ReDim sonuc(yillar.Count - 1)
        i = 0
        For Each yilDeger In yillar.Keys
            sonuc(i) = CInt(yilDeger)
            i = i + 1
        Next
    End If

    IstatistikYillariDizisi = sonuc
End Function

Function YukluAylarForYil(yil)
    Dim sonuc(), i, sayac, ayKlasor
    sayac = 0
    ReDim sonuc(11)

    For i = 1 To 12
        ayKlasor = AyKlasorFromNumara(i)
        If AyKlasordeDosyaVarMi(yil, ayKlasor) Then
            sonuc(sayac) = ayKlasor
            sayac = sayac + 1
        End If
    Next

    If sayac = 0 Then
        YukluAylarForYil = BosDizi()
    Else
        ReDim Preserve sonuc(sayac - 1)
        YukluAylarForYil = sonuc
    End If
End Function

Function GecmisYillariDizisi()
    Dim yillar, i, aylar, sonuc(), sayac, yilDeger
    yillar = IstatistikYillariDizisi()
    sayac = 0
    ReDim sonuc(UBound(yillar))

    For i = 0 To UBound(yillar)
        yilDeger = yillar(i)
        aylar = GecmisAylarForYil(yilDeger)
        If DiziDoluMu(aylar) Then
            sonuc(sayac) = yilDeger
            sayac = sayac + 1
        End If
    Next

    If sayac = 0 Then
        GecmisYillariDizisi = BosDizi()
    Else
        ReDim Preserve sonuc(sayac - 1)
        GecmisYillariDizisi = sonuc
    End If
End Function

Function GecmisAylarForYil(yil)
    Dim sonuc(), i, sayac, ayKlasor
    sayac = 0
    ReDim sonuc(11)

    For i = 1 To 12
        ayKlasor = AyKlasorFromNumara(i)
        If AyGecmisDonemMi(yil, ayKlasor) And AyKlasordeDosyaVarMi(yil, ayKlasor) Then
            sonuc(sayac) = ayKlasor
            sayac = sayac + 1
        End If
    Next

    If sayac = 0 Then
        GecmisAylarForYil = BosDizi()
    Else
        ReDim Preserve sonuc(sayac - 1)
        GecmisAylarForYil = sonuc
    End If
End Function

Function DosyaKayitBilgisi(yil, binaKodu, ayKlasor, dosyaAdi)
    Dim rs, sql, bilgi(3)
    bilgi(0) = ""
    bilgi(1) = 0
    bilgi(2) = ""

    On Error Resume Next
    If Not IsObject(conn) Then
        DosyaKayitBilgisi = bilgi
        Exit Function
    End If

    sql = "SELECT TOP 1 olusturma_tarihi, guncelleme_sayisi, guncelleme_tarihi FROM NobetListeDosyalar " & _
          "WHERE yil = " & CInt(yil) & " AND bina = '" & SqlEscape(binaKodu) & "' AND ay_klasor = '" & SqlEscape(ayKlasor) & _
          "' AND dosya_adi = '" & SqlEscape(dosyaAdi) & "'"

    Set rs = conn.Execute(sql)
    If Not rs.EOF Then
        If Not IsNull(rs("olusturma_tarihi")) Then bilgi(0) = rs("olusturma_tarihi")
        If Not IsNull(rs("guncelleme_sayisi")) Then bilgi(1) = rs("guncelleme_sayisi")
        If Not IsNull(rs("guncelleme_tarihi")) Then bilgi(2) = rs("guncelleme_tarihi")
    End If
    rs.Close
    Set rs = Nothing
    On Error GoTo 0

    DosyaKayitBilgisi = bilgi
End Function

Function TarihGoster(tarihDeger)
    If IsDate(tarihDeger) Then
        TarihGoster = FormatDateTime(tarihDeger, 0)
    Else
        TarihGoster = "-"
    End If
End Function

Sub RenderIstatistikTamListe(yil, ayKlasor)
    Dim i, satir, baslik, dosyaAdi, bilgi, aktif, webYolu, donemBaslik
    donemBaslik = AyBaslikFromKlasor(ayKlasor) & " " & yil

    Response.Write "<div class=""section-title"">Pendik E.A.H. Nöbet Listeleri</div>"
    Response.Write "<div class=""ay-baslik"">" & Server.HTMLEncode(donemBaslik) & "</div>"
    Response.Write "<table class=""liste-tablo"">"
    For i = 0 To UBound(pendikNobetListeleri)
        satir = pendikNobetListeleri(i)
        RenderIstatistikListeSatiri BINA_PENDIK, satir(0), satir(1), yil, ayKlasor
    Next
    Response.Write "</table>"

    Response.Write "<div class=""section-title"">Prof. Dr. Asaf Ataseven Ek Hizmet Binası</div>"
    Response.Write "<div class=""ay-baslik"">" & Server.HTMLEncode(donemBaslik) & "</div>"
    Response.Write "<table class=""liste-tablo"">"
    For i = 0 To UBound(basibuyukNobetListeleri)
        satir = basibuyukNobetListeleri(i)
        RenderIstatistikListeSatiri BINA_BASIBUYUK, satir(0), satir(1), yil, ayKlasor
    Next
    Response.Write "</table>"
End Sub

Sub RenderIstatistikListeSatiri(binaKodu, baslik, dosyaAdi, yil, ayKlasor)
    Dim aktif, webYolu, bilgi
    aktif = NobetDosyaAktif(binaKodu, dosyaAdi, yil, ayKlasor)
    bilgi = DosyaKayitBilgisi(yil, binaKodu, ayKlasor, dosyaAdi)

    Response.Write "<tr><td class=""yazi-stil"">"
    If aktif Then
        webYolu = NobetDosyaWebYolu(binaKodu, dosyaAdi, yil, ayKlasor)
        Response.Write "<a class=""duyuru-link"" target=""_blank"" href=""" & webYolu & """>" & Server.HTMLEncode(baslik) & "</a>"
    Else
        Response.Write "<span class=""duyuru-link-pasif"">" & Server.HTMLEncode(baslik) & "</span>"
    End If
    Response.Write "<div class=""stat-meta"">"
    Response.Write "İlk yükleme: " & TarihGoster(bilgi(0))
    Response.Write " &nbsp;|&nbsp; Güncelleme: " & bilgi(1)
    Response.Write " &nbsp;|&nbsp; Son: " & TarihGoster(bilgi(2))
    Response.Write "</div></td></tr>"
End Sub
%>
