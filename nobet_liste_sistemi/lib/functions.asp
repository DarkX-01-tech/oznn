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

Function GetYil()
    GetYil = Year(Now())
End Function

Function GetAyKlasorAdi()
    GetAyKlasorAdi = LCase(TurkceAyAdi(Month(Now())))
End Function

Function GetAyBaslikMetni()
    GetAyBaslikMetni = TurkceAyAdi(Month(Now())) & " " & GetYil()
End Function

Sub KlasorOlustur(fso, yol)
    If Not fso.FolderExists(yol) Then
        fso.CreateFolder yol
    End If
End Sub

Function AyKlasorFizikselYolu(binaKodu)
    AyKlasorFizikselYolu = ListelerKokYolu() & "\" & GetYil() & "\" & binaKodu & "\" & GetAyKlasorAdi()
End Function

Sub EnsureAyKlasoru(binaKodu)
    Dim fso, yol
    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    yol = ListelerKokYolu()
    KlasorOlustur fso, yol
    yol = yol & "\" & GetYil()
    KlasorOlustur fso, yol
    yol = yol & "\" & binaKodu
    KlasorOlustur fso, yol
    yol = yol & "\" & GetAyKlasorAdi()
    KlasorOlustur fso, yol

    Set fso = Nothing
End Sub

Sub EnsureTumAyKlasorleri()
    EnsureAyKlasoru BINA_PENDIK
    EnsureAyKlasoru BINA_BASIBUYUK
End Sub

Function NobetDosyaFizikselYolu(binaKodu, dosyaAdi)
    NobetDosyaFizikselYolu = AyKlasorFizikselYolu(binaKodu) & "\" & dosyaAdi
End Function

Function NobetDosyaMevcut(binaKodu, dosyaAdi)
    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    NobetDosyaMevcut = fso.FileExists(NobetDosyaFizikselYolu(binaKodu, dosyaAdi))
    Set fso = Nothing
End Function

Function NobetDosyaWebYolu(binaKodu, dosyaAdi)
    NobetDosyaWebYolu = MODUL_WEB_YOLU & LISTELER_KLASORU & "/" & GetYil() & "/" & binaKodu & "/" & GetAyKlasorAdi() & "/" & Server.URLEncode(dosyaAdi)
End Function

Sub NobetDosyaDbKaydet(binaKodu, dosyaAdi, baslik, yukleyen, aktif)
    On Error Resume Next
    If Not IsObject(conn) Then Exit Sub

    Dim rs, sql, aktifDeger
    If aktif Then
        aktifDeger = "True"
    Else
        aktifDeger = "False"
    End If

    sql = "SELECT id FROM NobetListeDosyalar WHERE yil = " & GetYil() & _
          " AND bina = '" & SqlEscape(binaKodu) & "' AND ay_klasor = '" & SqlEscape(GetAyKlasorAdi()) & _
          "' AND dosya_adi = '" & SqlEscape(dosyaAdi) & "'"

    Set rs = conn.Execute(sql)

    If Not rs.EOF Then
        sql = "UPDATE NobetListeDosyalar SET aktif = " & aktifDeger & ", baslik = '" & SqlEscape(baslik) & _
              "', yukleyen = '" & SqlEscape(yukleyen) & "', guncelleme_tarihi = Now() " & _
              "WHERE yil = " & GetYil() & " AND bina = '" & SqlEscape(binaKodu) & "' AND ay_klasor = '" & SqlEscape(GetAyKlasorAdi()) & _
              "' AND dosya_adi = '" & SqlEscape(dosyaAdi) & "'"
        conn.Execute sql
    Else
        sql = "INSERT INTO NobetListeDosyalar (yil, bina, ay_klasor, dosya_adi, baslik, aktif, yukleyen, olusturma_tarihi) VALUES (" & _
              GetYil() & ", '" & SqlEscape(binaKodu) & "', '" & SqlEscape(GetAyKlasorAdi()) & "', '" & SqlEscape(dosyaAdi) & _
              "', '" & SqlEscape(baslik) & "', " & aktifDeger & ", '" & SqlEscape(yukleyen) & "', Now())"
        conn.Execute sql
    End If

    If IsObject(rs) Then
        rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0
End Sub

Sub NobetDosyaDbSenkronize(binaKodu, dosyaAdi, baslik)
    Dim mevcut, yukleyen
    mevcut = NobetDosyaMevcut(binaKodu, dosyaAdi)
    yukleyen = ""
    If IsObject(Session) Then
        If Session(SESSION_ADMIN_KEY & "_ad") <> "" Then
            yukleyen = Session(SESSION_ADMIN_KEY & "_ad")
        End If
    End If
    NobetDosyaDbKaydet binaKodu, dosyaAdi, baslik, yukleyen, mevcut
End Sub

Function NobetDosyaAktif(binaKodu, dosyaAdi)
    NobetDosyaAktif = NobetDosyaMevcut(binaKodu, dosyaAdi)
End Function

Sub RenderNobetListeSatiri(binaKodu, baslik, dosyaAdi)
    Dim aktif, webYolu
    aktif = NobetDosyaAktif(binaKodu, dosyaAdi)
    NobetDosyaDbSenkronize binaKodu, dosyaAdi, baslik

    Response.Write "<tr>" & vbCrLf
    Response.Write "  <td class=""yazi-stil"">" & vbCrLf

    If aktif Then
        webYolu = NobetDosyaWebYolu(binaKodu, dosyaAdi)
        Response.Write "    <a class=""duyuru-link"" target=""_blank"" href=""" & webYolu & """>" & Server.HTMLEncode(baslik) & "</a>" & vbCrLf
    Else
        Response.Write "    <span class=""duyuru-link-pasif"" title=""Bu ay için dosya henüz yüklenmedi."">" & Server.HTMLEncode(baslik) & "</span>" & vbCrLf
    End If

    Response.Write "  </td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
End Sub

Sub RenderNobetListeTablosu(binaKodu, listeDizisi)
    Dim i, satir
    For i = 0 To UBound(listeDizisi)
        satir = listeDizisi(i)
        RenderNobetListeSatiri binaKodu, satir(0), satir(1)
    Next
End Sub

Function IzinVerilenUzantilar()
    IzinVerilenUzantilar = Array(".pdf", ".xls", ".xlsx")
End Function

Function DosyaUzantisiGecerliMi(dosyaAdi)
    Dim uzantilar, i, uzanti
    uzanti = LCase(Right(dosyaAdi, Len(dosyaAdi) - InStrRev(dosyaAdi, ".")))
    uzantilar = IzinVerilenUzantilar()

    DosyaUzantisiGecerliMi = False
    For i = 0 To UBound(uzantilar)
        If "." & uzanti = LCase(uzantilar(i)) Then
            DosyaUzantisiGecerliMi = True
            Exit Function
        End If
    Next
End Function
%>
