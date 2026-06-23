<!-- #include file="period.asp" -->
<%
Sub KlasorOlustur(fso, yol)
    If Not fso.FolderExists(yol) Then
        fso.CreateFolder yol
    End If
End Sub

Function AyKlasorFizikselYolu(binaKodu, yil, ayKlasor)
    AyKlasorFizikselYolu = ListelerKokYolu() & "\" & yil & "\" & binaKodu & "\" & ayKlasor
End Function

Sub EnsureAyKlasoru(binaKodu, yil, ayKlasor)
    Dim fso, yol
    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    yol = ListelerKokYolu()
    KlasorOlustur fso, yol
    yol = yol & "\" & yil
    KlasorOlustur fso, yol
    yol = yol & "\" & binaKodu
    KlasorOlustur fso, yol
    yol = yol & "\" & ayKlasor
    KlasorOlustur fso, yol

    Set fso = Nothing
End Sub

Sub EnsureTumAyKlasorleri()
    EnsureAyKlasoru BINA_PENDIK, GuncelYil(), GuncelAyKlasor()
    EnsureAyKlasoru BINA_BASIBUYUK, GuncelYil(), GuncelAyKlasor()
End Sub

Sub EnsureSeciliDonemKlasorleri()
    EnsureAyKlasoru BINA_PENDIK, GetSeciliYil(), GetSeciliAyKlasor()
    EnsureAyKlasoru BINA_BASIBUYUK, GetSeciliYil(), GetSeciliAyKlasor()
End Sub

Function NobetDosyaFizikselYolu(binaKodu, dosyaAdi, yil, ayKlasor)
    NobetDosyaFizikselYolu = AyKlasorFizikselYolu(binaKodu, yil, ayKlasor) & "\" & dosyaAdi
End Function

Function NobetDosyaMevcut(binaKodu, dosyaAdi, yil, ayKlasor)
    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    NobetDosyaMevcut = fso.FileExists(NobetDosyaFizikselYolu(binaKodu, dosyaAdi, yil, ayKlasor))
    Set fso = Nothing
End Function

Function NobetDosyaWebYolu(binaKodu, dosyaAdi, yil, ayKlasor)
    NobetDosyaWebYolu = MODUL_WEB_YOLU & LISTELER_KLASORU & "/" & yil & "/" & binaKodu & "/" & ayKlasor & "/" & Server.URLEncode(dosyaAdi)
End Function

Function SqlBool(deger)
    If deger Then
        SqlBool = "True"
    Else
        SqlBool = "False"
    End If
End Function

Sub NobetDosyaDbKaydet(binaKodu, dosyaAdi, baslik, yukleyen, aktif, yil, ayKlasor, islemTipi)
    On Error Resume Next
    If Not IsObject(conn) Then Exit Sub

    Dim rs, sql, mevcutKayit, yeniSayi
    Set rs = conn.Execute("SELECT id, guncelleme_sayisi FROM NobetListeDosyalar WHERE yil = " & CInt(yil) & _
          " AND bina = '" & SqlEscape(binaKodu) & "' AND ay_klasor = '" & SqlEscape(ayKlasor) & _
          "' AND dosya_adi = '" & SqlEscape(dosyaAdi) & "'")

    mevcutKayit = Not rs.EOF

    If mevcutKayit Then
        If islemTipi = "guncelle" And aktif Then
            yeniSayi = 1
            If Not IsNull(rs("guncelleme_sayisi")) Then
                yeniSayi = CLng(rs("guncelleme_sayisi")) + 1
            End If
            sql = "UPDATE NobetListeDosyalar SET aktif = True, baslik = '" & SqlEscape(baslik) & _
                  "', yukleyen = '" & SqlEscape(yukleyen) & "', guncelleme_tarihi = Now(), guncelleme_sayisi = " & yeniSayi & _
                  " WHERE yil = " & CInt(yil) & " AND bina = '" & SqlEscape(binaKodu) & "' AND ay_klasor = '" & SqlEscape(ayKlasor) & _
                  "' AND dosya_adi = '" & SqlEscape(dosyaAdi) & "'"
            conn.Execute sql
        ElseIf islemTipi = "sil" Then
            sql = "UPDATE NobetListeDosyalar SET aktif = False, yukleyen = '" & SqlEscape(yukleyen) & _
                  "', guncelleme_tarihi = Now() WHERE yil = " & CInt(yil) & " AND bina = '" & SqlEscape(binaKodu) & _
                  "' AND ay_klasor = '" & SqlEscape(ayKlasor) & "' AND dosya_adi = '" & SqlEscape(dosyaAdi) & "'"
            conn.Execute sql
        Else
            sql = "UPDATE NobetListeDosyalar SET aktif = " & SqlBool(aktif) & ", baslik = '" & SqlEscape(baslik) & _
                  "', yukleyen = '" & SqlEscape(yukleyen) & "', guncelleme_tarihi = Now() " & _
                  "WHERE yil = " & CInt(yil) & " AND bina = '" & SqlEscape(binaKodu) & "' AND ay_klasor = '" & SqlEscape(ayKlasor) & _
                  "' AND dosya_adi = '" & SqlEscape(dosyaAdi) & "'"
            conn.Execute sql
        End If
    ElseIf aktif Then
        sql = "INSERT INTO NobetListeDosyalar (yil, bina, ay_klasor, dosya_adi, baslik, aktif, yukleyen, olusturma_tarihi, guncelleme_sayisi) VALUES (" & _
              CInt(yil) & ", '" & SqlEscape(binaKodu) & "', '" & SqlEscape(ayKlasor) & "', '" & SqlEscape(dosyaAdi) & _
              "', '" & SqlEscape(baslik) & "', True, '" & SqlEscape(yukleyen) & "', Now(), 0)"
        conn.Execute sql
    End If

    If islemTipi <> "" Then
        LogNobetIslem yil, binaKodu, ayKlasor, dosyaAdi, baslik, islemTipi, yukleyen
    End If

    rs.Close
    Set rs = Nothing
    On Error GoTo 0
End Sub

Sub NobetDosyaDbSenkronize(binaKodu, dosyaAdi, baslik, yil, ayKlasor)
    Dim mevcut, yukleyen
    mevcut = NobetDosyaMevcut(binaKodu, dosyaAdi, yil, ayKlasor)
    yukleyen = ""
    If NobetCookieOku(SESSION_ADMIN_KEY & "_ad") <> "" Then
        yukleyen = NobetCookieOku(SESSION_ADMIN_KEY & "_ad")
    End If
    NobetDosyaDbKaydet binaKodu, dosyaAdi, baslik, yukleyen, mevcut, yil, ayKlasor, ""
End Sub

Function NobetDosyaAktif(binaKodu, dosyaAdi, yil, ayKlasor)
    NobetDosyaAktif = NobetDosyaMevcut(binaKodu, dosyaAdi, yil, ayKlasor)
End Function

Sub RenderNobetListeSatiri(binaKodu, baslik, dosyaAdi, yil, ayKlasor)
    Dim aktif, webYolu
    aktif = NobetDosyaAktif(binaKodu, dosyaAdi, yil, ayKlasor)

    Response.Write "<tr class=""nobet-liste-row"">"
    Response.Write "<td class=""nobet-liste-hucre"">"
    Response.Write "<span class=""liste-satir-icerik"">"
    Response.Write "<span class=""liste-icon"" aria-hidden=""true""></span>"
    If aktif Then
        webYolu = NobetDosyaWebYolu(binaKodu, dosyaAdi, yil, ayKlasor)
        Response.Write "<a class=""duyuru-link"" target=""_blank"" href=""" & webYolu & """>" & Server.HTMLEncode(baslik) & "</a>"
    Else
        Response.Write "<span class=""duyuru-link-pasif"" title=""Bu ay için dosya henüz yüklenmedi."">" & Server.HTMLEncode(baslik) & "</span>"
    End If
    Response.Write "</span></td></tr>"
End Sub

Sub RenderNobetListeTablosu(binaKodu, listeDizisi, yil, ayKlasor)
    Dim i, satir
    For i = 0 To UBound(listeDizisi)
        satir = listeDizisi(i)
        RenderNobetListeSatiri binaKodu, satir(0), satir(1), yil, ayKlasor
    Next
End Sub

Sub RenderTamListeGorunumu(yil, ayKlasor)
    Dim donemBaslik
    donemBaslik = AyBaslikFromKlasor(ayKlasor) & " " & yil

    Response.Write "<div class=""baslik"">PENDİK EĞİTİM &amp; ARAŞTIRMA HASTANESİ<br/>NÖBET LİSTELERİ</div>"
    Response.Write "<div class=""ay-baslik"">" & Server.HTMLEncode(donemBaslik) & "</div>"
    Response.Write "<table class=""liste-tablo"">"
    RenderNobetListeTablosu BINA_PENDIK, pendikNobetListeleri, yil, ayKlasor
    Response.Write "</table>"

    Response.Write "<div class=""baslik"">PROF. DR. ASAF ATASEVEN EK HİZMET BİNASI<br/>NÖBET LİSTELERİ</div>"
    Response.Write "<div class=""ay-baslik"">" & Server.HTMLEncode(donemBaslik) & "</div>"
    Response.Write "<table class=""liste-tablo"">"
    RenderNobetListeTablosu BINA_BASIBUYUK, basibuyukNobetListeleri, yil, ayKlasor
    Response.Write "<tr><td class=""yazi-stil no-icon""><span class=""duz-metn"">Pacs Destek (0531 682 44 36)</span></td></tr>"
    Response.Write "</table>"
End Sub
%>
