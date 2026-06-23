<%
' Nöbet listesi: ay klasörü, dosya kontrolü ve link üretimi

Const NOBET_LISTE_WEB_YOLU = "/Admin/nobet_liste_sistemi/"

Function NobetListeBasePhysicalPath()
    NobetListeBasePhysicalPath = Server.MapPath(NOBET_LISTE_WEB_YOLU)
End Function

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

Function GetAyKlasorAdi()
    Dim ayAdi
    ayAdi = TurkceAyAdi(Month(Now()))
    GetAyKlasorAdi = LCase(ayAdi) & " " & Year(Now())
End Function

Function GetAyBaslikMetni()
    GetAyBaslikMetni = TurkceAyAdi(Month(Now())) & " " & Year(Now())
End Function

Sub EnsureAyKlasoru()
    Dim fso, klasorAdi, tamYol, anaKlasor

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    anaKlasor = NobetListeBasePhysicalPath()

    If Not fso.FolderExists(anaKlasor) Then
        fso.CreateFolder anaKlasor
    End If

    klasorAdi = GetAyKlasorAdi()
    tamYol = anaKlasor & "\" & klasorAdi

    If Not fso.FolderExists(tamYol) Then
        fso.CreateFolder tamYol
    End If

    Set fso = Nothing
End Sub

Function NobetDosyaFizikselYolu(dosyaAdi)
    NobetDosyaFizikselYolu = NobetListeBasePhysicalPath() & "\" & GetAyKlasorAdi() & "\" & dosyaAdi
End Function

Function NobetDosyaMevcut(dosyaAdi)
    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    NobetDosyaMevcut = fso.FileExists(NobetDosyaFizikselYolu(dosyaAdi))
    Set fso = Nothing
End Function

Function NobetDosyaWebYolu(dosyaAdi)
    Dim klasorAdi
    klasorAdi = GetAyKlasorAdi()
    NobetDosyaWebYolu = NOBET_LISTE_WEB_YOLU & Server.URLEncode(klasorAdi) & "/" & Server.URLEncode(dosyaAdi)
End Function

Function NobetDosyaDbKaydiVar(dosyaAdi)
    On Error Resume Next
    NobetDosyaDbKaydiVar = False

    If IsObject(conn) Then
        Dim rs, sql, ayKlasor
        ayKlasor = GetAyKlasorAdi()
        sql = "SELECT TOP 1 1 AS VarMi FROM NobetListeDosyalar " & _
              "WHERE ay_klasor = '" & Replace(ayKlasor, "'", "''") & "' " & _
              "AND dosya_adi = '" & Replace(dosyaAdi, "'", "''") & "' " & _
              "AND aktif = 1"

        Set rs = conn.Execute(sql)
        If Not rs.EOF Then
            NobetDosyaDbKaydiVar = True
        End If
        rs.Close
        Set rs = Nothing
    End If

    On Error GoTo 0
End Function

Function NobetDosyaAktif(dosyaAdi)
    ' Dosya fiziksel olarak mevcutsa link aktif olur.
    ' İsteğe bağlı: veritabanı kaydı da kontrol edilebilir.
    NobetDosyaAktif = NobetDosyaMevcut(dosyaAdi)
End Function

Sub NobetDosyaDbSenkronize(dosyaAdi)
    On Error Resume Next

    If Not IsObject(conn) Then Exit Sub

    Dim sql, ayKlasor, mevcut, aktifDeger
    ayKlasor = GetAyKlasorAdi()
    mevcut = NobetDosyaMevcut(dosyaAdi)

    If mevcut Then
        aktifDeger = "1"
    Else
        aktifDeger = "0"
    End If

    sql = "IF EXISTS (SELECT 1 FROM NobetListeDosyalar " & _
          "WHERE ay_klasor = '" & Replace(ayKlasor, "'", "''") & "' " & _
          "AND dosya_adi = '" & Replace(dosyaAdi, "'", "''") & "') " & _
          "UPDATE NobetListeDosyalar SET aktif = " & aktifDeger & ", " & _
          "guncelleme_tarihi = GETDATE() " & _
          "WHERE ay_klasor = '" & Replace(ayKlasor, "'", "''") & "' " & _
          "AND dosya_adi = '" & Replace(dosyaAdi, "'", "''") & "' " & _
          "ELSE INSERT INTO NobetListeDosyalar (ay_klasor, dosya_adi, aktif) " & _
          "VALUES ('" & Replace(ayKlasor, "'", "''") & "', '" & Replace(dosyaAdi, "'", "''") & "', " & _
          aktifDeger & ")"

    conn.Execute sql
    On Error GoTo 0
End Sub

Sub RenderNobetListeSatiri(baslik, dosyaAdi)
    Dim aktif, webYolu
    aktif = NobetDosyaAktif(dosyaAdi)
    NobetDosyaDbSenkronize dosyaAdi

    Response.Write "<tr>" & vbCrLf
    Response.Write "  <td class=""yazi-stil"">" & vbCrLf

    If aktif Then
        webYolu = NobetDosyaWebYolu(dosyaAdi)
        Response.Write "    <a class=""duyuru-link"" target=""_blank"" href=""" & webYolu & """>" & Server.HTMLEncode(baslik) & "</a>" & vbCrLf
    Else
        Response.Write "    <span class=""duyuru-link-pasif"" title=""Bu ay için dosya henüz yüklenmedi."">" & Server.HTMLEncode(baslik) & "</span>" & vbCrLf
    End If

    Response.Write "  </td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
End Sub
%>
