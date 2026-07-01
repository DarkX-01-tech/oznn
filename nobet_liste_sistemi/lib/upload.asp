<%
Function DosyaYukleVeKaydet(hedefTamYol, ByRef hataMesaji)
    Dim upload, dosya, bilesenAdi, basarili, tmpKlasor
    basarili = False
    hataMesaji = ""
    tmpKlasor = ModulFizikselYol() & "\tmp"

    On Error Resume Next

    Set upload = Server.CreateObject("Persits.Upload")
    bilesenAdi = "Persits.Upload"
    If Err.Number <> 0 Then
        Err.Clear
        Set upload = Server.CreateObject("ABCUpload4.XForm")
        bilesenAdi = "ABCUpload4.XForm"
    End If

    If Err.Number = 0 And Not upload Is Nothing Then
        If bilesenAdi = "Persits.Upload" Then
            upload.OverwriteFiles = True
            upload.Save tmpKlasor
            Set dosya = upload.Files("dosya")
            If dosya Is Nothing Then
                If upload.Files.Count > 0 Then
                    Set dosya = upload.Files(1)
                End If
            End If
            If Not dosya Is Nothing Then
                dosya.SaveAs hedefTamYol
                If Err.Number = 0 Then basarili = True
            Else
                hataMesaji = "Yüklenecek dosya seçilmedi."
            End If
        Else
            upload.AbsolutePath = True
            upload.Overwrite = True
            If upload("dosya").Value <> "" Then
                upload("dosya").Save hedefTamYol
                If Err.Number = 0 Then basarili = True
            Else
                hataMesaji = "Yüklenecek dosya seçilmedi."
            End If
        End If
    Else
        hataMesaji = "Sunucuda dosya yükleme bileşeni bulunamadı. Persits.Upload veya ABCUpload kurulumu gerekir."
    End If

    If Err.Number <> 0 And hataMesaji = "" Then
        hataMesaji = "Yükleme hatası: " & Err.Description
    End If

    On Error GoTo 0
    DosyaYukleVeKaydet = basarili
End Function
%>
