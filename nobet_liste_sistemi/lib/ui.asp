<%
Function FlashMesajAl()
    Select Case LCase(Trim(Request("mesaj")))
        Case "yuklendi"
            FlashMesajAl = "Dosya başarıyla yüklendi. Liste linki aktif hale geldi."
        Case "guncellendi"
            FlashMesajAl = "Dosya başarıyla güncellendi."
        Case "silindi"
            FlashMesajAl = "Dosya silindi. Liste linki pasif hale geldi."
        Case Else
            FlashMesajAl = ""
    End Select
End Function

Sub FlashMesajGoster()
    Dim mesaj
    mesaj = FlashMesajAl()
    If mesaj <> "" Then
        Response.Write "<div class=""alert alert-success flash-alert"">" & Server.HTMLEncode(mesaj) & "</div>"
    End If
End Sub
%>
