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

Sub AdminUstLinkler()
    Response.Write "<div class=""admin-menu"">"
    Response.Write "<a href=""panel.asp"">Yönetim Paneli</a>"
    Response.Write "<a href=""../index.asp"" target=""_blank"">Liste Sayfası</a>"
    Response.Write "<a href=""logout.asp"">Çıkış</a>"
    Response.Write "</div>"
End Sub

Sub DonemOzetKartlariGoster(yil, ayKlasor)
    Response.Write "<div class=""summary-grid"">"
    Response.Write "<div class=""summary-card""><span class=""label"">Yıl</span><span class=""value"">" & yil & "</span></div>"
    Response.Write "<div class=""summary-card""><span class=""label"">Ay</span><span class=""value"">" & Server.HTMLEncode(AyBaslikFromKlasor(ayKlasor)) & "</span></div>"
    Response.Write "<div class=""summary-card""><span class=""label"">Dönem</span><span class=""value"">" & Server.HTMLEncode(AyBaslikFromKlasor(ayKlasor) & " " & yil) & "</span></div>"
    Response.Write "</div>"
End Sub
%>
