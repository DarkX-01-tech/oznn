<%
' Duyuru içeriğini yazar. Tablolar her zaman görünür; sadece görseller gizlenebilir.
Sub YazDuyuruIcerik(icerik, gorsellerKapaliMi, gorselleriGoster, ByRef imageCounter)
    Dim linkId, contentId, linkId2, contentId2
    Dim textContent, imageContent, imgPos
    Dim icerikteTabloVar

    If Trim(icerik) = "" Then Exit Sub

    icerikteTabloVar = (InStr(LCase(icerik), "<table") > 0)

    If InStr(LCase(icerik), "<img") > 0 Then
        imageCounter = imageCounter + 1
        linkId = "imgLink" & imageCounter
        contentId = "imgContent" & imageCounter
        imgPos = InStr(LCase(icerik), "<img")

        If imgPos > 1 Then
            textContent = Left(icerik, imgPos - 1)
            imageContent = Mid(icerik, imgPos)
        Else
            textContent = ""
            imageContent = icerik
        End If

        Response.Write "<tr><td class='duyuru-icerik'>"
        If Trim(textContent) <> "" Then Response.Write textContent

        If gorselleriGoster Then
            Response.Write imageContent
        Else
            Response.Write "<div class='image-toggle-link' id='" & linkId & "' onclick=""toggleImage('" & linkId & "', '" & contentId & "')"">"
            Response.Write "  <span class='text'>Görseli Açmak İçin Tıklayınız</span>"
            Response.Write "</div>"
            Response.Write "<div class='image-content' id='" & contentId & "'>"
            Response.Write imageContent
            Response.Write "</div>"
        End If
        Response.Write "</td></tr>"
    Else
        Response.Write "<tr><td class='duyuru-icerik'>"
        ' Tablo varsa içerik her zaman doğrudan gösterilir
        If gorsellerKapaliMi And Not icerikteTabloVar Then
            imageCounter = imageCounter + 1
            linkId2 = "imgLink" & imageCounter
            contentId2 = "imgContent" & imageCounter
            Response.Write "<div class='image-toggle-link' id='" & linkId2 & "' onclick=""toggleImage('" & linkId2 & "', '" & contentId2 & "')"">"
            Response.Write "  <span class='text'>İçeriği Açmak İçin Tıklayınız</span>"
            Response.Write "</div>"
            Response.Write "<div class='image-content' id='" & contentId2 & "'>" & icerik & "</div>"
        Else
            Response.Write icerik
        End If
        Response.Write "</td></tr>"
    End If
End Sub
%>
