<%
Function FlashMesajAl()
    Select Case LCase(Trim(Request("mesaj")))
        Case "yuklendi"
            FlashMesajAl = "Dosya başarıyla yüklendi. Liste linki aktif hale geldi."
        Case "guncellendi"
            FlashMesajAl = "Dosya başarıyla güncellendi."
        Case "silindi"
            FlashMesajAl = "Dosya silindi. Liste linki pasif hale geldi."
        Case "donem"
            FlashMesajAl = "Dönem seçimi kaydedildi."
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

Sub AdminNavGoster(aktifSayfa)
    Dim navClass
    Response.Write "<div class=""admin-menu"">"
    Response.Write "<a href=""panel.asp"""
    If aktifSayfa = "panel" Then Response.Write " class=""active"""
    Response.Write ">Yönetim</a>"
    Response.Write "<a href=""listeler.asp"""
    If aktifSayfa = "listeler" Then Response.Write " class=""active"""
    Response.Write ">Listeler</a>"
    Response.Write "<a href=""istatistikler.asp"""
    If aktifSayfa = "istatistikler" Then Response.Write " class=""active"""
    Response.Write ">İstatistikler</a>"
    Response.Write "<a href=""gecmis.asp"""
    If aktifSayfa = "gecmis" Then Response.Write " class=""active"""
    Response.Write ">Geçmiş Dönem</a>"
    Response.Write "<a href=""../index.asp"" target=""_blank"">Liste Sayfası</a>"
    Response.Write "<a href=""logout.asp"">Çıkış</a>"
    Response.Write "</div>"
End Sub

Sub DonemOzetKartlariGoster()
    Response.Write "<div class=""summary-grid"">"
    Response.Write "<div class=""summary-card""><span class=""label"">Yıl</span><span class=""value"">" & GetSeciliYil() & "</span></div>"
    Response.Write "<div class=""summary-card""><span class=""label"">Ay</span><span class=""value"">" & Server.HTMLEncode(GetSeciliAyBaslik()) & "</span></div>"
    Response.Write "<div class=""summary-card""><span class=""label"">Dönem</span><span class=""value"">" & Server.HTMLEncode(GetSeciliDonemBaslik()) & "</span></div>"
    Response.Write "</div>"
End Sub
%>
