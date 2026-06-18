<%@ Language="VBScript" %>
<!DOCTYPE html>
<html>
<%
session("ok") = false
%>
<!-- #include file="admin/database/Connection.asp" -->
<!-- #include file="ayarlar.asp" -->
<!-- #include file="duyuru-icerik-render.inc" -->
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Language" content="tr">
<title>MÜ Pendik E.A.H. Portal</title>
<link rel="icon" href="images/hastane_portal_logo.png"/>

<style type="text/css">
    @font-face {
        font-family: 'Open Sans';
        src: url('/webfonts/OpenSans-Regular.ttf') format('truetype');
        font-weight: 400;
        font-style: normal;
    }
    @font-face {
        font-family: 'Open Sans';
        src: url('/webfonts/OpenSans-SemiBold.ttf') format('truetype');
        font-weight: 600;
        font-style: normal;
    }
    @font-face {
        font-family: 'Open Sans';
        src: url('/webfonts/OpenSans-ExtraBold.ttf') format('truetype');
        font-weight: 800;
        font-style: normal;
    }

    body {
        background: url('<%= bgresim %>') no-repeat center center fixed;
        background-size: cover;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        margin: 0;
    }

    .golgeliKutu {
        box-shadow: 0 8px 16px rgba(0,0,0,0.2);
        width: 900px;
        margin: 40px auto;
        background-color: #ffffff;
        border-radius: 8px;
        overflow: hidden;
    }

    .baslik {
        font-size: 24px;
        color: #850303;
        text-align: center;
        margin: 25px 0;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
        padding-bottom: 12px;
        text-shadow: 1px 1px 4px rgba(0,0,0,0.3);
    }
    .baslik::after {
        content: "";
        display: block;
        width: 725px;
        height: 2px;
        background-color: #343a40;
        margin-top: 15px;
        margin-left: auto;
        margin-right: auto;
    }

    .duyuru-baslik {
        font-size: 16px;
        color: #333;
        margin-bottom: 5px;
        font-weight: 600;
        border-bottom: 2px solid #850303;
        padding-bottom: 4px;
        display: inline-block;
    }

    .duyuru-icerik {
        font-size: 12px;
        color: #555;
        font-weight: 600;
        line-height: 1.6;
        margin-bottom: 10px;
        text-align: justify;
        overflow-wrap: break-word;
        word-break: break-word;
        white-space: normal;
    }
    .duyuru-icerik b,
    .duyuru-icerik strong  { font-weight: 800 !important; }

    <!--#include file="duyuru-tablo.css.inc"-->

    .baslik b, .baslik strong,
    .duyuru-baslik b, .duyuru-baslik strong { font-weight: inherit !important; }

    .duyuru-tarih {
        font-size: 12px;
        color: #888;
        text-align: right;
        padding: 1em 0;
    }

    a {
        text-decoration: none;
        color: #1e8b99;
        font-weight: 600;
    }
    a:hover { color: #0d5f6b; }

    .search-container {
        float: right;
        margin-right: 20px;
        margin-top: -20px;
    }
    .search-container input[type="text"] {
        padding: 10px 15px;
        border: 1px solid #ccc;
        border-radius: 30px;
        font-size: 16px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
    }

    #scrollTopBtn {
        position: fixed;
        bottom: 30px;
        right: 30px;
        width: 50px;
        height: 50px;
        background-color: #25abb9;
        color: #fff;
        font-size: 24px;
        text-align: center;
        line-height: 50px;
        border-radius: 50%;
        cursor: pointer;
        opacity: 0;
        visibility: hidden;
        z-index: 999;
    }

    #apDiv1 {
        position: fixed;
        left: 2%;
        top: 2%;
        color: #BF0A2D;
        font-weight: bold;
        background: rgba(255,255,255,0.8);
        padding: 10px;
        border-radius: 5px;
        z-index: 1;
    }

    .duyuru-container {
        max-height: 1220px;
        overflow-y: auto;
        padding-right: 10px;
        margin: 10px auto;
        max-width: 710px;
        width: 100%;
        text-align: left;
        padding: 10px;
        background-color: #ffffff;
        display: flex;
        flex-direction: column;
        gap: 10px;
    }
    .duyuru-container::-webkit-scrollbar { width:6px; }
    .duyuru-container::-webkit-scrollbar-thumb { background:#45b8c3; border-radius:5px; }

    #noMatchesRow td {
        font-size: 16px;
        font-weight: 700;
        color: #c00;
        text-align: center;
        padding: 25px 0;
    }

    .image-wrapper { position: relative; display: block; margin: 10px 0; }
    .image-tooltip {
        position: absolute;
        bottom: 20px;
        left: 50%;
        transform: translateX(-50%);
        background: rgba(69, 184, 195, 0.95);
        color: #fff;
        padding: 6px 12px;
        border-radius: 15px;
        font-size: 11px;
        opacity: 0;
        pointer-events: none;
        z-index: 100;
    }
    .image-wrapper:hover .image-tooltip { opacity: 1; }

    .duyuru-icerik img {
        max-width: 100%;
        width: 100%;
        height: auto;
        display: block;
        border-radius: 8px;
        margin: 10px auto;
        cursor: pointer;
        border: 2px solid #f0f0f0;
        box-sizing: border-box;
    }

    .image-modal {
        display: none;
        position: fixed;
        z-index: 10000;
        left: 0; top: 0;
        width: 100%; height: 100%;
        background-color: rgba(0, 0, 0, 0.92);
        opacity: 0;
        transition: opacity 0.3s ease;
    }
    .image-modal.show { opacity: 1; }
    .modal-content {
        position: absolute;
        top: 50%; left: 50%;
        border-radius: 8px;
        opacity: 0;
        cursor: grab;
    }
    .image-modal.show .modal-content {
        transform: translate(-50%, -50%) scale(1);
        opacity: 1;
    }
    .modal-close {
        position: absolute;
        top: 20px; right: 30px;
        color: #fff;
        font-size: 36px;
        cursor: pointer;
        z-index: 10001;
    }
    .modal-prev, .modal-next {
        position: absolute;
        top: 50%;
        width: 50px; height: 50px;
        margin-top: -25px;
        color: white;
        font-size: 26px;
        border-radius: 50%;
        background: rgba(69, 184, 195, 0.9);
        display: none;
        align-items: center;
        justify-content: center;
        cursor: pointer;
    }
    .modal-next { right: 25px; }
    .modal-prev { left: 25px; }
    .zoom-controls {
        position: absolute;
        bottom: 30px;
        left: 50%;
        transform: translateX(-50%);
        display: none;
        gap: 10px;
        z-index: 10002;
    }
    .zoom-btn {
        width: 40px; height: 40px;
        background: rgba(69, 184, 195, 0.9);
        color: white;
        border-radius: 50%;
        border: none;
        cursor: pointer;
    }
    .zoom-indicator {
        position: absolute;
        top: 30px;
        left: 50%;
        transform: translateX(-50%);
        background: rgba(69, 184, 195, 0.9);
        color: white;
        padding: 6px 15px;
        border-radius: 15px;
        font-size: 13px;
        opacity: 0;
        z-index: 10002;
    }
    .zoom-indicator.show { opacity: 1; }

    .image-toggle-link {
        display: inline-block;
        color: #1e8b99;
        font-size: 13px;
        font-weight: 600;
        margin: 8px 0;
        padding: 6px 10px;
        cursor: pointer;
        border-left: 3px solid #1e8b99;
        background: linear-gradient(90deg, rgba(30,139,153,0.12) 0%, transparent 100%);
        border-radius: 0 6px 6px 0;
    }
    .image-content {
        max-height: 0;
        overflow: hidden;
        opacity: 0;
        margin: 0;
        transition: max-height 0.35s ease, opacity 0.25s ease, margin 0.25s ease;
    }
    .image-content.show {
        max-height: 1200px;
        opacity: 1;
        margin: 10px 0 5px 0;
    }
</style>

<script src="/js/duyuru-modal.js" type="text/javascript"></script>
<script src="/js/duyuru-search.js" type="text/javascript"></script>
</head>

<body>

<div id="imageModal" class="image-modal">
    <span class="modal-close" onclick="closeModal()">&times;</span>
    <div id="zoomControls" class="zoom-controls">
        <button class="zoom-btn zoom-out" onclick="zoomOut()">-</button>
        <button class="zoom-btn zoom-reset" onclick="resetZoom()">0</button>
        <button class="zoom-btn zoom-in" onclick="zoomIn()">+</button>
    </div>
    <div id="zoomIndicator" class="zoom-indicator">100%</div>
    <img class="modal-content" id="modalImage" alt="">
    <a class="modal-prev" id="modalPrev" onclick="changeImage(-1)">&#10094;</a>
    <a class="modal-next" id="modalNext" onclick="changeImage(1)">&#10095;</a>
</div>

<div id="scrollTopBtn" onclick="scrollToTop()">^</div>

<div align="center" class="golgeliKutu">
    <table id="Table_01" width="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td colspan="3"><img src="images/muhst_06.png" width="900" height="165" alt=""></td>
        </tr>
        <tr>
            <td bgcolor="#FFFFFF" width="166" height="604" valign="top">
                <!--#include file="sol.asp"-->
            </td>
            <td bgcolor="#FFFFFF" width="734" height="604" valign="top" align="center">

                <div class="baslik">TÜM DUYURULAR</div>

                <div class="search-container">
                    <form>
                        <input type="text" id="search-bar" placeholder="Duyuru Ara...">
                    </form>
                </div>

                <div class="duyuru-container">
                    <table border="0" width="100%" id="table2" class="announcements-table" style="border-collapse: collapse">
                        <%
                        Dim searchTerm
                        searchTerm = Trim(Request("q"))

                        Dim sqld, rsDuy, imageCounter
                        imageCounter = 0
                        Set rsDuy = Server.CreateObject("ADODB.RecordSet")

                        If searchTerm <> "" Then
                            sqld = "SELECT TOP 850 * FROM Sanat_Duyuru " & _
                                   "WHERE (strd_baslik LIKE '%" & searchTerm & "%' " & _
                                   "   OR strd_duyuru LIKE '%" & searchTerm & "%') " & _
                                   "  AND strd_tarih <= Now() " & _
                                   "ORDER BY strd_tarih DESC, strd_id DESC"
                        Else
                            sqld = "SELECT TOP 850 * FROM Sanat_Duyuru " & _
                                   "WHERE strd_tarih <= Now() " & _
                                   "ORDER BY strd_tarih DESC, strd_id DESC"
                        End If

                        rsDuy.Open sqld, conn, 1, 3

                        Do While Not rsDuy.EOF
                            Dim baslik, icerik, tarihVal
                            baslik = rsDuy("strd_baslik") & ""
                            icerik = rsDuy("strd_duyuru") & ""

                            If Not IsNull(rsDuy("strd_tarih")) Then tarihVal = rsDuy("strd_tarih") Else tarihVal = ""

                            If Trim(baslik) <> "" Then
                                Response.Write "<tr><td class='duyuru-baslik'>" & baslik & "</td></tr>"
                            Else
                                Response.Write "<tr><td style='height:5px;'></td></tr>"
                            End If

                            Dim linkIcerikVal, listeStiliVal, tumIcerik, linkSatirlar, linkHtml, satir, parcalar, metin, url, tekUrl
                            linkIcerikVal = rsDuy("strd_link_icerik") & ""
                            listeStiliVal = rsDuy("strd_liste_stili") & ""
                            If Len(listeStiliVal) = 0 Then listeStiliVal = "disc"
                            tumIcerik = icerik

                            If Len(Trim(linkIcerikVal)) > 0 Then
                                linkSatirlar = Split(linkIcerikVal, vbCrLf)
                                linkHtml = "<ul style='list-style-type: " & listeStiliVal & "; padding-left: 25px; margin: 10px 0;'>"
                                For Each satir In linkSatirlar
                                    satir = Trim(satir)
                                    If Len(satir) > 0 Then
                                        If InStr(satir, "|") > 0 Then
                                            parcalar = Split(satir, "|")
                                            metin = Trim(parcalar(0))
                                            url = Trim(parcalar(1))
                                            If Left(LCase(url), 4) <> "http" And Left(LCase(url), 4) <> "pdf/" Then url = "pdf/" & url
                                            linkHtml = linkHtml & "<li><a target='_blank' href='" & url & "'>" & metin & "</a></li>"
                                        Else
                                            If Left(LCase(satir), 4) = "http" Or Left(LCase(satir), 4) = "pdf/" Or InStr(satir, ".") > 0 Then
                                                tekUrl = satir
                                                If Left(LCase(tekUrl), 4) <> "http" And Left(LCase(tekUrl), 4) <> "pdf/" Then tekUrl = "pdf/" & tekUrl
                                                linkHtml = linkHtml & "<li><a target='_blank' href='" & tekUrl & "'>" & satir & "</a></li>"
                                            Else
                                                linkHtml = linkHtml & "<li>" & satir & "</li>"
                                            End If
                                        End If
                                    End If
                                Next
                                linkHtml = linkHtml & "</ul>"
                                tumIcerik = tumIcerik & linkHtml
                            End If
                            icerik = tumIcerik

                            Dim yayinTarihi, gunFarki, gorselleriGoster, ozelGunMu, gorsellerKapaliMi
                            gunFarki = 999
                            ozelGunMu = False
                            gorsellerKapaliMi = False
                            If Not IsNull(rsDuy("strd_tarih")) Then
                                yayinTarihi = rsDuy("strd_tarih")
                                gunFarki = DateDiff("d", yayinTarihi, Now())
                            End If
                            If Not IsNull(rsDuy("strd_ozel_gun")) Then ozelGunMu = rsDuy("strd_ozel_gun")
                            If Not IsNull(rsDuy("strd_gorsel_kapali")) Then gorsellerKapaliMi = rsDuy("strd_gorsel_kapali")

                            If gorsellerKapaliMi Then
                                gorselleriGoster = False
                            ElseIf ozelGunMu Then
                                gorselleriGoster = True
                            Else
                                gorselleriGoster = (gunFarki < 3)
                            End If

                            If Trim(icerik) <> "" Then
                                Call YazDuyuruIcerik(icerik, gorsellerKapaliMi, gorselleriGoster, imageCounter)
                            Else
                                Response.Write "<tr><td style='height:5px;'></td></tr>"
                            End If

                            If Len(tarihVal) > 0 Then
                                Response.Write "<tr><td class='duyuru-tarih'>" & _
                                    Right("0" & Day(tarihVal), 2) & "/" & _
                                    Right("0" & Month(tarihVal), 2) & "/" & _
                                    Year(tarihVal) & "</td></tr>"
                            Else
                                Response.Write "<tr><td style='height:5px;'></td></tr>"
                            End If

                            rsDuy.MoveNext
                        Loop
                        rsDuy.Close
                        Set rsDuy = Nothing
                        %>

                        <tr id="noMatchesRow" style="display:none;">
                            <td>Sonuç bulunamadı.</td>
                        </tr>
                    </table>
                </div>

            </td>
        </tr>
    </table>
</div>

<div id="apDiv1">
<%
Dim ArrayIPLocalStart(2), ArrayIPLocalEnd(2), ArrayIPClient, IPClient, blnLocal, i
Ipclient = Request.ServerVariables("REMOTE_ADDR")
Response.Write "Lokal IP <br>Adresiniz:<br>  " & Ipclient & "<BR>"
blnLocal = False
ArrayIPLocalStart(0) = "010.000.000.000"
ArrayIPLocalEnd(0)   = "010.255.255.255"
ArrayIPLocalStart(1) = "172.016.000.000"
ArrayIPLocalEnd(1)   = "172.031.000.000"
ArrayIPLocalStart(2) = "192.168.000.000"
ArrayIPLocalEnd(2)   = "192.168.255.000"
ArrayIPClient = Split(Ipclient,".")
For i = LBound(ArrayIPClient) To UBound(ArrayIPClient)
    ArrayIPClient(i) = String(3 - Len(ArrayIPClient(i)), "0") & ArrayIPClient(i)
Next
IPClient = Join(ArrayIPClient, "")
If Trim(Ipclient) <> "" Then
    For i = LBound(ArrayIPLocalStart) To UBound(ArrayIPLocalStart)
        ArrayIPLocalStart(i) = Replace(ArrayIPLocalStart(i),".","")
        ArrayIPLocalEnd(i)   = Replace(ArrayIPLocalEnd(i),".","")
        If IPClient >= ArrayIPLocalStart(i) And IPClient <= ArrayIPLocalEnd(i) Then
            blnLocal = True
            Exit For
        End If
    Next
End If
Response.Write "**********"
%>
</div>
</body>
</html>
