<%@ Language="VBScript" %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>MÜ Pendik E.A.H. Portal</title>
    <link rel="icon" href="images/hastane_portal_logo.png"/>
    <script src="/js/jquery-3.7.1.min.js" type="text/javascript"></script>

    <%
    session("ok") = false
    %>

    <!-- #include file="admin/database/Connection.asp" -->
    <!-- #include file="ayarlar.asp" -->

    <style>
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
          background-color: #fff;
          border-radius: 8px;
          overflow: hidden;
          transition: box-shadow 0.3s, background-color 0.3s;
        }
        .golgeliKutu:hover {
          background-color: #f9f9f9;
          box-shadow: 0 12px 24px rgba(0,0,0,0.3);
        }

        .baslik {
          font-size: 24px;
          color: #850303;
          text-align: center;
          margin: -10px 0;
          font-weight: 800;
          text-transform: uppercase;
          letter-spacing: 1px;
          padding: 10px 0;
          text-shadow: 1px 1px 4px rgba(0,0,0,0.3);
          transition: color 0.3s, border-color 0.3s;
        }
        .baslik:hover { color: #343a40; border-color: #c70039; }
        .baslik::after {
          content: "";
          display: block;
          width: 725px;
          height: 2px;
          background-color: #343a40;
          margin: 0px auto 0 auto;
        }

        .duyuru-baslik {
          font-size: 16px;
          color: #333;
          margin-bottom: 5px;
          font-weight: 600;
          text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
          border-bottom: 2px solid #850303;
          padding-bottom: 4px;
          display: inline-block;
          transition: color 0.3s, border-color 0.3s;
        }
        .duyuru-baslik:hover { color: #850303; border-color: #343a40; }

        .duyuru-icerik {
          font-size: 12px;
          color: #555;
          font-weight: 600;
          line-height: 1.6;
          margin-bottom: 10px;
          text-align: justify;
          text-shadow: 1px 1px 2px rgba(0,0,0,0.05);
          overflow-wrap: break-word;
          word-break: break-word;
          white-space: normal;
        }
        .duyuru-icerik b,
        .duyuru-icerik strong { font-weight: 800 !important; }

        .duyuru-icerik table,
        .duyuru-tablo-always table {
          width: 100% !important;
          max-width: 100% !important;
          border-collapse: collapse !important;
          margin: 12px 0 !important;
          border: 1px solid #e5e5e5 !important;
          text-align: left !important;
          font-family: 'Open Sans', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important;
          font-size: 12px !important;
          line-height: 1.6 !important;
          font-weight: 400 !important;
          color: #555 !important;
          text-shadow: none !important;
          background: #fff !important;
        }
        .duyuru-tablo-always {
          display: block !important;
          visibility: visible !important;
          opacity: 1 !important;
          max-height: none !important;
          overflow: visible !important;
          margin: 12px 0 !important;
        }
        .duyuru-icerik table td,
        .duyuru-icerik table th,
        .duyuru-tablo-always table td,
        .duyuru-tablo-always table th {
          border: none !important;
          border-bottom: 1px solid #e5e5e5 !important;
          padding: 8px 12px !important;
          text-align: left !important;
          font-family: 'Open Sans', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important;
          font-size: 12px !important;
          line-height: 1.6 !important;
          font-weight: 400 !important;
          color: #555 !important;
          vertical-align: middle !important;
          word-break: break-word !important;
          text-shadow: none !important;
          background-color: #ffffff !important;
        }
        .duyuru-icerik table tr:last-child td,
        .duyuru-icerik table tr:last-child th,
        .duyuru-tablo-always table tr:last-child td,
        .duyuru-tablo-always table tr:last-child th {
          border-bottom: none !important;
        }
        .duyuru-icerik table tr:first-child td,
        .duyuru-icerik table tr:first-child th,
        .duyuru-tablo-always table tr:first-child td,
        .duyuru-tablo-always table tr:first-child th {
          background-color: #f2f2f2 !important;
          font-weight: 600 !important;
          text-align: center !important;
          border-bottom: 1px solid #e5e5e5 !important;
        }
        .duyuru-icerik table tr:not(:first-child) td:first-child,
        .duyuru-tablo-always table tr:not(:first-child) td:first-child {
          font-weight: 600 !important;
          width: 42% !important;
        }
        .duyuru-icerik table tr:not(:first-child) td:last-child,
        .duyuru-tablo-always table tr:not(:first-child) td:last-child {
          font-weight: 400 !important;
        }
        .duyuru-icerik table b,
        .duyuru-icerik table strong,
        .duyuru-tablo-always table b,
        .duyuru-tablo-always table strong {
          font-weight: 600 !important;
        }
        .duyuru-icerik .image-content.show table,
        .duyuru-icerik > table,
        .duyuru-icerik table,
        .duyuru-tablo-always table {
          display: table !important;
          visibility: visible !important;
          opacity: 1 !important;
        }

        .duyuru-tarih {
          font-size: 12px;
          color: #888;
          text-align: right;
          text-shadow: 1px 1px 2px rgba(0,0,0,0.05);
          padding-top: 1em;
          padding-bottom: 1em;
        }

        a {
          text-decoration: none;
          color: #1e8b99;
          transition: all 0.25s ease;
          font-weight: 600;
          border-bottom: 1px solid transparent;
        }
        a:hover {
          color: #0d5f6b;
          border-bottom: 1px solid #0d5f6b;
        }

        .duyuru-container {
          max-height: 1330px;
          overflow-y: auto;
          padding-right: 10px;
          margin: 10px auto;
          max-width: 710px;
          width: 100%;
          text-align: left;
          padding: 10px;
          background-color: #fff;
          display: flex;
          flex-direction: column;
          gap: 10px;
        }
        .duyuru-container::-webkit-scrollbar { width: 6px; }
        .duyuru-container::-webkit-scrollbar-track { background: #ffffff; }
        .duyuru-container::-webkit-scrollbar-thumb { background: #45b8c3; border-radius: 5px; }
        .duyuru-container::-webkit-scrollbar-thumb:hover { background: #25abb9; }
        .duyuru-container::-webkit-scrollbar-button { display: none; }

        #apDiv1 {
          position: fixed;
          left: 2%;
          top: 2%;
          width: auto;
          height: auto;
          z-index: 1;
          color: #BF0A2D;
          font-weight: bold;
          background: rgba(255,255,255,0.8);
          padding: 10px;
          border-radius: 5px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.2);
        }

        .navbar {
          width: 100%;
          background-color: #ffffff;
          border-bottom: 1px solid #ddd;
          box-shadow: 2px 0 5px rgba(0,0,0,0.1);
          padding-top: 10px;
          padding-bottom: 10px;
          box-sizing: border-box;
          text-align: center;
        }
        .navbar ul { padding: 0; margin: 0; list-style: none; display: inline-block; }
        .navbar > ul > li { position: relative; display: inline-block; margin-right: 15px; }
        .navbar > ul > li:last-child { margin-right: 0; }
        .navbar > ul > li > a {
          display: block;
          padding: 8px 16px;
          font-size: 14px;
          font-weight: 600;
          color: #fff;
          background: #25abb9;
          border-radius: 15px;
          text-decoration: none;
          box-shadow: 0 4px 8px rgba(37,171,185,0.3);
        }
        .navbar > ul > li > a:hover { background: #1e8c99; }

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
          box-shadow: 0 4px 8px rgba(0,0,0,0.3);
          cursor: pointer;
          opacity: 0;
          visibility: hidden;
          z-index: 999;
        }

        .image-wrapper { position: relative; display: block; margin: 10px 0; }
        .image-tooltip {
          position: absolute;
          bottom: 20px;
          left: 50%;
          transform: translateX(-50%) translateY(10px);
          background: rgba(69, 184, 195, 0.95);
          color: #fff;
          padding: 6px 12px;
          border-radius: 15px;
          font-size: 11px;
          font-weight: 600;
          opacity: 0;
          pointer-events: none;
          transition: all 0.25s ease;
          z-index: 100;
        }
        .image-wrapper:hover .image-tooltip { opacity: 1; transform: translateX(-50%) translateY(0); }

        .duyuru-icerik img {
          max-width: 100%;
          width: 100%;
          height: auto;
          display: block;
          border-radius: 8px;
          box-shadow: 0 3px 10px rgba(0,0,0,0.12);
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
          transform-origin: center center;
          border-radius: 8px;
          box-shadow: 0 15px 50px rgba(0,0,0,0.5);
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
          cursor: pointer;
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
    <script>
        document.addEventListener("DOMContentLoaded", function() {
            setTimeout(function() {
                window.scrollTo({ top: 0, behavior: 'smooth' });
            }, 15000);
        });

        document.addEventListener("DOMContentLoaded", function() {
            var scrollTopBtn = document.getElementById("scrollTopBtn");
            if (!scrollTopBtn) return;
            window.addEventListener("scroll", function() {
                if (window.scrollY > 300) {
                    scrollTopBtn.style.visibility = "visible";
                    scrollTopBtn.style.opacity = "1";
                } else {
                    scrollTopBtn.style.opacity = "0";
                    scrollTopBtn.style.visibility = "hidden";
                }
            });
        });

        function scrollToTop() {
            window.scrollTo({ top: 0, behavior: "smooth" });
        }

        var currentImageIndex = 0;
        var currentAnnouncementImages = [];
        var currentZoom = 1;
        var isDragging = false;
        var startX, startY;
        var translateX = 0, translateY = 0;
        var zoomIndicatorTimeout;

        document.addEventListener("DOMContentLoaded", function() {
            setTimeout(function() {
                attachImageModalEvents();
            }, 500);

            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape') closeModal();
            });

            document.addEventListener('keydown', function(e) {
                var modal = document.getElementById('imageModal');
                if (modal && modal.style.display === 'block') {
                    if (e.key === 'ArrowLeft') changeImage(-1);
                    else if (e.key === 'ArrowRight') changeImage(1);
                    else if (e.key === '+' || e.key === '=') zoomIn();
                    else if (e.key === '-' || e.key === '_') zoomOut();
                    else if (e.key === '0') resetZoom();
                }
            });

            document.addEventListener('wheel', function(e) {
                var modal = document.getElementById('imageModal');
                if (modal && modal.style.display === 'block') {
                    e.preventDefault();
                    if (e.deltaY < 0) zoomIn();
                    else zoomOut();
                }
            }, { passive: false });
        });

        function attachImageModalEvents() {
            document.querySelectorAll('.duyuru-icerik').forEach(function(announcement) {
                var images = announcement.querySelectorAll('img');
                if (images.length === 0) return;

                var announcementImages = Array.from(images);
                images.forEach(function(img, localIndex) {
                    if (!img.parentElement.classList.contains('image-wrapper')) {
                        var wrapper = document.createElement('div');
                        wrapper.className = 'image-wrapper';
                        img.parentNode.insertBefore(wrapper, img);
                        wrapper.appendChild(img);

                        var tooltip = document.createElement('div');
                        tooltip.className = 'image-tooltip';
                        tooltip.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><circle cx="10" cy="10" r="7"/><path d="M15 15L21 21"/></svg><span>Detaylı Görüntüle</span>';
                        wrapper.appendChild(tooltip);
                    }

                    img.onclick = function() {
                        currentAnnouncementImages = announcementImages;
                        openModal(localIndex);
                    };
                });
            });
        }

        function openModal(index) {
            if (currentAnnouncementImages.length === 0) return;
            currentImageIndex = index;
            resetZoom();

            var modal = document.getElementById('imageModal');
            var modalImg = document.getElementById('modalImage');
            var zoomControls = document.getElementById('zoomControls');

            modal.style.display = "block";
            document.body.style.overflow = "hidden";
            setTimeout(function() { modal.classList.add('show'); }, 10);
            modalImg.src = currentAnnouncementImages[index].src;
            modalImg.onload = function() {
                fitImageToScreen();
                zoomControls.style.display = 'flex';
            };
            updateNavigationButtons();
            setupImageDragging();
        }

        function closeModal() {
            var modal = document.getElementById('imageModal');
            var zoomControls = document.getElementById('zoomControls');
            modal.classList.remove('show');
            zoomControls.style.display = 'none';
            setTimeout(function() {
                modal.style.display = "none";
                document.body.style.overflow = "auto";
                resetZoom();
            }, 300);
        }

        function changeImage(direction) {
            currentImageIndex += direction;
            if (currentImageIndex >= currentAnnouncementImages.length) currentImageIndex = 0;
            else if (currentImageIndex < 0) currentImageIndex = currentAnnouncementImages.length - 1;

            var modalImg = document.getElementById('modalImage');
            var modal = document.getElementById('imageModal');
            modal.classList.remove('show');
            resetZoom();
            setTimeout(function() {
                modalImg.src = currentAnnouncementImages[currentImageIndex].src;
                modalImg.onload = function() { fitImageToScreen(); };
                modal.classList.add('show');
            }, 200);
        }

        function updateNavigationButtons() {
            var prevBtn = document.getElementById('modalPrev');
            var nextBtn = document.getElementById('modalNext');
            if (currentAnnouncementImages.length > 1) {
                prevBtn.style.display = 'flex';
                nextBtn.style.display = 'flex';
            } else {
                prevBtn.style.display = 'none';
                nextBtn.style.display = 'none';
            }
        }

        function zoomIn() { if (currentZoom < 5) { currentZoom += 0.25; applyZoom(); } }
        function zoomOut() { if (currentZoom > 0.5) { currentZoom -= 0.25; applyZoom(); } }
        function resetZoom() { currentZoom = 1; translateX = 0; translateY = 0; applyZoom(); }

        function applyZoom() {
            var modalImg = document.getElementById('modalImage');
            if (currentZoom === 1) {
                fitImageToScreen();
                modalImg.classList.remove('zoomed');
            } else {
                modalImg.style.transform = 'translate(calc(-50% + ' + translateX + 'px), calc(-50% + ' + translateY + 'px)) scale(' + currentZoom + ')';
                modalImg.classList.add('zoomed');
            }
            showZoomIndicator();
        }

        function fitImageToScreen() {
            var modalImg = document.getElementById('modalImage');
            var scaleX = (window.innerWidth * 0.9) / modalImg.naturalWidth;
            var scaleY = (window.innerHeight * 0.85) / modalImg.naturalHeight;
            var scale = Math.min(scaleX, scaleY, 1);
            modalImg.style.width = (modalImg.naturalWidth * scale) + 'px';
            modalImg.style.height = (modalImg.naturalHeight * scale) + 'px';
            modalImg.style.transform = 'translate(-50%, -50%) scale(1)';
            translateX = 0;
            translateY = 0;
        }

        function showZoomIndicator() {
            var indicator = document.getElementById('zoomIndicator');
            indicator.textContent = Math.round(currentZoom * 100) + '%';
            indicator.classList.add('show');
            clearTimeout(zoomIndicatorTimeout);
            zoomIndicatorTimeout = setTimeout(function() { indicator.classList.remove('show'); }, 1000);
        }

        function setupImageDragging() {
            var modalImg = document.getElementById('modalImage');
            modalImg.addEventListener('mousedown', startDrag);
            document.addEventListener('mousemove', drag);
            document.addEventListener('mouseup', endDrag);
            modalImg.addEventListener('touchstart', startDrag);
            document.addEventListener('touchmove', drag);
            document.addEventListener('touchend', endDrag);
        }

        function startDrag(e) {
            if (currentZoom <= 1) return;
            isDragging = true;
            if (e.type === 'touchstart') {
                startX = e.touches[0].clientX - translateX;
                startY = e.touches[0].clientY - translateY;
            } else {
                startX = e.clientX - translateX;
                startY = e.clientY - translateY;
                e.preventDefault();
            }
        }

        function drag(e) {
            if (!isDragging) return;
            var modalImg = document.getElementById('modalImage');
            if (e.type === 'touchmove') {
                translateX = e.touches[0].clientX - startX;
                translateY = e.touches[0].clientY - startY;
            } else {
                translateX = e.clientX - startX;
                translateY = e.clientY - startY;
            }
            modalImg.style.transform = 'translate(calc(-50% + ' + translateX + 'px), calc(-50% + ' + translateY + 'px)) scale(' + currentZoom + ')';
        }

        function endDrag() { isDragging = false; }

        window.onclick = function(event) {
            var modal = document.getElementById('imageModal');
            if (event.target === modal) closeModal();
        };

        function toggleImage(linkId, contentId) {
            var link = document.getElementById(linkId);
            var content = document.getElementById(contentId);
            if (!link || !content) return;
            var isOpen = content.classList.contains('show');
            if (isOpen) {
                content.classList.remove('show');
                link.classList.remove('active');
            } else {
                content.classList.add('show');
                link.classList.add('active');
            }
        }
    </script>
</head>
<body>

<div id="imageModal" class="image-modal">
    <span class="modal-close" onclick="closeModal()">&times;</span>
    <div id="zoomControls" class="zoom-controls">
        <button class="zoom-btn zoom-out" onclick="zoomOut()" title="Uzaklaştır (-)">-</button>
        <button class="zoom-btn zoom-reset" onclick="resetZoom()" title="Sıfırla (0)">0</button>
        <button class="zoom-btn zoom-in" onclick="zoomIn()" title="Yakınlaştır (+)">+</button>
    </div>
    <div id="zoomIndicator" class="zoom-indicator">100%</div>
    <img class="modal-content" id="modalImage" alt="">
    <a class="modal-prev" id="modalPrev" onclick="changeImage(-1)">&#10094;</a>
    <a class="modal-next" id="modalNext" onclick="changeImage(1)">&#10095;</a>
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

<div id="scrollTopBtn" onclick="scrollToTop()">^</div>

<div align="center" class="golgeliKutu">
    <table id="Table_01" width="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td colspan="3">
                <img border="0" src="images/muhst_06.png" width="900" height="165" alt="">
            </td>
        </tr>
        <tr>
            <td colspan="3">
                <div class="navbar">
                    <ul>
                        <li><a href="https://dys.saglik.gov.tr/">DYS</a></li>
                        <li><a href="http://10.201.64.42:8077">PAGO</a></li>
                        <li><a href="https://eposta.saglik.gov.tr/owakontrol/">SB E-Posta</a></li>
                        <li><a href="http://10.201.64.71:8040">Pendik HBYS Kur</a></li>
                        <li><a href="http://10.231.96.71:8040">Bsb HBYS Kur</a></li>
                        <li><a href="https://marmaraeah.saglik.gov.tr/">MÜ Pendik E.A.H.</a></li>
                        <li><a href="http://10.210.122.20/default.php">HES Eğitim</a></li>
                    </ul>
                </div>
            </td>
        </tr>
        <tr>
            <td bgcolor="#FFFFFF" width="166" height="604" valign="top">
                <!--#include file="sol.asp"-->
            </td>
            <td bgcolor="#FFFFFF" width="734" height="604" valign="top" align="center">
                <div class="baslik">ANASAYFA</div>
                <div class="duyuru-container">
                    <table border="0" width="100%" id="table2" style="border-collapse: collapse">
                        <%
                        On Error Resume Next
                        Dim rsDuyuru, sqlD, baslik, icerik, tarihVal, imageCounter
                        Set rsDuyuru = Server.CreateObject("ADODB.Recordset")
                        imageCounter = 0

                        sqlD = "SELECT TOP 8 *, " & _
                               "IIF(strd_sabitli = True AND strd_sabit_baslangic <= Now() AND strd_sabit_bitis >= Now(), 1, 0) AS sabit_aktif " & _
                               "FROM Sanat_Duyuru " & _
                               "WHERE strd_tarih <= Now() AND (strd_gizli = 0 OR strd_gizli IS NULL) " & _
                               "ORDER BY IIF(strd_sabitli = True AND strd_sabit_baslangic <= Now() AND strd_sabit_bitis >= Now(), 1, 0) DESC, strd_tarih DESC, strd_id DESC"

                        rsDuyuru.Open sqlD, conn, 1, 3
                        If Err.Number <> 0 Then
                            Response.Write "<tr><td style='color:red;'>Veritabanı hatası: " & Err.Description & "</td></tr>"
                            Response.End
                        End If

                        Function FindTableEnd(s, startPos)
                            Dim depth, pos, nextOpen, nextClose, sLower
                            sLower = LCase(s & "")
                            depth = 1
                            pos = startPos + 6
                            Do While depth > 0
                                nextOpen = InStr(pos, sLower, "<table")
                                nextClose = InStr(pos, sLower, "</table")
                                If nextClose = 0 Then
                                    FindTableEnd = 0
                                    Exit Function
                                End If
                                If nextOpen > 0 And nextOpen < nextClose Then
                                    depth = depth + 1
                                    pos = nextOpen + 6
                                Else
                                    depth = depth - 1
                                    If depth = 0 Then
                                        FindTableEnd = nextClose + Len("</table>")
                                        Exit Function
                                    End If
                                    pos = nextClose + Len("</table>")
                                End If
                            Loop
                            FindTableEnd = 0
                        End Function

                        Function ExtractTables(html)
                            Dim result, pos, endPos, block, s
                            s = html & ""
                            s = Replace(s, "&lt;table", "<table", 1, -1, vbTextCompare)
                            s = Replace(s, "&lt;/table", "</table", 1, -1, vbTextCompare)
                            result = ""
                            Do While InStr(LCase(s), "<table") > 0
                                pos = InStr(LCase(s), "<table")
                                endPos = FindTableEnd(s, pos)
                                If endPos = 0 Then Exit Do
                                block = Mid(s, pos, endPos - pos)
                                result = result & block
                                s = Left(s, pos - 1) & Mid(s, endPos)
                            Loop
                            ExtractTables = result
                        End Function

                        Function RemoveTables(html)
                            Dim result, pos, endPos
                            result = html & ""
                            result = Replace(result, "&lt;table", "<table", 1, -1, vbTextCompare)
                            result = Replace(result, "&lt;/table", "</table", 1, -1, vbTextCompare)
                            Do While InStr(LCase(result), "<table") > 0
                                pos = InStr(LCase(result), "<table")
                                endPos = FindTableEnd(result, pos)
                                If endPos = 0 Then Exit Do
                                result = Left(result, pos - 1) & Mid(result, endPos)
                            Loop
                            RemoveTables = result
                        End Function

                        If Not rsDuyuru.EOF Then
                            Do While Not rsDuyuru.EOF
                                baslik = rsDuyuru("strd_baslik") & ""
                                icerik = rsDuyuru("strd_duyuru") & ""
                                If Not IsNull(rsDuyuru("strd_tarih")) Then tarihVal = rsDuyuru("strd_tarih") Else tarihVal = ""

                                Dim sabitAktif
                                sabitAktif = rsDuyuru("sabit_aktif")
                                If IsNull(sabitAktif) Then sabitAktif = 0

                                If Trim(baslik) <> "" Then
                                    Response.Write "<tr><td class='duyuru-baslik'>" & baslik & "</td></tr>"
                                End If

                                Dim linkIcerikVal, listeStiliVal, tumIcerik, linkSatirlar, linkHtml, satir, parcalar, metin, url, tekUrl
                                linkIcerikVal = rsDuyuru("strd_link_icerik") & ""
                                listeStiliVal = rsDuyuru("strd_liste_stili") & ""
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
                                                If Left(LCase(satir), 4) = "http" Or Left(LCase(satir), 4) = "pdf/" Or InStr(satir, ".  ") > 0 Then
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
                                If Not IsNull(rsDuyuru("strd_tarih")) Then
                                    yayinTarihi = rsDuyuru("strd_tarih")
                                    gunFarki = DateDiff("d", yayinTarihi, Now())
                                End If
                                If Not IsNull(rsDuyuru("strd_ozel_gun")) Then ozelGunMu = rsDuyuru("strd_ozel_gun")
                                If Not IsNull(rsDuyuru("strd_gorsel_kapali")) Then gorsellerKapaliMi = rsDuyuru("strd_gorsel_kapali")

                                If gorsellerKapaliMi Then
                                    gorselleriGoster = False
                                ElseIf ozelGunMu Then
                                    gorselleriGoster = True
                                ElseIf gunFarki < 3 Then
                                    gorselleriGoster = True
                                Else
                                    gorselleriGoster = False
                                End If

                                If Trim(icerik) <> "" Then
                                    Dim tablesAlways, bodyContent
                                    tablesAlways = ExtractTables(icerik)
                                    If Len(tablesAlways) > 0 Then
                                        bodyContent = RemoveTables(icerik)
                                    Else
                                        bodyContent = icerik
                                    End If

                                    Response.Write "<tr><td class='duyuru-icerik'>"

                                    If InStr(LCase(bodyContent), "<img") > 0 Then
                                        imageCounter = imageCounter + 1
                                        Dim linkId, contentId, textContent, imageContent, imgPos
                                        linkId = "imgLink" & imageCounter
                                        contentId = "imgContent" & imageCounter
                                        imgPos = InStr(LCase(bodyContent), "<img")
                                        If imgPos > 1 Then
                                            textContent = Left(bodyContent, imgPos - 1)
                                            imageContent = Mid(bodyContent, imgPos)
                                        Else
                                            textContent = ""
                                            imageContent = bodyContent
                                        End If

                                        If Len(tablesAlways) = 0 And InStr(LCase(imageContent), "<table") > 0 Then
                                            tablesAlways = ExtractTables(imageContent)
                                            If Len(tablesAlways) > 0 Then
                                                imageContent = RemoveTables(imageContent)
                                            Else
                                                gorselleriGoster = True
                                            End If
                                        End If

                                        If Trim(textContent) <> "" Then Response.Write textContent
                                        If Len(tablesAlways) > 0 Then
                                            Response.Write "<div class='duyuru-tablo-always'>" & tablesAlways & "</div>"
                                        End If
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
                                    Else
                                        If Trim(bodyContent) <> "" Then
                                            If gorsellerKapaliMi And Len(tablesAlways) = 0 And InStr(LCase(bodyContent), "<table") = 0 Then
                                                imageCounter = imageCounter + 1
                                                Dim linkId2, contentId2
                                                linkId2 = "imgLink" & imageCounter
                                                contentId2 = "imgContent" & imageCounter
                                                Response.Write "<div class='image-toggle-link' id='" & linkId2 & "' onclick=""toggleImage('" & linkId2 & "', '" & contentId2 & "')"">"
                                                Response.Write "  <span class='text'>İçeriği Açmak İçin Tıklayınız</span>"
                                                Response.Write "</div>"
                                                Response.Write "<div class='image-content' id='" & contentId2 & "'>" & bodyContent & "</div>"
                                            Else
                                                Response.Write bodyContent
                                            End If
                                        End If
                                        If Len(tablesAlways) > 0 Then
                                            Response.Write "<div class='duyuru-tablo-always'>" & tablesAlways & "</div>"
                                        End If
                                    End If

                                    Response.Write "</td></tr>"
                                End If

                                If CLng(sabitAktif) = 0 Then
                                    If Len(tarihVal) > 0 Then
                                        Response.Write "<tr><td class='duyuru-tarih'>" & _
                                            Right("0" & Day(tarihVal),2) & "/" & _
                                            Right("0" & Month(tarihVal),2) & "/" & _
                                            Year(tarihVal) & "</td></tr>"
                                    End If
                                End If
                                Response.Write "<tr><td style='height:5px;'></td></tr>"
                                rsDuyuru.MoveNext
                            Loop
                        Else
                            Response.Write "<tr><td>Hiç duyuru kaydı bulunamadı.</td></tr>"
                        End If
                        rsDuyuru.Close
                        Set rsDuyuru = Nothing
                        %>
                    </table>
                </div>
            </td>
        </tr>
    </table>
</div>
</body>
</html>
