<%
session("ok") = false
%>
<!-- #include file="admin/database/Connection.asp" -->
<!-- #include file="ayarlar.asp" -->
<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>MÜ Pendik E.A.H. Portal</title>
<!-- tablo-guncelleme-20260618-v2 -->
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
        margin: 25px 0;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
        padding-bottom: 12px;
        text-shadow: 1px 1px 4px rgba(0,0,0,0.3);
        transition: color 0.3s, border-color 0.3s;
    }
    .baslik:hover          { color:#343a40; border-color:#c70039; }
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
        text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
        border-bottom: 2px solid #850303;
        padding-bottom: 4px;
        display: inline-block;
        transition: color 0.3s, border-color 0.3s;
    }
    .duyuru-baslik:hover   { color:#850303; border-color:#343a40; }

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
    .duyuru-icerik strong  { font-weight: 800 !important; }

    .duyuru-icerik table {
        width: 100%;
        max-width: 100%;
        border-collapse: collapse;
        margin: 12px 0;
        text-align: left;
        font-size: 12px;
        line-height: 1.6;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-weight: 400;
        color: #555;
        border: 1px solid #ddd;
    }
    .duyuru-icerik table td,
    .duyuru-icerik table th {
        border: 1px solid #ddd;
        padding: 8px 12px;
        text-align: left !important;
        font-size: 12px;
        line-height: 1.6;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-weight: 400;
        color: #555;
        vertical-align: top;
        word-break: break-word;
    }
    .duyuru-icerik table th,
    .duyuru-icerik table tr:first-child td {
        background-color: #f0f0f0;
        font-weight: 600;
        text-align: center !important;
    }
    .duyuru-icerik table td:first-child {
        font-weight: 600;
        width: 40%;
    }
    .duyuru-icerik table td:last-child {
        font-weight: 400;
    }
    .duyuru-icerik table b,
    .duyuru-icerik table strong {
        font-weight: 600 !important;
    }
    .duyuru-icerik .image-content.show table,
    .duyuru-icerik > table {
        display: table !important;
        visibility: visible !important;
        opacity: 1 !important;
    }

    .baslik b, .baslik strong,
    .duyuru-baslik b, .duyuru-baslik strong { font-weight: inherit !important; }

    .duyuru-tarih {
        font-size: 12px;
        color: #888;
        text-align: right;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.05);
        padding: 1em 0;
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
    .search-container {
        float: right;
        margin-right: 20px;
        margin-top: -20px;
        position: relative;
    }
    .search-container form              { display:inline-block; position:relative; }
    .search-container input[type="text"]{
        padding: 10px 15px;
        border: 1px solid #ccc;
        border-radius: 30px;
        font-size: 16px;
        transition: border-color 0.3s, box-shadow 0.3s;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
    }
    .search-container input[type="text"]:focus { border-color:#45b8c3; outline:none; box-shadow:0 4px 8px rgba(0,0,0,0.2); }
    .search-container input[type="text"]:hover { border-color:#45b8c3; }

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
        transition: all 0.3s ease, opacity 0.5s ease;
        opacity: 0;
        visibility: hidden;
        z-index: 999;
    }
    #scrollTopBtn:hover {
        background-color: #1e8c99;
        box-shadow: 0 6px 12px rgba(0,0,0,0.5);
        transform: scale(1.1);
    }

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
        transition: background 0.3s, box-shadow 0.3s;
    }
    #apDiv1:hover        { background:rgba(255,255,255,1); box-shadow:0 4px 8px rgba(0,0,0,0.4); }

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
    .duyuru-container::-webkit-scrollbar        { width:6px; }
    .duyuru-container::-webkit-scrollbar-track  { background:#ffffff; }
    .duyuru-container::-webkit-scrollbar-thumb  { background:#45b8c3; border-radius:5px; }
    .duyuru-container::-webkit-scrollbar-thumb:hover { background:#25abb9; }
    .duyuru-container::-webkit-scrollbar-button { display:none; }

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
        transform: translateX(-50%) translateY(10px);
        background: rgba(69, 184, 195, 0.95);
        color: #fff;
        padding: 6px 12px;
        border-radius: 15px;
        font-size: 11px;
        font-weight: 600;
        white-space: nowrap;
        opacity: 0;
        pointer-events: none;
        transition: all 0.25s ease;
        box-shadow: 0 3px 10px rgba(0,0,0,0.2);
        z-index: 100;
        display: flex;
        align-items: center;
        gap: 5px;
    }

    .image-tooltip svg {
        width: 14px;
        height: 14px;
        fill: none;
        stroke: #fff;
        stroke-width: 2;
        flex-shrink: 0;
    }

    .image-wrapper:hover .image-tooltip {
        opacity: 1;
        transform: translateX(-50%) translateY(0);
    }

    .duyuru-icerik img {
        max-width: 100%;
        width: 100%;
        height: auto;
        display: block;
        border-radius: 8px;
        box-shadow: 0 3px 10px rgba(0,0,0,0.12);
        margin: 10px auto;
        cursor: pointer;
        transition: all 0.25s ease;
        border: 2px solid #f0f0f0;
        box-sizing: border-box;
    }
    
    .duyuru-icerik img:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(0,0,0,0.18);
        border-color: #45b8c3;
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
        margin: auto;
        display: block;
        max-width: none;
        max-height: none;
        position: absolute;
        top: 50%;
        left: 50%;
        transform-origin: center center;
        transition: transform 0.3s ease, opacity 0.3s ease;
        border-radius: 8px;
        box-shadow: 0 15px 50px rgba(0,0,0,0.5);
        opacity: 0;
        cursor: grab;
    }

    .modal-content:active {
        cursor: grabbing;
    }

    .modal-content.zoomed {
        cursor: move;
    }
    .image-modal.show .modal-content {
        transform: translate(-50%, -50%) scale(1);
        opacity: 1;
    }
    .modal-close {
        position: absolute;
        top: 20px;
        right: 30px;
        color: #fff;
        font-size: 36px;
        font-weight: bold;
        cursor: pointer;
        transition: all 0.25s;
        z-index: 10001;
        width: 45px;
        height: 45px;
        background: rgba(69, 184, 195, 0.9);
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        line-height: 1;
    }

    .modal-close:hover {
        background: rgba(37, 171, 185, 1);
        transform: rotate(90deg) scale(1.1);
    }

    .modal-prev, .modal-next {
        cursor: pointer;
        position: absolute;
        top: 50%;
        width: 50px;
        height: 50px;
        padding: 0;
        margin-top: -25px;
        color: white;
        font-weight: bold;
        font-size: 26px;
        transition: all 0.25s ease;
        border-radius: 50%;
        user-select: none;
        background: rgba(69, 184, 195, 0.9);
        display: none;
        align-items: center;
        justify-content: center;
        border: 2px solid rgba(255,255,255,0.2);
    }
    .modal-next { right: 25px; }
    .modal-prev { left: 25px; }
    .modal-prev:hover, .modal-next:hover {
        background: rgba(37, 171, 185, 1);
        transform: scale(1.1);
    }

    .modal-prev:active, .modal-next:active {
        transform: scale(0.95);
    }

    .zoom-controls {
        position: absolute;
        bottom: 30px;
        left: 50%;
        transform: translateX(-50%);
        display: none;
        gap: 10px;
        z-index: 10002;
        background: rgba(0, 0, 0, 0.5);
        padding: 8px 12px;
        border-radius: 25px;
        backdrop-filter: blur(5px);
    }

    .zoom-btn {
        width: 40px;
        height: 40px;
        background: rgba(69, 184, 195, 0.9);
        color: white;
        border: 2px solid rgba(255,255,255,0.2);
        border-radius: 50%;
        font-size: 20px;
        font-weight: bold;
        cursor: pointer;
        transition: all 0.25s ease;
        display: flex;
        align-items: center;
        justify-content: center;
        user-select: none;
        position: relative;
        overflow: visible;
    }

    .zoom-btn:hover {
        background: rgba(37, 171, 185, 1);
        transform: scale(1.1);
    }

    .zoom-btn:active {
        transform: scale(0.95);
    }

    .zoom-btn svg {
        width: 20px;
        height: 20px;
        fill: none;
        stroke: white;
        stroke-width: 3;
        stroke-linecap: round;
        transition: all 0.25s ease;
    }

    .zoom-out:hover svg {
        opacity: 0.3;
    }

    .zoom-in:hover svg {
        transform: scale(1.3);
    }

    .zoom-reset svg {
        width: 22px;
        height: 22px;
        fill: none;
        stroke: white;
        stroke-width: 2.5;
        stroke-linecap: round;
        stroke-linejoin: round;
        transition: transform 0.25s ease;
    }

    .zoom-reset:hover svg {
        transform: rotate(360deg);
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
        font-weight: 600;
        z-index: 10002;
        opacity: 0;
        transition: opacity 0.3s ease;
        pointer-events: none;
    }
    .zoom-indicator.show { opacity: 1; }

    /* --- Resim aç / kapa butonu --- */
.image-toggle-link {
  display: inline-block;
  color: #1e8b99;
  font-size: 13px;
  font-weight: 600;
  margin: 8px 0;
  padding: 6px 10px;
  cursor: pointer;
  user-select: none;
  transition: all 0.2s ease;
  border-left: 3px solid #1e8b99;
  background: linear-gradient(90deg, rgba(30,139,153,0.12) 0%, transparent 100%);
  border-radius: 0 6px 6px 0;
}

.image-toggle-link:hover {
  color: #0d5f6b;
  border-left-color: #0d5f6b;
  padding-left: 12px;
}

/* Açılır / kapanır resim alanı */
.image-content {
    max-height: 0;
    overflow: hidden;
    opacity: 0;
    margin: 0;
    transition:
        max-height 0.35s ease,
        opacity 0.25s ease,
        margin 0.25s ease;
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

function normalize(txt) {
    return txt
        .toLocaleLowerCase('tr-TR')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^\w\s]/g, '');
}

function debounce(fn, delay) {
    var t;
    return function() {
        clearTimeout(t);
        var args = arguments;
        var ctx = this;
        t = setTimeout(function() { fn.apply(ctx, args); }, delay);
    };
}

function resetSearchExpanded() {
    document.querySelectorAll('.image-content[data-search-open]').forEach(function(box) {
        box.classList.remove('show');
        box.removeAttribute('data-search-open');
    });
}

function expandTablesInRow(tr) {
    if (!tr) return;
    tr.querySelectorAll('.image-content').forEach(function(box) {
        if (box.querySelector('table')) {
            box.classList.add('show');
            box.setAttribute('data-search-open', '1');
            var toggle = box.previousElementSibling;
            if (toggle && toggle.classList.contains('image-toggle-link')) {
                toggle.classList.add('active');
            }
        }
    });
}

function getAnnouncementRows(trBaslik) {
    var rows = [trBaslik];
    var tr = trBaslik.nextElementSibling;
    while (tr && !tr.querySelector('.duyuru-baslik')) {
        rows.push(tr);
        tr = tr.nextElementSibling;
    }
    return rows;
}

function searchAnnouncements() {
    var raw = document.getElementById('search-bar').value.trim();
    var tokens = normalize(raw).split(/\s+/).filter(Boolean);
    var rows = document.querySelectorAll('.announcements-table tr');
    var noRow = document.getElementById('noMatchesRow');

    resetSearchExpanded();

    if (tokens.length === 0) {
        rows.forEach(function(r) { r.style.display = 'table-row'; });
        if (noRow) noRow.style.display = 'none';
        return;
    }

    rows.forEach(function(r) { r.style.display = 'none'; });
    var matchFound = false;

    document.querySelectorAll('.duyuru-baslik').forEach(function(td) {
        var trBaslik = td.parentElement;
        var announcementRows = getAnnouncementRows(trBaslik);
        var combo = '';
        announcementRows.forEach(function(r) {
            combo += ' ' + r.innerText;
        });
        combo = normalize(combo);
        var hit = tokens.every(function(tok) { return combo.indexOf(tok) !== -1; });

        if (hit) {
            announcementRows.forEach(function(r) {
                r.style.display = 'table-row';
                expandTablesInRow(r);
            });
            matchFound = true;
        }
    });

    if (noRow) noRow.style.display = matchFound ? 'none' : 'table-row';
}

document.addEventListener('DOMContentLoaded', function() {
    var searchBar = document.getElementById('search-bar');
    if (searchBar) {
        searchBar.addEventListener('keyup', debounce(searchAnnouncements, 300));
    }
});
</script>
</head>

<body>

<div id="imageModal" class="image-modal">
    <span class="modal-close" onclick="closeModal()">&times;</span>
    <div id="zoomControls" class="zoom-controls">
        <button class="zoom-btn zoom-out" onclick="zoomOut()" title="Uzaklaştır (-)">
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
                <line x1="5" y1="12" x2="19" y2="12"></line>
            </svg>
        </button>
        <button class="zoom-btn zoom-reset" onclick="resetZoom()" title="Sıfırla (0)">
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
                <path d="M1 4v6h6"></path>
                <path d="M23 20v-6h-6"></path>
                <path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4l-4.64 4.36A9 9 0 0 1 3.51 15"></path>
            </svg>
        </button>
        <button class="zoom-btn zoom-in" onclick="zoomIn()" title="Yakınlaştır (+)">
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
                <line x1="12" y1="5" x2="12" y2="19"></line>
                <line x1="5" y1="12" x2="19" y2="12"></line>
            </svg>
        </button>
    </div>
    
    <div id="zoomIndicator" class="zoom-indicator">100%</div>
    
    <img class="modal-content" id="modalImage">
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
                                If InStr(LCase(icerik), "<img") > 0 Then
                                    imageCounter = imageCounter + 1
                                    Dim linkId, contentId, textContent, imageContent, imgPos
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
                                        Response.Write "Görseli Açmak İçin Tıklayınız"
                                        Response.Write "</div>"
                                        Response.Write "<div class='image-content' id='" & contentId & "'>"
                                        Response.Write imageContent
                                        Response.Write "</div>"
                                    End If
                                    Response.Write "</td></tr>"
                                Else
                                    Response.Write "<tr><td class='duyuru-icerik'>"
                                    If InStr(LCase(icerik), "<table") > 0 Then
                                        Response.Write icerik
                                    ElseIf gorsellerKapaliMi Then
                                        Dim linkId2, contentId2
                                        imageCounter = imageCounter + 1
                                        linkId2 = "imgLink" & imageCounter
                                        contentId2 = "imgContent" & imageCounter
                                        Response.Write "<div class='image-toggle-link' id='" & linkId2 & "' onclick=""toggleImage('" & linkId2 & "', '" & contentId2 & "')"">"
                                        Response.Write "İçeriği Açmak İçin Tıklayınız"
                                        Response.Write "</div>"
                                        Response.Write "<div class='image-content' id='" & contentId2 & "'>" & icerik & "</div>"
                                    Else
                                        Response.Write icerik
                                    End If
                                    Response.Write "</td></tr>"
                                End If
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
