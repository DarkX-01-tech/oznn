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
