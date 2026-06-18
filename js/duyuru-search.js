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

function searchAnnouncements() {
    var raw = document.getElementById('search-bar').value.trim();
    var tokens = normalize(raw).split(/\s+/).filter(Boolean);
    var rows = document.querySelectorAll('.announcements-table tr');
    var noRow = document.getElementById('noMatchesRow');

    if (tokens.length === 0) {
        rows.forEach(function(r) { r.style.display = 'table-row'; });
        if (noRow) noRow.style.display = 'none';
        return;
    }

    rows.forEach(function(r) { r.style.display = 'none'; });
    var matchFound = false;

    document.querySelectorAll('.duyuru-baslik').forEach(function(td) {
        var trBaslik = td.parentElement;
        var trIcerik = trBaslik.nextElementSibling;
        var trTarih = trIcerik ? trIcerik.nextElementSibling : null;
        var combo = normalize(td.innerText + ' ' + (trIcerik ? trIcerik.innerText : ''));
        var hit = tokens.every(function(tok) { return combo.indexOf(tok) !== -1; });

        if (hit) {
            [trBaslik, trIcerik, trTarih].forEach(function(r) { if (r) r.style.display = 'table-row'; });
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
