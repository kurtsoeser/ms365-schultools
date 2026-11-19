/*
 * Seiten-Skript für tools/gast-zugaenge.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
(function () {
    var hash = String(window.location.hash || '');
    var target = 'gaeste-verwalten.html' + (window.location.search || '') + (hash || '#tab=teams');
    window.location.replace(target);
})();
