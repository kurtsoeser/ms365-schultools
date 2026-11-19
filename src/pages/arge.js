/*
 * Seiten-Skript für tools/arge.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
(function () {
    var t = 'arge-fachgruppen.html' + (window.location.search || '') + (window.location.hash || '');
    window.location.replace(t);
})();
