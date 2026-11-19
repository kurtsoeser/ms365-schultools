/*
 * Seiten-Skript für tools/archiv/teams-archiv.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
(function () {
    var t = '../teams-archiv.html' + (window.location.search || '') + (window.location.hash || '');
    window.location.replace(t);
})();
