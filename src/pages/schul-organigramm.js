/*
 * Seiten-Skript für tools/schul-organigramm.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
(function () {
    try {
        location.replace('../tenant.html#organigramm');
    } catch (e) {
        location.href = '../tenant.html#organigramm';
    }
})();
