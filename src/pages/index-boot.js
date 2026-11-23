/*
 * Seiten-Skript für index.html (klassisches Skript, kein Modul).
 * Theme und Marke setzt src/shared/theme-boot.js davor.
 * Hier nur die Dashboard-Bereichs-Attribute, damit kein IT-Flash vor der Persona entsteht.
 */
(function () {
    try {
        var d = document.documentElement;
        /* Eine IT-Ansicht: Gruppenleiste. Katalog nur noch für Suche. */
        d.setAttribute('data-dash-main-section', 'tasks');
        try {
            localStorage.setItem('ms365-dash-main-section-v1', 'tasks');
        } catch (e2) {}
        /* Portal-first: kein IT-Flash vor Persona-Resolve */
        d.setAttribute('data-dash-app-only', '');
    } catch (e) {}
})();
