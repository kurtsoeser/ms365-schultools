/*
 * Seiten-Skript für index.html (klassisches Skript, kein Modul).
 * Theme und Marke setzt src/shared/theme-boot.js davor.
 * Hier nur die Dashboard-Bereichs-Attribute, damit kein IT-Flash vor der Persona entsteht.
 */
(function () {
    try {
        var d = document.documentElement;
        var sec = localStorage.getItem('ms365-dash-main-section-v1');
        if (sec !== 'tasks' && sec !== 'catalog' && sec !== 'playbooks') {
            sec = localStorage.getItem('ms365-dash-it-focus-v1') === '0' ? 'catalog' : 'tasks';
        }
        d.setAttribute('data-dash-main-section', sec);
        /* Portal-first: kein IT-Flash vor Persona-Resolve */
        d.setAttribute('data-dash-app-only', '');
    } catch (e) {}
})();
