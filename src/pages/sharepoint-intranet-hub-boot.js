/*
 * Frühes Seiten-Skript für tools/sharepoint-intranet-hub.html.
 * Setzt data-ih-audience vor dem ersten Paint und fällt nach Timeout auf IT zurück,
 * falls die Modul-Prüfung nicht rechtzeitig antwortet (CSP ohne unsafe-inline).
 */
(function () {
    document.documentElement.setAttribute('data-ih-audience', 'pending');
    setTimeout(function () {
        if (document.documentElement.getAttribute('data-ih-audience') === 'pending') {
            document.documentElement.setAttribute('data-ih-audience', 'it');
        }
    }, 12000);
})();
