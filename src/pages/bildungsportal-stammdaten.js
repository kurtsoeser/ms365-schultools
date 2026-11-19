/*
 * Seiten-Skript für tools/bildungsportal-stammdaten.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
(function () {
    'use strict';
    function notify(msg) {
        if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(msg, { kind: 'info', title: 'Bildungsportal' });
        } else if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(msg, { kind: 'info' });
        }
    }
    ['bpBtnConnect', 'bpBtnSync'].forEach(function (id) {
        var btn = document.getElementById(id);
        if (!btn) return;
        btn.addEventListener('click', function () {
            notify(
                'Die direkte Bildungsportal-Anbindung ist noch in Entwicklung. Bitte nutzen Sie vorerst den Datei-Import im Playbook.'
            );
        });
    });
})();
