/*
 * Seiten-Skript für tools/sharepoint-liste-stammdaten.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
document.addEventListener('DOMContentLoaded', function () {
    function refreshCounts() {
        var el = document.getElementById('spsCounts');
        if (!el || typeof window.ms365TenantSettingsLoad !== 'function') return;
        try {
            var s = window.ms365TenantSettingsLoad();
            var st = (s && Array.isArray(s.students) ? s.students : []).length;
            var fa = (s && Array.isArray(s.subjects) ? s.subjects : []).filter(function (x) {
                return x && String(x.code || '').trim();
            }).length;
            var ar = (s && Array.isArray(s.arges) ? s.arges : []).filter(function (x) {
                return x && String(x.code || '').trim();
            }).length;
            var kl = (s && Array.isArray(s.classes) ? s.classes : []).filter(function (x) {
                return x && (String(x.code || '').trim() || String(x.name || '').trim());
            }).length;
            el.innerHTML =
                '<i class="bi bi-bar-chart"></i>' +
                st + ' Schülerinnen · ' + fa + ' Fächer/FG · ' + ar + ' ARGEs · ' + kl + ' Klassen';
        } catch (e) {
            el.innerHTML = '<i class="bi bi-bar-chart"></i>Stammdaten prüfen';
        }
    }
    setTimeout(refreshCounts, 120);
    setTimeout(refreshCounts, 600);
});
