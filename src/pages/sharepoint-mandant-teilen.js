/*
 * Seiten-Skript für tools/sharepoint-mandant-teilen.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
(function () {
    function syncDomainCards() {
        var sel = document.getElementById('fDomainMode');
        if (!sel) return;
        var v = sel.value;
        document.querySelectorAll('[data-domain-mode]').forEach(function (btn) {
            btn.classList.toggle('is-active', btn.getAttribute('data-domain-mode') === v);
        });
    }
    document.querySelectorAll('[data-domain-mode]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            var sel = document.getElementById('fDomainMode');
            if (!sel) return;
            sel.value = btn.getAttribute('data-domain-mode');
            sel.dispatchEvent(new Event('change', { bubbles: true }));
            syncDomainCards();
        });
    });
    var sel = document.getElementById('fDomainMode');
    if (sel) sel.addEventListener('change', syncDomainCards);
    document.addEventListener('DOMContentLoaded', syncDomainCards);
    setTimeout(syncDomainCards, 400);
})();
