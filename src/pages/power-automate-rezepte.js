/*
 * Seiten-Skript für tools/power-automate-rezepte.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
(function () {
    var cat = window.ms365PaRecipes;
    if (!cat) return;

    // Die konkreten Keys je Rezept stehen im STORAGE_CATALOG (browser-backup.js).
    function doneKeyFor(id) {
        return 'ms365-pa-done-' + id;
    }
    function isDone(id) {
        try { return localStorage.getItem(doneKeyFor(id)) === '1'; } catch (e) { return false; }
    }
    function clearDone(id) {
        try { localStorage.removeItem(doneKeyFor(id)); } catch (e) {}
    }

    function escapeHtml(s) {
        return String(s || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function hrefFor(r) {
        return 'pa-' + r.id + '.html';
    }

    function render() {
        var host = document.getElementById('paHub');
        if (!host) return;
        var cards = [];
        var done = 0;
        var total = 0;

        // Setup-Pakete (Freistellung, Lehrer-Freistellung, Schulaktivitäten, …)
        var extras = Array.isArray(cat.hubExtras)
            ? cat.hubExtras
            : cat.hubExtra
              ? [cat.hubExtra]
              : [];
        extras.forEach(function (extra) {
            total++;
            var frDone = false;
            var doneKey = extra.doneKey || ('ms365-pa-done-' + extra.id);
            try { frDone = localStorage.getItem(doneKey) === '1'; } catch (e) {}
            if (frDone) done++;
            cards.push(
                '<article class="pa-card">' +
                '<h3><i class="bi ' + escapeHtml(extra.icon) + '"></i>' + escapeHtml(extra.title) + '</h3>' +
                '<p class="muted" style="margin:0;">' + escapeHtml(extra.summary) + '</p>' +
                '<p class="pa-status' + (frDone ? ' is-done' : '') + '">' + (frDone ? '✓ markiert' : 'offen') + '</p>' +
                '<a class="btn btn-sm" href="' + escapeHtml(extra.href) + '">Öffnen</a>' +
                '</article>'
            );
        });

        cat.list.forEach(function (r) {
            total++;
            var on = isDone(r.id);
            if (on) done++;
            cards.push(
                '<article class="pa-card">' +
                '<h3><i class="bi ' + escapeHtml(r.icon) + '"></i>' + escapeHtml(r.title) + '</h3>' +
                '<p class="muted" style="margin:0;">' + escapeHtml(r.summary) + '</p>' +
                '<p class="pa-status' + (on ? ' is-done' : '') + '">' + (on ? '✓ ' + escapeHtml(r.doneLabel || 'umgesetzt') : 'offen') + '</p>' +
                '<a class="btn btn-sm" href="' + escapeHtml(hrefFor(r)) + '">Öffnen &amp; ausführen</a>' +
                '</article>'
            );
        });

        host.innerHTML = cards.join('');
        var p = document.getElementById('paProgress');
        if (p) p.textContent = done + ' von ' + total + ' markiert';
    }

    var reset = document.getElementById('paReset');
    if (reset) {
        reset.addEventListener('click', function () {
            if (!confirm('Alle Automations-Status zurücksetzen?')) return;
            cat.list.forEach(function (r) { clearDone(r.id); });
            try { localStorage.removeItem('ms365-pa-done-freistellung'); } catch (e) {}
            try { localStorage.setItem('ms365-power-automate-recipes-v1', '{}'); } catch (e2) {}
            render();
        });
    }
    render();
})();
