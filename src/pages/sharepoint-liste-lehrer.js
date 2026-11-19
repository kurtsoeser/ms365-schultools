/*
 * Seiten-Skript für tools/sharepoint-liste-lehrer.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
document.addEventListener('DOMContentLoaded', function () {
    var PREVIEW_MAX = 8;

    function esc(s) {
        return String(s == null ? '' : s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function loadTeachers() {
        if (typeof window.ms365TenantSettingsLoad !== 'function') return [];
        try {
            var s = window.ms365TenantSettingsLoad();
            return (s && Array.isArray(s.teachers) ? s.teachers : []).filter(function (t) {
                return t && String(t.code || '').trim();
            });
        } catch (e) {
            return [];
        }
    }

    function syncPreviewTitle() {
        var input = document.getElementById('splListName');
        var titleEl = document.getElementById('splPreviewListTitle');
        if (!titleEl) return;
        var name = input && String(input.value || '').trim() ? String(input.value).trim() : 'Lehrerinnen';
        titleEl.textContent = name;
    }

    function refreshPreview() {
        var countEl = document.getElementById('splTeacherCount');
        var body = document.getElementById('splPreviewBody');
        var hint = document.getElementById('splPreviewHint');
        var teachers = loadTeachers();
        var n = teachers.length;

        if (countEl) {
            countEl.innerHTML =
                '<i class="bi bi-person-badge"></i>' +
                (n ? n + ' Lehrkräfte in Stammdaten' : 'Keine Lehrkräfte in Stammdaten');
        }

        syncPreviewTitle();

        if (!body) return;

        if (!n) {
            body.innerHTML =
                '<tr><td colspan="5" style="color:var(--muted);">Noch keine Lehrkräfte in den Stammdaten. ' +
                '<a href="../tenant.html#teachers">Stammdaten öffnen</a></td></tr>';
            if (hint) hint.textContent = 'Sobald Lehrkräfte in den Stammdaten stehen, erscheinen sie hier.';
            return;
        }

        var shown = teachers.slice(0, PREVIEW_MAX);
        var rows = shown.map(function (t) {
            var name = String(t.name || '').trim() || String(t.code || '').trim();
            var code = String(t.code || '').trim();
            var email = String(t.email || '').trim();
            var personIcon = email
                ? '<i class="bi bi-person-check" title="M365-Benutzer über E-Mail"></i>'
                : '<i class="bi bi-person" title="Keine E-Mail – Personenfeld leer" style="opacity:0.45;"></i>';
            return (
                '<tr>' +
                '<td>' + esc(name) + '</td>' +
                '<td><span class="pill">' + esc(code) + '</span></td>' +
                '<td>' + esc(email || '—') + '</td>' +
                '<td>' + esc(email || '—') + '</td>' +
                '<td>' + personIcon + '</td>' +
                '</tr>'
            );
        });

        if (n > PREVIEW_MAX) {
            rows.push(
                '<tr><td colspan="5" style="color:var(--muted);font-size:0.9em;">… und ' +
                    (n - PREVIEW_MAX) +
                    ' weitere aus den Stammdaten</td></tr>'
            );
        }

        body.innerHTML = rows.join('');
        if (hint) {
            hint.textContent =
                n === 1
                    ? '1 Lehrkraft aus Ihren Stammdaten – so wirkt die Liste in SharePoint.'
                    : n + ' Lehrkräfte aus Ihren Stammdaten – so wirkt die Liste in SharePoint.';
        }
    }

    var listName = document.getElementById('splListName');
    if (listName) {
        listName.addEventListener('input', syncPreviewTitle);
        listName.addEventListener('change', syncPreviewTitle);
    }

    setTimeout(refreshPreview, 120);
    setTimeout(refreshPreview, 600);
});
