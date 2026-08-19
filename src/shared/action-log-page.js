(function () {
    'use strict';

    function setStatus(text, kind) {
        const el = document.getElementById('actionLogStatus');
        if (!el) return;
        el.style.display = text ? 'block' : 'none';
        el.textContent = text || '';
        el.dataset.kind = kind || 'info';
    }

    function renderActionLog() {
        const tbody = document.getElementById('actionLogTableBody');
        if (!tbody) return;
        tbody.replaceChildren();
        const rows = window.ms365ActionLog && typeof window.ms365ActionLog.list === 'function' ? window.ms365ActionLog.list(200) : [];
        if (!rows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 5;
            td.style.color = '#6c757d';
            td.textContent = 'Noch keine Einträge.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            return;
        }
        rows.forEach(function (row) {
            const tr = document.createElement('tr');
            const when = row.at ? String(row.at).replace('T', ' ').slice(0, 19) : '';
            const result = row.result === 'error' ? 'Fehler' : row.result === 'skip' ? 'Übersprungen' : 'OK';
            if (row.result === 'ok') tr.style.background = 'color-mix(in srgb, #0d8050 6%, transparent)';
            if (row.result === 'error') tr.style.background = 'color-mix(in srgb, #c0392b 7%, transparent)';

            [when, row.tool || 'app', row.action || '', row.target || '', row.summary || result].forEach(function (value) {
                const td = document.createElement('td');
                td.textContent = String(value || '');
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    }

    function exportJson() {
        if (!window.ms365ActionLog || typeof window.ms365ActionLog.exportJson !== 'function') return;
        const blob = new Blob([window.ms365ActionLog.exportJson()], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'ms365-aktionsprotokoll.json';
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(function () {
            URL.revokeObjectURL(url);
        }, 250);
    }

    function bind() {
        const btnRefresh = document.getElementById('actionLogRefresh');
        const btnExport = document.getElementById('actionLogExport');
        const btnClear = document.getElementById('actionLogClear');

        if (btnRefresh) {
            btnRefresh.addEventListener('click', function () {
                renderActionLog();
                setStatus('Aktionsprotokoll aktualisiert.', 'ok');
            });
        }
        if (btnExport) {
            btnExport.addEventListener('click', function () {
                exportJson();
                setStatus('Aktionsprotokoll exportiert.', 'ok');
            });
        }
        if (btnClear) {
            btnClear.addEventListener('click', function () {
                if (window.ms365ActionLog && typeof window.ms365ActionLog.clear === 'function') {
                    window.ms365ActionLog.clear();
                }
                renderActionLog();
                setStatus('Aktionsprotokoll geleert.', 'ok');
            });
        }

        renderActionLog();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bind);
    } else {
        bind();
    }
})();
