import {
    loadReleaseNotes,
    appendReleaseNote,
    saveReleaseNotes,
    setLastSeenAt
} from './release-notes-store.js';
import { loadAccessOverride, saveAccessOverride } from './access-override-store.js';

// Note: Die Imports oben wirken unhandlich; sie werden hier absichtlich getrennt gehalten,
// damit Vite/ESM sicher die Modulfunktionalität beibehält.

function $(id) {
    return document.getElementById(id);
}

function parsePinsFromTextarea(text) {
    return String(text || '')
        .split(/\r\n|\n|\r/)
        .map((x) => String(x || '').trim())
        .filter(Boolean);
}

function renderPins(pins) {
    const el = $('adminUserPinsText');
    if (!el) return;
    el.value = Array.isArray(pins) ? pins.join('\n') : '';
}

function renderNotesList(notes) {
    const wrap = $('adminReleaseNotesList');
    if (!wrap) return;
    wrap.replaceChildren();

    if (!notes.length) {
        const p = document.createElement('p');
        p.style.color = '#6c757d';
        p.textContent = 'Noch keine Release-Notes vorhanden.';
        wrap.appendChild(p);
        return;
    }

    notes.forEach((n) => {
        const div = document.createElement('div');
        div.style.border = '1px solid var(--border)';
        div.style.borderRadius = '12px';
        div.style.padding = '12px 12px';
        div.style.background = '#fff';

        const h = document.createElement('h4');
        h.style.margin = '0 0 6px';
        h.style.fontSize = '0.98em';
        h.style.color = '#32325d';
        h.textContent = n.title || '(ohne Titel)';

        const meta = document.createElement('div');
        meta.style.color = 'var(--muted)';
        meta.style.fontSize = '0.85em';
        meta.style.marginBottom = '8px';
        const d = n.at ? new Date(n.at) : null;
        meta.textContent = d && !Number.isNaN(d.getTime()) ? `Stand: ${d.toLocaleString('de-AT')}` : 'Stand: -';

        const pre = document.createElement('pre');
        pre.style.margin = '0';
        pre.style.whiteSpace = 'pre-wrap';
        pre.style.fontFamily = "Consolas, 'Segoe UI', monospace";
        pre.style.fontSize = '0.9em';
        pre.style.lineHeight = '1.4';
        pre.textContent = n.body || '';

        div.appendChild(h);
        div.appendChild(meta);
        div.appendChild(pre);
        wrap.appendChild(div);
    });
}

function init() {
    const config = typeof window !== 'undefined' ? window.MS365_ACCESS_CONFIG : null;
    const override = loadAccessOverride(localStorage) || null;

    const enabledBox = $('adminAccessEnabled');
    if (enabledBox) {
        enabledBox.checked = !!(override && typeof override.enabled === 'boolean' ? override.enabled : config && config.enabled !== false);
    }

    const initialPins = override && Array.isArray(override.pins) ? override.pins : Array.isArray(config && config.pins) ? config.pins : [];
    renderPins(initialPins);

    const btnSavePins = $('adminSavePinsBtn');
    if (btnSavePins) {
        btnSavePins.addEventListener('click', function () {
            const pins = parsePinsFromTextarea($('adminUserPinsText') && $('adminUserPinsText').value);
            const enabled = $('adminAccessEnabled') ? $('adminAccessEnabled').checked : true;
            saveAccessOverride({ enabled: enabled, pins: pins }, localStorage);
            renderPins(pins);
            alert('User-PINs gespeichert (lokal in diesem Browserprofil).');
        });
    }

    function downloadJson(payload, filename) {
        try {
            const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename || 'ms365-schooltool-admin-export.json';
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(function () {
                URL.revokeObjectURL(url);
            }, 250);
        } catch {
            alert('Export fehlgeschlagen.');
        }
    }

    function refreshPins() {
        if (!$('adminUserPinsText')) return;
        const ov = loadAccessOverride(localStorage);
        const effectivePins = ov && Array.isArray(ov.pins) && ov.pins.length ? ov.pins : Array.isArray(config?.pins) ? config.pins : [];
        renderPins(effectivePins);
        if ($('adminAccessEnabled') && ov && typeof ov.enabled === 'boolean') $('adminAccessEnabled').checked = ov.enabled;
    }

    const btnAddNote = $('adminAddReleaseNoteBtn');
    if (btnAddNote) {
        btnAddNote.addEventListener('click', function () {
            const title = $('adminReleaseTitleInput') ? $('adminReleaseTitleInput').value : '';
            const body = $('adminReleaseBodyInput') ? $('adminReleaseBodyInput').value : '';
            const t = String(title || '').trim();
            const b = String(body || '').trim();
            if (!t || !b) {
                alert('Bitte Titel und Text für die Release-Note eingeben.');
                return;
            }
            appendReleaseNote({ title: t, body: b, at: new Date().toISOString() }, localStorage);
            if ($('adminReleaseTitleInput')) $('adminReleaseTitleInput').value = '';
            if ($('adminReleaseBodyInput')) $('adminReleaseBodyInput').value = '';
            refresh();
        });
    }

    const btnClearNotes = $('adminClearReleaseNotesBtn');
    if (btnClearNotes) {
        btnClearNotes.addEventListener('click', function () {
            if (!confirm('Release-Notes wirklich löschen?')) return;
            saveReleaseNotes([], localStorage);
            // last-seen zurücksetzen, damit User beim nächsten Öffnen wieder etwas sehen.
            setLastSeenAt('', localStorage);
            refresh();
        });
    }

    const btnLogout = $('adminLogoutBtn');
    if (btnLogout) {
        btnLogout.addEventListener('click', function () {
            try {
                sessionStorage.removeItem('ms365-admin-access-granted-v1');
                sessionStorage.removeItem('ms365-access-granted-v1');
            } catch {
                // ignore
            }
            location.replace('welcome.html');
        });
    }

    function refresh() {
        const notes = loadReleaseNotes(localStorage);
        renderNotesList(notes);
        refreshPins();
    }

    const btnExportAccess = $('adminExportAccessBtn');
    if (btnExportAccess) {
        btnExportAccess.addEventListener('click', function () {
            const enabled = $('adminAccessEnabled') ? $('adminAccessEnabled').checked : true;
            const pins = parsePinsFromTextarea($('adminUserPinsText') && $('adminUserPinsText').value);
            const releaseNotes = loadReleaseNotes(localStorage);
            downloadJson(
                {
                    exportedAt: new Date().toISOString(),
                    accessOverride: { enabled: enabled, pins: pins },
                    releaseNotes: releaseNotes
                },
                'ms365-schooltool-access-and-release-notes.json'
            );
        });
    }

    const fileInput = $('adminImportAccessFile');
    const btnImport = $('adminImportAccessBtn');
    if (btnImport && fileInput) {
        btnImport.addEventListener('click', function () {
            const file = fileInput.files && fileInput.files[0];
            if (!file) {
                alert('Bitte zuerst eine JSON-Datei auswählen.');
                return;
            }
            const reader = new FileReader();
            reader.onload = function () {
                try {
                    const text = String(reader.result || '');
                    const data = JSON.parse(text);
                    if (data && data.accessOverride) {
                        saveAccessOverride(data.accessOverride, localStorage);
                    }
                    if (data && data.releaseNotes) {
                        saveReleaseNotes(data.releaseNotes, localStorage);
                        // Damit die neu hinzugefügten Hinweise für die neue Schule sichtbar sind:
                        setLastSeenAt('', localStorage);
                    }
                    alert('Import abgeschlossen.');
                    refresh();
                } catch {
                    alert('Import fehlgeschlagen: Ungültiges JSON-Format.');
                }
            };
            reader.readAsText(file);
        });
    }

    refresh();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();

