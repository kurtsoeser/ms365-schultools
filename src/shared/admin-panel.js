import {
    loadReleaseNotes,
    appendReleaseNote,
    saveReleaseNotes,
    updateReleaseNote,
    deleteReleaseNote,
    setLastSeenAt,
    loadMergedReleaseNotes,
    sanitizeReleaseHtml,
    toPublishedJson
} from './release-notes-store.js';
import { loadAccessOverride, saveAccessOverride, getActiveUserAccessConfig } from './access-override-store.js';

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

/** @type {Array<{ src: string, alt: string }>} */
let draftImages = [];
/** @type {Array<object>} */
let notesCache = [];
/** @type {() => void | Promise<void>} */
let refreshNotes = function () {};

function selectedKind() {
    const el = document.querySelector('input[name="adminRnKind"]:checked');
    return el ? String(el.value) : 'feature';
}

function setSelectedKind(kind) {
    const want = String(kind || 'feature');
    document.querySelectorAll('input[name="adminRnKind"]').forEach((input) => {
        input.checked = input.value === want;
    });
}

function editorEl() {
    return $('adminReleaseBodyEditor');
}

function getEditorHtml() {
    const el = editorEl();
    return sanitizeReleaseHtml(el ? el.innerHTML : '');
}

function setEditorHtml(html) {
    const el = editorEl();
    if (el) el.innerHTML = sanitizeReleaseHtml(html || '');
}

function renderImagePreviews() {
    const wrap = $('adminRnImagePreview');
    if (!wrap) return;
    wrap.replaceChildren();
    draftImages.forEach((img, idx) => {
        const fig = document.createElement('figure');
        fig.className = 'admin-rn-thumb';
        const image = document.createElement('img');
        image.src = img.src;
        image.alt = img.alt || 'Screenshot';
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'admin-rn-thumb__remove';
        btn.setAttribute('aria-label', 'Screenshot entfernen');
        btn.innerHTML = '&times;';
        btn.addEventListener('click', () => {
            draftImages.splice(idx, 1);
            renderImagePreviews();
        });
        fig.appendChild(image);
        fig.appendChild(btn);
        wrap.appendChild(fig);
    });
}

function resetEditor() {
    if ($('adminReleaseTitleInput')) $('adminReleaseTitleInput').value = '';
    setEditorHtml('');
    draftImages = [];
    renderImagePreviews();
    setSelectedKind('feature');
    if ($('adminRnEditingId')) $('adminRnEditingId').value = '';
    const heading = $('adminRnEditorHeading');
    if (heading) heading.textContent = 'Neuer Eintrag';
    const saveLabel = $('adminRnSaveLabel');
    if (saveLabel) saveLabel.textContent = 'Speichern';
    const cancel = $('adminRnCancelEditBtn');
    if (cancel) cancel.hidden = true;
}

function fillEditor(note) {
    if ($('adminReleaseTitleInput')) $('adminReleaseTitleInput').value = note.title || '';
    setEditorHtml(note.bodyHtml || '');
    draftImages = Array.isArray(note.images) ? note.images.map((x) => ({ src: x.src, alt: x.alt || '' })) : [];
    renderImagePreviews();
    setSelectedKind(note.kind || 'feature');
    if ($('adminRnEditingId')) $('adminRnEditingId').value = note.id || '';
    const heading = $('adminRnEditorHeading');
    if (heading) heading.textContent = 'Eintrag bearbeiten';
    const saveLabel = $('adminRnSaveLabel');
    if (saveLabel) saveLabel.textContent = 'Aktualisieren';
    const cancel = $('adminRnCancelEditBtn');
    if (cancel) cancel.hidden = false;
}

function renderNotesList(notes) {
    const wrap = $('adminReleaseNotesList');
    if (!wrap) return;
    wrap.replaceChildren();

    if (!notes.length) {
        const p = document.createElement('p');
        p.className = 'admin-app__empty';
        p.textContent = 'Noch keine Release-Notes vorhanden.';
        wrap.appendChild(p);
        return;
    }

    const ui = window.ms365ReleaseNotesUi;
    notes.forEach((n) => {
        if (ui && typeof ui.renderNoteCard === 'function') {
            wrap.appendChild(
                ui.renderNoteCard(n, {
                    editable: true,
                    onEdit: fillEditor,
                    onDelete: (note) => {
                        if (!confirm('Diesen Eintrag lokal löschen?')) return;
                        deleteReleaseNote(note.id, localStorage);
                        refreshNotes();
                    }
                })
            );
            return;
        }
        const div = document.createElement('article');
        div.className = 'admin-app__note';
        div.textContent = n.title || '(ohne Titel)';
        wrap.appendChild(div);
    });
}

async function addImagesFromFiles(fileList) {
    const ui = window.ms365ReleaseNotesUi;
    const files = Array.from(fileList || []).filter((f) => f && /^image\//i.test(f.type));
    if (!files.length) return;
    if (!ui || typeof ui.compressImageFile !== 'function') {
        alert('Bild-Hilfe nicht geladen.');
        return;
    }
    for (const file of files) {
        if (draftImages.length >= 6) break;
        try {
            const compressed = await ui.compressImageFile(file, file.name);
            if (compressed.bytes > 900000) {
                alert('Screenshot nach Kompression noch zu groß – bitte kleineres Bild wählen.');
                continue;
            }
            draftImages.push({ src: compressed.src, alt: compressed.alt });
        } catch (e) {
            alert('Bild fehlgeschlagen: ' + ((e && e.message) || e));
        }
    }
    renderImagePreviews();
}

function wireReleaseEditor() {
    document.querySelectorAll('.admin-rn-toolbar [data-rn-cmd]').forEach((btn) => {
        btn.addEventListener('click', (ev) => {
            ev.preventDefault();
            const cmd = btn.getAttribute('data-rn-cmd');
            const ed = editorEl();
            if (!ed) return;
            ed.focus();
            if (cmd === 'createLink') {
                const url = window.prompt('Link-URL (https://…)', 'https://');
                if (!url) return;
                document.execCommand('createLink', false, url);
            } else {
                document.execCommand(cmd, false, null);
            }
        });
    });

    const drop = $('adminRnDropzone');
    const pick = $('adminRnPickImageBtn');
    const fileInput = $('adminRnImageInput');
    if (pick && fileInput) {
        pick.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', () => {
            addImagesFromFiles(fileInput.files);
            fileInput.value = '';
        });
    }
    if (drop) {
        ['dragenter', 'dragover'].forEach((evt) => {
            drop.addEventListener(evt, (e) => {
                e.preventDefault();
                drop.classList.add('is-dragover');
            });
        });
        ['dragleave', 'drop'].forEach((evt) => {
            drop.addEventListener(evt, (e) => {
                e.preventDefault();
                drop.classList.remove('is-dragover');
            });
        });
        drop.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            if (dt && dt.files) addImagesFromFiles(dt.files);
        });
        drop.addEventListener('paste', (e) => {
            const items = e.clipboardData && e.clipboardData.items;
            if (!items) return;
            const files = [];
            for (const item of items) {
                if (item.type && item.type.indexOf('image') === 0) {
                    const f = item.getAsFile();
                    if (f) files.push(f);
                }
            }
            if (files.length) {
                e.preventDefault();
                addImagesFromFiles(files);
            }
        });
    }

    const ed = editorEl();
    if (ed) {
        ed.addEventListener('paste', (e) => {
            const items = e.clipboardData && e.clipboardData.items;
            if (!items) return;
            for (const item of items) {
                if (item.type && item.type.indexOf('image') === 0) {
                    e.preventDefault();
                    const f = item.getAsFile();
                    if (f) addImagesFromFiles([f]);
                    return;
                }
            }
        });
    }

    const cancel = $('adminRnCancelEditBtn');
    if (cancel) cancel.addEventListener('click', resetEditor);

    const exportBtn = $('adminExportReleaseNotesBtn');
    if (exportBtn) {
        exportBtn.addEventListener('click', () => {
            const json = toPublishedJson(notesCache.length ? notesCache : loadReleaseNotes(localStorage));
            const blob = new Blob([json], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'release-notes.json';
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 250);
            alert('Datei speichern als public/release-notes.json und committen – dann sehen alle Schulen die Notes nach dem Deploy.');
        });
    }
}

const ADMIN_TAB_META = {
    licenses: {
        eyebrow: 'Betreiber',
        title: 'Schulen & Lizenzen',
        subtitle: 'Freischaltungen direkt in der Tabelle verwalten.'
    },
    access: {
        eyebrow: 'Zugang',
        title: 'User-Zugänge',
        subtitle: 'PIN optional – Freischaltung über Tenant-Lizenz; Admin nur per Betreiber-Konto.'
    },
    setup: {
        eyebrow: 'Technik',
        title: 'Setup & Diagnose',
        subtitle: 'Was noch fehlt: SharePoint-Spalten und API – Freischaltungen laufen über Lizenzen.'
    },
    templates: {
        eyebrow: 'Katalog',
        title: 'Kursteam-Vorlagen',
        subtitle: 'Zentrale Kanal-Vorlagen für alle Schulen. Schulen lesen sie, du pflegst sie hier.'
    },
    notes: {
        eyebrow: 'Kommunikation',
        title: 'Neuigkeiten',
        subtitle: 'Release-Notes mit Editor und Screenshots – auch automatisch aus GitHub-Commits.'
    }
};

const ADMIN_TAB_STORAGE_KEY = 'ms365-admin-active-tab-v1';

function setAdminTab(tabId) {
    const id = ADMIN_TAB_META[tabId] ? tabId : 'licenses';
    const meta = ADMIN_TAB_META[id];

    document.querySelectorAll('.admin-app__nav-btn[data-admin-tab]').forEach((btn) => {
        const selected = btn.getAttribute('data-admin-tab') === id;
        btn.setAttribute('aria-selected', selected ? 'true' : 'false');
    });

    document.querySelectorAll('.admin-app__panel[data-admin-panel]').forEach((panel) => {
        const active = panel.getAttribute('data-admin-panel') === id;
        panel.classList.toggle('active', active);
        if (active) panel.removeAttribute('hidden');
        else panel.setAttribute('hidden', '');
    });

    const eyebrow = $('adminAppEyebrow');
    const title = $('adminAppTitle');
    const subtitle = $('adminAppSubtitle');
    if (eyebrow) eyebrow.textContent = meta.eyebrow;
    if (title) title.textContent = meta.title;
    if (subtitle) subtitle.textContent = meta.subtitle;

    try {
        localStorage.setItem(ADMIN_TAB_STORAGE_KEY, id);
    } catch {
        // ignore
    }

    try {
        if (location.hash.replace(/^#/, '') !== id) {
            history.replaceState(null, '', '#' + id);
        }
    } catch {
        // ignore
    }
}

function tabFromLocation() {
    try {
        const hash = String(location.hash || '')
            .replace(/^#/, '')
            .trim()
            .toLowerCase();
        if (hash && ADMIN_TAB_META[hash]) return hash;
        const q = new URLSearchParams(location.search).get('tab');
        if (q && ADMIN_TAB_META[q]) return q;
    } catch {
        // ignore
    }
    return null;
}

function initAdminTabs() {
    const nav = $('adminAppTabs');
    if (!nav) return;

    nav.addEventListener('click', (ev) => {
        const btn = ev.target && ev.target.closest ? ev.target.closest('.admin-app__nav-btn[data-admin-tab]') : null;
        if (!btn || !nav.contains(btn)) return;
        setAdminTab(btn.getAttribute('data-admin-tab'));
    });

    nav.addEventListener('keydown', (ev) => {
        const buttons = Array.from(nav.querySelectorAll('.admin-app__nav-btn[data-admin-tab]'));
        const current = document.activeElement;
        const idx = buttons.indexOf(current);
        if (idx < 0) return;
        let next = -1;
        if (ev.key === 'ArrowDown' || ev.key === 'ArrowRight') next = (idx + 1) % buttons.length;
        if (ev.key === 'ArrowUp' || ev.key === 'ArrowLeft') next = (idx - 1 + buttons.length) % buttons.length;
        if (ev.key === 'Home') next = 0;
        if (ev.key === 'End') next = buttons.length - 1;
        if (next < 0) return;
        ev.preventDefault();
        buttons[next].focus();
        setAdminTab(buttons[next].getAttribute('data-admin-tab'));
    });

    window.addEventListener('hashchange', () => {
        const fromHash = tabFromLocation();
        if (fromHash) setAdminTab(fromHash);
    });

    let initial = tabFromLocation() || 'licenses';
    if (!tabFromLocation()) {
        try {
            const stored = localStorage.getItem(ADMIN_TAB_STORAGE_KEY);
            if (stored && ADMIN_TAB_META[stored]) initial = stored;
        } catch {
            // ignore
        }
    }
    setAdminTab(initial);
}

window.ms365AdminSetTab = setAdminTab;

function init() {
    initAdminTabs();
    const config = typeof window !== 'undefined' ? window.MS365_ACCESS_CONFIG : null;
    const override = loadAccessOverride(localStorage) || null;

    const enabledBox = $('adminAccessEnabled');
    if (enabledBox) {
        const active = getActiveUserAccessConfig(config);
        enabledBox.checked = !!active.enabled;
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

    wireReleaseEditor();

    const btnAddNote = $('adminAddReleaseNoteBtn');
    if (btnAddNote) {
        btnAddNote.addEventListener('click', function () {
            const title = $('adminReleaseTitleInput') ? $('adminReleaseTitleInput').value : '';
            const t = String(title || '').trim();
            const bodyHtml = getEditorHtml();
            const plain = String(editorEl() ? editorEl().innerText : '').trim();
            if (!t || (!bodyHtml && !plain && !draftImages.length)) {
                alert('Bitte Titel und Text (oder Screenshot) eingeben.');
                return;
            }
            const editingId = $('adminRnEditingId') ? String($('adminRnEditingId').value || '').trim() : '';
            const payload = {
                title: t,
                bodyHtml: bodyHtml || '<p>' + plain.replace(/</g, '&lt;') + '</p>',
                kind: selectedKind(),
                source: 'local',
                images: draftImages.slice(),
                at: new Date().toISOString()
            };
            if (editingId) {
                updateReleaseNote(editingId, payload, localStorage);
            } else {
                appendReleaseNote(payload, localStorage);
            }
            resetEditor();
            refresh();
        });
    }

    const btnClearNotes = $('adminClearReleaseNotesBtn');
    if (btnClearNotes) {
        btnClearNotes.addEventListener('click', function () {
            if (!confirm('Nur lokal gespeicherte Release-Notes löschen? (Die veröffentlichte JSON bleibt.)')) return;
            saveReleaseNotes([], localStorage);
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

    async function refresh() {
        try {
            notesCache = await loadMergedReleaseNotes(localStorage);
        } catch {
            notesCache = loadReleaseNotes(localStorage);
        }
        renderNotesList(notesCache);
        refreshPins();
    }
    refreshNotes = refresh;

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

