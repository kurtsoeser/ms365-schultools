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

/**
 * Betreiber-Admin: Zugang nur License-API (Operator-UPN/OID).
 * Schulen: Tenant-Lizenz nach MS365-Login. Legacy-Tab #access → Lizenzen.
 */

function $(id) {
    return document.getElementById(id);
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
    if (saveLabel) saveLabel.textContent = 'Lokal speichern';
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
    if (saveLabel) saveLabel.textContent = 'Lokal aktualisieren';
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
    overview: {
        eyebrow: 'Betreiber',
        title: 'Überblick',
        subtitle: 'Ablauf, Lücken und Kontakte der Schulen.'
    },
    licenses: {
        eyebrow: 'Betreiber',
        title: 'Schulen',
        subtitle: 'Schulen finden, öffnen und freischalten.'
    },
    modules: {
        eyebrow: 'Produkt',
        title: 'Pakete',
        subtitle: 'Paket wählen und Werkzeuge zuordnen. Schulen schalten Pakete unter Lizenzen frei.'
    },
    templates: {
        eyebrow: 'Katalog',
        title: 'Kursteam-Vorlagen',
        subtitle: 'Zentrale Kanal-Vorlagen und Materialien für alle Schulen.'
    },
    notes: {
        eyebrow: 'Kommunikation',
        title: 'Neuigkeiten',
        subtitle: 'Release-Notes mit Editor und Screenshots – auch automatisch aus GitHub-Commits.'
    },
    feedback: {
        eyebrow: 'Support',
        title: 'Feedback',
        subtitle: 'Inbox mit Screenshots – Status führen, optional per Mail antworten.'
    },
    setup: {
        eyebrow: 'Technik',
        title: 'Setup & Diagnose',
        subtitle: 'SharePoint-Schema und License-API prüfen – Freischaltungen unter Lizenzen.'
    }
};

/** Legacy PIN-Tab (#access) – entfernt in Welle 0. */
const ADMIN_TAB_ALIASES = { access: 'licenses' };

const ADMIN_TAB_STORAGE_KEY = 'ms365-admin-active-tab-v1';

/**
 * @param {string} tabId
 * @returns {string}
 */
function normalizeAdminTab(tabId) {
    const raw = String(tabId || '')
        .trim()
        .toLowerCase();
    const mapped = ADMIN_TAB_ALIASES[raw] || raw;
    return ADMIN_TAB_META[mapped] ? mapped : 'licenses';
}

/**
 * @param {string} tabId
 * @param {{ resetNested?: boolean }} [opts] resetNested: Nested-Hash (#licenses/school/…) auf Tab-Wurzel
 */
function setAdminTab(tabId, opts) {
    const id = normalizeAdminTab(tabId);
    const meta = ADMIN_TAB_META[id];
    const resetNested = !!(opts && opts.resetNested);

    document.querySelectorAll('.admin-app__nav-btn[data-admin-tab]').forEach((btn) => {
        const selected = btn.getAttribute('data-admin-tab') === id;
        btn.setAttribute('aria-selected', selected ? 'true' : 'false');
        btn.setAttribute('tabindex', selected ? '0' : '-1');
    });

    document.querySelectorAll('.admin-app__panel[data-admin-panel]').forEach((panel) => {
        const active = panel.getAttribute('data-admin-panel') === id;
        panel.classList.toggle('active', active);
        if (active) {
            panel.removeAttribute('hidden');
            panel.setAttribute('tabindex', '-1');
        } else {
            panel.setAttribute('hidden', '');
            panel.removeAttribute('tabindex');
        }
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

    if (id === 'feedback') {
        try {
            if (
                window.ms365AdminFeedback &&
                typeof window.ms365AdminFeedback.onTabActivated === 'function'
            ) {
                window.ms365AdminFeedback.onTabActivated();
            }
        } catch {
            // ignore
        }
    }

    try {
        const cur = String(location.hash || '')
            .replace(/^#/, '')
            .trim()
            .toLowerCase();
        const root = cur.split('/')[0] || '';
        const nested = cur.indexOf('/') >= 0;
        if (root !== id) {
            history.replaceState(null, '', '#' + id);
        } else if (resetNested && nested) {
            history.replaceState(null, '', '#' + id);
            try {
                window.dispatchEvent(new HashChangeEvent('hashchange'));
            } catch {
                window.dispatchEvent(new Event('hashchange'));
            }
        } else if (!cur) {
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
        if (hash) {
            const root = hash.split('/')[0];
            if (root) return normalizeAdminTab(root);
        }
        const q = new URLSearchParams(location.search).get('tab');
        if (q) return normalizeAdminTab(q);
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
        setAdminTab(btn.getAttribute('data-admin-tab'), { resetNested: true });
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
        setAdminTab(buttons[next].getAttribute('data-admin-tab'), { resetNested: true });
    });

    window.addEventListener('hashchange', () => {
        const fromHash = tabFromLocation();
        if (fromHash) setAdminTab(fromHash);
    });

    let initial = tabFromLocation() || 'overview';
    if (!tabFromLocation()) {
        try {
            const stored = localStorage.getItem(ADMIN_TAB_STORAGE_KEY);
            if (stored) initial = normalizeAdminTab(stored);
        } catch {
            // ignore
        }
    }
    setAdminTab(initial);
}

window.ms365AdminSetTab = setAdminTab;

function syncAdminSessionChrome() {
    let loggedIn = false;
    try {
        loggedIn = typeof window.ms365AuthIsLoggedIn === 'function' && !!window.ms365AuthIsLoggedIn();
    } catch {
        loggedIn = false;
    }
    document.body.classList.toggle('admin-session', loggedIn);
}

function init() {
    initAdminTabs();
    syncAdminSessionChrome();
    window.addEventListener('ms365-auth-state-changed', syncAdminSessionChrome);
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
    }
    refreshNotes = refresh;

    refresh();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();

