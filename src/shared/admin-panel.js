import {
    loadReleaseNotes,
    saveReleaseNotes,
    loadMergedReleaseNotes,
    sanitizeReleaseHtml
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

function selectedAudience() {
    const el = document.querySelector('input[name="adminRnAudience"]:checked');
    return el ? String(el.value) : 'all';
}

function setSelectedAudience(audience) {
    const want = String(audience || 'all');
    document.querySelectorAll('input[name="adminRnAudience"]').forEach((input) => {
        input.checked = input.value === want;
    });
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
    setSelectedAudience('all');
    if ($('adminRnEditingId')) $('adminRnEditingId').value = '';
    const heading = $('adminRnEditorHeading');
    if (heading) heading.textContent = 'Neuer Eintrag';
    const saveLabel = $('adminRnSaveLabel');
    if (saveLabel) saveLabel.textContent = 'Veröffentlichen';
    const cancel = $('adminRnCancelEditBtn');
    if (cancel) cancel.hidden = true;
}

function fillEditor(note) {
    if ($('adminReleaseTitleInput')) $('adminReleaseTitleInput').value = note.title || '';
    setEditorHtml(note.bodyHtml || '');
    draftImages = Array.isArray(note.images) ? note.images.map((x) => ({ src: x.src, alt: x.alt || '' })) : [];
    renderImagePreviews();
    setSelectedKind(note.kind || 'feature');
    setSelectedAudience(note.audience || 'all');
    if ($('adminRnEditingId')) $('adminRnEditingId').value = note.id || '';
    const heading = $('adminRnEditorHeading');
    if (heading) heading.textContent = 'Eintrag bearbeiten';
    const saveLabel = $('adminRnSaveLabel');
    if (saveLabel) saveLabel.textContent = 'Änderung veröffentlichen';
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
                        const title = note && note.title ? note.title : 'Diesen Eintrag';
                        if (!confirm('„' + title + '“ von der Homepage und aus der App entfernen?')) return;
                        const next = notesCache.filter((n) => n.id !== note.id);
                        const editingId = $('adminRnEditingId') ? String($('adminRnEditingId').value || '') : '';
                        publishNotes(next, 'Eintrag ist entfernt. Homepage und App zeigen den neuen Stand.').then((ok) => {
                            if (ok && editingId && note && editingId === note.id) resetEditor();
                        });
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

function setPublishStatus(text) {
    const el = $('adminRnStatus');
    if (el) el.textContent = text || '';
}

function editorDraft() {
    const title = $('adminReleaseTitleInput') ? $('adminReleaseTitleInput').value : '';
    const t = String(title || '').trim();
    const bodyHtml = getEditorHtml();
    const plain = String(editorEl() ? editorEl().innerText : '').trim();
    if (!t && !bodyHtml && !plain && !draftImages.length) return null;
    if (!t || (!bodyHtml && !plain && !draftImages.length)) {
        const err = new Error('Bitte Titel und Text (oder Screenshot) eingeben.');
        err.status = 400;
        throw err;
    }
    const editingId = $('adminRnEditingId') ? String($('adminRnEditingId').value || '').trim() : '';
    const existing = editingId ? notesCache.find((n) => n.id === editingId) : null;
    return {
        id: editingId,
        title: t,
        bodyHtml: bodyHtml || '<p>' + plain.replace(/</g, '&lt;') + '</p>',
        kind: selectedKind(),
        audience: selectedAudience(),
        source: 'manual',
        images: draftImages.slice(),
        at: existing && existing.at ? existing.at : new Date().toISOString()
    };
}

function listWithDraft(draft) {
    const notes = notesCache.slice();
    if (!draft) return notes;
    if (draft.id) {
        const idx = notes.findIndex((n) => n.id === draft.id);
        if (idx >= 0) notes[idx] = Object.assign({}, notes[idx], draft);
        else notes.unshift(draft);
        return notes;
    }
    notes.unshift(Object.assign({}, draft, { id: 'rn_' + Date.now().toString(36) }));
    return notes;
}

async function publishNotes(notes, okMessage) {
    const api = window.ms365LicenseApi;
    const cfg = window.MS365_LICENSE_API;
    if (!api || !cfg || !String(cfg.baseUrl || '').trim() || typeof api.publishReleaseNotes !== 'function') {
        alert('Veröffentlichen braucht die verbundene License-API. In dieser Ansicht ist sie nicht erreichbar.');
        return false;
    }
    const btn = $('adminAddReleaseNoteBtn');
    if (btn) btn.disabled = true;
    setPublishStatus('Wird veröffentlicht …');
    try {
        const token = await api.acquireLicenseToken({ popup: true });
        const result = await api.publishReleaseNotes(token, notes);
        saveReleaseNotes(notes, localStorage);
        const count = Array.isArray(notes) ? notes.length : 0;
        const pub =
            result && result.publicCount != null
                ? result.publicCount
                : notes.filter((n) => !n.audience || n.audience === 'all').length;
        setPublishStatus(
            okMessage ||
                'Veröffentlicht. ' + count + ' Einträge in der App, ' + pub + ' davon auf der Homepage.'
        );
        await refreshNotes();
        return true;
    } catch (e) {
        setPublishStatus('');
        alert('Veröffentlichen fehlgeschlagen: ' + (e && e.message ? e.message : e));
        return false;
    } finally {
        if (btn) btn.disabled = false;
    }
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
        btnAddNote.addEventListener('click', async function () {
            let draft = null;
            try {
                draft = editorDraft();
            } catch (e) {
                alert(e && e.message ? e.message : e);
                return;
            }
            if (!draft) {
                alert('Bitte Titel und Text (oder Screenshot) eingeben.');
                return;
            }
            const editing = !!draft.id;
            const where =
                draft.audience === 'it'
                    ? 'in der App für die Schul-IT'
                    : draft.audience === 'enduser'
                      ? 'in der App für Endnutzer'
                      : 'in der App und auf der Homepage';
            const ok = await publishNotes(
                listWithDraft(draft),
                (editing ? 'Änderung ist ' : 'Eintrag ist ') + where + '.'
            );
            if (ok) resetEditor();
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

