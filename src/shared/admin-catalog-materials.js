/**
 * Admin: Materialien (Ordner/Dateien) in MS365-Katalog/materialien.
 */
import {
    fetchMaterials,
    createMaterialFolder,
    uploadMaterialFile,
    deleteMaterialItem
} from '../tools/kursteam-templates/kursteam-templates-catalog.js';
import { defaultMaterialsPath, normalizeMaterialsPath } from '../tools/kursteam-templates/kursteam-templates-logic.js';

function $(id) {
    return document.getElementById(id);
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

const mat = {
    path: 'materialien',
    items: [],
    busy: false
};

function setMatBanner(text) {
    const el = $('adminMatBanner');
    if (!el) return;
    if (!text) {
        el.hidden = true;
        el.textContent = '';
        return;
    }
    el.hidden = false;
    el.textContent = text;
}

function setMatProgress(opts) {
    const box = $('adminMatProgress');
    const fill = $('adminMatProgressFill');
    const text = $('adminMatProgressText');
    if (!box || !text) return;
    if (opts && opts.hidden) {
        box.hidden = true;
        box.removeAttribute('data-state');
        if (fill) fill.style.width = '0%';
        text.textContent = '';
        return;
    }
    box.hidden = false;
    const state = (opts && opts.state) || '';
    if (state) box.setAttribute('data-state', state);
    else box.removeAttribute('data-state');
    if (fill && opts && Number.isFinite(opts.pct)) {
        fill.style.width = Math.max(0, Math.min(100, opts.pct)) + '%';
    }
    text.textContent = (opts && opts.message) || '';
}

function formatBytes(n) {
    const v = Number(n) || 0;
    if (v < 1024) return v + ' B';
    if (v < 1024 * 1024) return (v / 1024).toFixed(1) + ' KB';
    return (v / (1024 * 1024)).toFixed(1) + ' MB';
}

function renderMatList() {
    const ul = $('adminMatList');
    const pathEl = $('adminMatPath');
    const up = $('adminMatUp');
    if (pathEl) pathEl.textContent = mat.path;
    if (up) up.hidden = !mat.path || mat.path === 'materialien';
    if (!ul) return;
    if (!mat.items.length) {
        ul.innerHTML =
            '<li class="admin-catalog__empty" style="padding:14px;">Ordner ist leer. Neuen Ordner anlegen oder Datei hochladen.</li>';
        return;
    }
    ul.innerHTML = mat.items
        .map((item) => {
            const meta = item.isFolder ? 'Ordner' : formatBytes(item.size);
            const icon = item.isFolder ? 'bi-folder' : 'bi-file-earmark';
            return (
                '<li>' +
                '<button type="button" class="admin-mat__open" data-mat-path="' +
                escapeHtml(item.path) +
                '" data-mat-folder="' +
                (item.isFolder ? '1' : '0') +
                '"><span><i class="bi ' +
                icon +
                '"></i> ' +
                escapeHtml(item.name) +
                '</span><span class="muted">' +
                escapeHtml(meta) +
                '</span></button>' +
                '<button type="button" class="btn btn-sm alt" data-mat-del="' +
                escapeHtml(item.path) +
                '" title="Löschen"><i class="bi bi-trash"></i></button>' +
                '</li>'
            );
        })
        .join('');
}

async function loadMat(path) {
    setMatBanner('');
    const data = await fetchMaterials(path || 'materialien');
    mat.path = data.path || path || 'materialien';
    mat.items = Array.isArray(data.items) ? data.items : [];
    if (data.missing) {
        setMatBanner(data.message || 'Ordner fehlt noch – mit „Neuer Ordner“ oder über eine Vorlage anlegen.');
    }
    renderMatList();
    return data;
}

function setCatTab(tab) {
    const id = tab === 'materials' ? 'materials' : 'channels';
    document.querySelectorAll('[data-admin-cat-tab]').forEach((btn) => {
        const on = btn.getAttribute('data-admin-cat-tab') === id;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    document.querySelectorAll('[data-admin-cat-panel]').forEach((panel) => {
        panel.hidden = panel.getAttribute('data-admin-cat-panel') !== id;
    });
    const chTools = $('adminCatToolbarChannels');
    const matTools = $('adminCatToolbarMaterials');
    if (chTools) chTools.hidden = id !== 'channels';
    if (matTools) matTools.hidden = id !== 'materials';
    if (id === 'materials') {
        loadMat(mat.path || 'materialien').catch((err) => {
            setMatBanner((err && err.message) || String(err));
        });
    }
}

/**
 * Von der Kanal-Vorlage: Ordner sicherstellen und Materialien-Tab öffnen.
 * @param {string} templateId
 * @param {string} [preferredPath]
 */
export async function openMaterialsForTemplate(templateId, preferredPath) {
    const want =
        normalizeMaterialsPath(preferredPath) || defaultMaterialsPath(templateId || 'vorlage');
    setCatTab('materials');
    setMatProgress({ pct: 20, message: 'Ordner wird angelegt bzw. geöffnet …' });
    try {
        const data = await createMaterialFolder({ ensurePath: want });
        mat.path = data.path || want;
        mat.items = Array.isArray(data.items) ? data.items : [];
        setMatBanner('Ordner für die Vorlage: ' + mat.path);
        setMatProgress({ state: 'ok', pct: 100, message: 'Ordner bereit.' });
        renderMatList();
        return mat.path;
    } catch (err) {
        setMatProgress({
            state: 'error',
            pct: 100,
            message: (err && err.message) || String(err)
        });
        throw err;
    }
}

function bind() {
    document.querySelectorAll('[data-admin-cat-tab]').forEach((btn) => {
        btn.addEventListener('click', () => {
            setCatTab(btn.getAttribute('data-admin-cat-tab') || 'channels');
        });
    });

    $('adminMatReload')?.addEventListener('click', () => {
        loadMat(mat.path).catch((err) => setMatBanner((err && err.message) || String(err)));
    });
    $('adminMatUp')?.addEventListener('click', () => {
        const parts = String(mat.path || 'materialien').split('/').filter(Boolean);
        if (parts.length <= 1) return;
        parts.pop();
        loadMat(parts.join('/')).catch((err) => setMatBanner((err && err.message) || String(err)));
    });
    $('adminMatNewFolder')?.addEventListener('click', async () => {
        const name = window.prompt('Name des neuen Ordners:', '');
        if (name == null) return;
        const trimmed = String(name).trim();
        if (!trimmed) return;
        try {
            setMatProgress({ pct: 30, message: 'Ordner wird angelegt …' });
            const data = await createMaterialFolder({ parentPath: mat.path, name: trimmed });
            mat.path = data.path || mat.path;
            mat.items = Array.isArray(data.items) ? data.items : [];
            // createMaterialFolder returns listing of parent – reload parent then enter new folder
            await loadMat(mat.path);
            const child = (mat.path.replace(/\/+$/, '') + '/' + trimmed).replace(/\/+/g, '/');
            await loadMat(child);
            setMatProgress({ state: 'ok', pct: 100, message: 'Ordner angelegt.' });
        } catch (err) {
            setMatProgress({
                state: 'error',
                pct: 100,
                message: (err && err.message) || String(err)
            });
        }
    });
    $('adminMatUpload')?.addEventListener('change', async (e) => {
        const file = e.target.files && e.target.files[0];
        e.target.value = '';
        if (!file) return;
        const target = (mat.path.replace(/\/+$/, '') + '/' + file.name).replace(/\/+/g, '/');
        try {
            mat.busy = true;
            setMatProgress({ pct: 15, message: 'Lade „' + file.name + '“ hoch …' });
            const buf = new Uint8Array(await file.arrayBuffer());
            setMatProgress({ pct: 55, message: 'Schreibe nach SharePoint …' });
            await uploadMaterialFile(target, buf, file.type || 'application/octet-stream');
            await loadMat(mat.path);
            setMatProgress({
                state: 'ok',
                pct: 100,
                message: 'Hochgeladen: ' + file.name
            });
        } catch (err) {
            setMatProgress({
                state: 'error',
                pct: 100,
                message: (err && err.message) || String(err)
            });
        } finally {
            mat.busy = false;
        }
    });
    $('adminMatList')?.addEventListener('click', async (e) => {
        const del = e.target.closest('[data-mat-del]');
        if (del) {
            const path = del.getAttribute('data-mat-del') || '';
            if (!path) return;
            if (!window.confirm('„' + path.split('/').pop() + '“ wirklich löschen?')) return;
            try {
                const result = await deleteMaterialItem(path);
                await loadMat(result.parent || mat.path);
                setMatProgress({ state: 'ok', pct: 100, message: 'Gelöscht.' });
            } catch (err) {
                setMatProgress({
                    state: 'error',
                    pct: 100,
                    message: (err && err.message) || String(err)
                });
            }
            return;
        }
        const open = e.target.closest('[data-mat-path]');
        if (!open) return;
        const path = open.getAttribute('data-mat-path') || '';
        const isFolder = open.getAttribute('data-mat-folder') === '1';
        if (!isFolder) return;
        loadMat(path).catch((err) => setMatBanner((err && err.message) || String(err)));
    });

    window.ms365AdminCatalogMaterials = {
        openForTemplate: openMaterialsForTemplate,
        setTab: setCatTab,
        reload: () => loadMat(mat.path)
    };
}

function init() {
    if (!$('adminPanelTemplates')) return;
    bind();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
