/**
 * Admin: Materialien (Ordner/Dateien) in MS365-Katalog/materialien.
 * Explorer-Layout: links Ordnerbaum, rechts Inhalt des gewählten Ordners.
 */
import {
    fetchMaterials,
    createMaterialFolder,
    uploadMaterialFile,
    deleteMaterialItem
} from '../tools/kursteam-templates/kursteam-templates-catalog.js';
import { defaultMaterialsPath, normalizeMaterialsPath } from '../tools/kursteam-templates/kursteam-templates-logic.js';

const ROOT = 'materialien';

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

/** @type {{ path: string, items: Array, busy: boolean, cache: Map<string, {folders: Array, files: Array}>, expanded: Set<string> }} */
const mat = {
    path: ROOT,
    items: [],
    busy: false,
    cache: new Map(),
    expanded: new Set([ROOT])
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

function folderName(path) {
    const parts = String(path || '').split('/').filter(Boolean);
    return parts[parts.length - 1] || ROOT;
}

function parentPath(path) {
    const parts = String(path || ROOT).split('/').filter(Boolean);
    if (parts.length <= 1) return ROOT;
    parts.pop();
    return parts.join('/');
}

function ancestorPaths(path) {
    const parts = String(path || ROOT).split('/').filter(Boolean);
    const out = [];
    for (let i = 1; i <= parts.length; i++) out.push(parts.slice(0, i).join('/'));
    return out;
}

function rememberListing(path, items) {
    const folders = [];
    const files = [];
    (items || []).forEach((item) => {
        if (item.isFolder) folders.push(item);
        else files.push(item);
    });
    folders.sort((a, b) => a.name.localeCompare(b.name, 'de'));
    files.sort((a, b) => a.name.localeCompare(b.name, 'de'));
    mat.cache.set(path, { folders, files });
}

function invalidateSubtree(path) {
    const prefix = String(path || ROOT).replace(/\/+$/, '');
    for (const key of [...mat.cache.keys()]) {
        if (key === prefix || key.startsWith(prefix + '/')) mat.cache.delete(key);
    }
}

function renderTree() {
    const nav = $('adminMatTree');
    if (!nav) return;

    function nodeHtml(path, depth) {
        const cached = mat.cache.get(path);
        const folders = cached ? cached.folders : [];
        const expanded = mat.expanded.has(path);
        const selected = mat.path === path;
        const hasKids = !cached || folders.length > 0;
        const chevron = !hasKids && cached
            ? '<span class="admin-mat-tree__spacer" aria-hidden="true"></span>'
            : '<button type="button" class="admin-mat-tree__toggle" data-mat-toggle="' +
              escapeHtml(path) +
              '" aria-expanded="' +
              (expanded ? 'true' : 'false') +
              '" title="' +
              (expanded ? 'Zuklappen' : 'Aufklappen') +
              '"><i class="bi bi-caret-' +
              (expanded ? 'down' : 'right') +
              '-fill" aria-hidden="true"></i></button>';

        let html =
            '<div class="admin-mat-tree__row' +
            (selected ? ' is-selected' : '') +
            '" style="--depth:' +
            depth +
            '">' +
            chevron +
            '<button type="button" class="admin-mat-tree__label" data-mat-select="' +
            escapeHtml(path) +
            '"><i class="bi ' +
            (expanded || selected ? 'bi-folder2-open' : 'bi-folder') +
            '" aria-hidden="true"></i><span>' +
            escapeHtml(folderName(path)) +
            '</span></button></div>';

        if (expanded && folders.length) {
            html +=
                '<div class="admin-mat-tree__children">' +
                folders.map((f) => nodeHtml(f.path, depth + 1)).join('') +
                '</div>';
        } else if (expanded && !cached) {
            html +=
                '<div class="admin-mat-tree__children"><div class="admin-mat-tree__loading" style="--depth:' +
                (depth + 1) +
                '">Lädt …</div></div>';
        }
        return html;
    }

    nav.innerHTML = nodeHtml(ROOT, 0);
}

function renderFilePane() {
    const ul = $('adminMatList');
    const pathEl = $('adminMatPath');
    const titleEl = $('adminMatPaneTitle');
    const metaEl = $('adminMatPaneMeta');
    if (pathEl) pathEl.textContent = mat.path;
    if (titleEl) titleEl.textContent = folderName(mat.path);

    const folders = mat.items.filter((i) => i.isFolder);
    const files = mat.items.filter((i) => !i.isFolder);
    if (metaEl) {
        const bits = [];
        if (folders.length) bits.push(folders.length + ' Ordner');
        if (files.length) bits.push(files.length + ' Datei' + (files.length === 1 ? '' : 'en'));
        metaEl.textContent = bits.length ? bits.join(' · ') : 'leer';
    }

    if (!ul) return;
    if (!mat.items.length) {
        ul.innerHTML =
            '<li class="admin-mat__empty">Dieser Ordner ist leer. Neuen Ordner anlegen oder Datei hochladen.</li>';
        return;
    }

    const rows = [];
    folders.forEach((item) => {
        rows.push(
            '<li class="admin-mat__item admin-mat__item--folder">' +
                '<button type="button" class="admin-mat__open" data-mat-path="' +
                escapeHtml(item.path) +
                '" data-mat-folder="1"><span><i class="bi bi-folder" aria-hidden="true"></i> ' +
                escapeHtml(item.name) +
                '</span><span class="muted">Ordner</span></button>' +
                '<button type="button" class="btn btn-sm alt" data-mat-del="' +
                escapeHtml(item.path) +
                '" title="Löschen"><i class="bi bi-trash"></i></button></li>'
        );
    });
    files.forEach((item) => {
        rows.push(
            '<li class="admin-mat__item admin-mat__item--file">' +
                '<span class="admin-mat__file"><span><i class="bi bi-file-earmark" aria-hidden="true"></i> ' +
                escapeHtml(item.name) +
                '</span><span class="muted">' +
                escapeHtml(formatBytes(item.size)) +
                '</span></span>' +
                '<button type="button" class="btn btn-sm alt" data-mat-del="' +
                escapeHtml(item.path) +
                '" title="Löschen"><i class="bi bi-trash"></i></button></li>'
        );
    });
    ul.innerHTML = rows.join('');
}

function renderMatUi() {
    renderTree();
    renderFilePane();
}

async function ensureFolderCached(path) {
    if (mat.cache.has(path)) return mat.cache.get(path);
    const data = await fetchMaterials(path || ROOT);
    rememberListing(data.path || path || ROOT, data.items || []);
    return mat.cache.get(data.path || path || ROOT);
}

async function expandToPath(path) {
    const want = path || ROOT;
    for (const p of ancestorPaths(want)) {
        mat.expanded.add(p);
        try {
            await ensureFolderCached(p);
        } catch {
            /* Baum so weit wie möglich */
        }
    }
}

async function loadMat(path, opts) {
    const target = path || ROOT;
    setMatBanner('');
    const data = await fetchMaterials(target);
    mat.path = data.path || target;
    mat.items = Array.isArray(data.items) ? data.items : [];
    rememberListing(mat.path, mat.items);
    mat.expanded.add(mat.path);
    if (opts && opts.expandAncestors) {
        await expandToPath(mat.path);
    } else {
        // Elternknoten brauchen Einträge, sonst fehlt der Pfad im Baum
        const parent = parentPath(mat.path);
        if (parent !== mat.path && !mat.cache.has(parent)) {
            try {
                await ensureFolderCached(parent);
            } catch {
                /* ignore */
            }
        }
        ancestorPaths(mat.path).forEach((p) => mat.expanded.add(p));
    }
    if (data.missing) {
        setMatBanner(data.message || 'Ordner fehlt noch – mit „Neuer Ordner“ oder über eine Vorlage anlegen.');
    }
    renderMatUi();
    return data;
}

async function toggleFolder(path) {
    if (mat.expanded.has(path)) {
        mat.expanded.delete(path);
        renderTree();
        return;
    }
    mat.expanded.add(path);
    renderTree();
    try {
        await ensureFolderCached(path);
        renderTree();
    } catch (err) {
        mat.expanded.delete(path);
        setMatBanner((err && err.message) || String(err));
        renderTree();
    }
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
    if (id === 'materials') {
        loadMat(mat.path || ROOT, { expandAncestors: true }).catch((err) => {
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
        invalidateSubtree(ROOT);
        mat.path = data.path || want;
        mat.items = Array.isArray(data.items) ? data.items : [];
        rememberListing(mat.path, mat.items);
        await expandToPath(mat.path);
        setMatBanner('Ordner für die Vorlage: ' + mat.path);
        setMatProgress({ state: 'ok', pct: 100, message: 'Ordner bereit.' });
        renderMatUi();
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
        invalidateSubtree(mat.path);
        loadMat(mat.path, { expandAncestors: true }).catch((err) =>
            setMatBanner((err && err.message) || String(err))
        );
    });

    $('adminMatNewFolder')?.addEventListener('click', async () => {
        const name = window.prompt('Name des neuen Ordners:', '');
        if (name == null) return;
        const trimmed = String(name).trim();
        if (!trimmed) return;
        try {
            setMatProgress({ pct: 30, message: 'Ordner wird angelegt …' });
            await createMaterialFolder({ parentPath: mat.path, name: trimmed });
            invalidateSubtree(mat.path);
            const child = (mat.path.replace(/\/+$/, '') + '/' + trimmed).replace(/\/+/g, '/');
            await loadMat(child, { expandAncestors: true });
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
            invalidateSubtree(mat.path);
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

    $('adminMatTree')?.addEventListener('click', async (e) => {
        const toggle = e.target.closest('[data-mat-toggle]');
        if (toggle) {
            e.preventDefault();
            const path = toggle.getAttribute('data-mat-toggle') || ROOT;
            await toggleFolder(path);
            return;
        }
        const select = e.target.closest('[data-mat-select]');
        if (!select) return;
        const path = select.getAttribute('data-mat-select') || ROOT;
        loadMat(path).catch((err) => setMatBanner((err && err.message) || String(err)));
    });

    $('adminMatList')?.addEventListener('click', async (e) => {
        const del = e.target.closest('[data-mat-del]');
        if (del) {
            const path = del.getAttribute('data-mat-del') || '';
            if (!path) return;
            if (!window.confirm('„' + path.split('/').pop() + '“ wirklich löschen?')) return;
            try {
                const result = await deleteMaterialItem(path);
                const parent = result.parent || parentPath(path);
                invalidateSubtree(parent);
                if (mat.path === path || mat.path.startsWith(path + '/')) {
                    await loadMat(parent, { expandAncestors: true });
                } else {
                    await loadMat(mat.path);
                }
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
        reload: () => loadMat(mat.path, { expandAncestors: true })
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
