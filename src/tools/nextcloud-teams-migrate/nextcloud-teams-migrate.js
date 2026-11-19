/**
 * Nextcloud → Teams UI: lokaler Helfer (WebDAV) + ZIP/Ordner-Fallback.
 */
import { escapeHtml } from '../../shared/utils/strings.js';
import {
    formatBytes,
    suggestChannelMappings,
    buildMigrationPlan,
    validateNcMigrateSelection,
    normalizeNextcloudBase,
    norm
} from './nextcloud-teams-migrate-logic.js';
import { listChildren, inventoryFiles, downloadBlob } from './nextcloud-teams-migrate-webdav.js';
import {
    listVirtualChildren,
    inventoryVirtual,
    readVirtualBlob,
    buildFsFromFileList,
    buildFsFromJsZip,
    folderSamplesFromSelection,
    isLocalSourceMode
} from './nextcloud-teams-migrate-local.js';
import {
    helperHealth,
    helperListChildren,
    helperInventoryFiles,
    helperDownloadBlob,
    DEFAULT_HELPER_BASE
} from './nextcloud-teams-migrate-helper.js';
import {
    getGroupDrive,
    listChannels,
    ensureChannel,
    getChannelFilesFolder,
    ensureFolderPath,
    ensureChildFolderPath,
    uploadBlob,
    withRetry
} from './nextcloud-teams-migrate-graph.js';

const LS_KEY = 'ms365-nc-teams-migrate-v1';
const SS_PASS = 'ms365-nc-teams-migrate-pass';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.Read.All',
    'https://graph.microsoft.com/Channel.ReadBasic.All',
    'https://graph.microsoft.com/Channel.Create',
    'https://graph.microsoft.com/Files.ReadWrite.All',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

function $(id) {
    return document.getElementById(id);
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function log(msg) {
    const el = $('ncLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(msg || '');
}

const state = {
    /** @type {'helper'|'zip'|'folder'|'webdav'} */
    sourceMode: 'helper',
    connected: false,
    helperBase: DEFAULT_HELPER_BASE,
    helperUser: '',
    /** @type {import('./nextcloud-teams-migrate-local.js').VirtualFs | null} */
    vfs: null,
    crumbs: /** @type {Array<{ relPath: string, name: string }>} */ ([
        { relPath: '', name: 'Stamm' }
    ]),
    items: /** @type {Array<{ relPath: string, name: string, isFolder: boolean, size: number }>} */ ([]),
    selected: /** @type {Set<string>} */ (new Set()),
    dest: { groupId: '', label: '', driveId: '' },
    channelMaps: /** @type {ReturnType<typeof suggestChannelMappings>} */ ([]),
    lastPlan: /** @type {ReturnType<typeof buildMigrationPlan> | null} */ (null),
    running: false,
    abort: false
};

function loadSavedConn() {
    try {
        const raw = localStorage.getItem(LS_KEY);
        if (!raw) return;
        const o = JSON.parse(raw);
        if ($('ncUrl') && o.url) $('ncUrl').value = o.url;
        if ($('ncUser') && o.user) $('ncUser').value = o.user;
        if ($('ncProxy') && o.proxy) $('ncProxy').value = o.proxy;
        if ($('ncDestRoot') && o.destRoot) $('ncDestRoot').value = o.destRoot;
        if ($('ncHelperUrl') && o.helperUrl) $('ncHelperUrl').value = o.helperUrl;
        if ($('ncMapChannels') && typeof o.mapChannels === 'boolean') {
            $('ncMapChannels').checked = o.mapChannels;
        }
        if (
            o.sourceMode === 'helper' ||
            o.sourceMode === 'zip' ||
            o.sourceMode === 'folder' ||
            o.sourceMode === 'webdav'
        ) {
            setSourceMode(o.sourceMode === 'webdav' ? 'helper' : o.sourceMode, false);
        }
    } catch {
        /* ignore */
    }
    try {
        const p = sessionStorage.getItem(SS_PASS);
        if (p && $('ncPass')) $('ncPass').value = p;
    } catch {
        /* ignore */
    }
}

function saveConn() {
    const payload = {
        sourceMode: state.sourceMode,
        helperUrl: ($('ncHelperUrl') && $('ncHelperUrl').value) || DEFAULT_HELPER_BASE,
        url: ($('ncUrl') && $('ncUrl').value) || '',
        user: ($('ncUser') && $('ncUser').value) || '',
        proxy: ($('ncProxy') && $('ncProxy').value) || '',
        destRoot: ($('ncDestRoot') && $('ncDestRoot').value) || 'Import aus Nextcloud',
        mapChannels: !!( $('ncMapChannels') && $('ncMapChannels').checked )
    };
    localStorage.setItem(LS_KEY, JSON.stringify(payload));
    try {
        const pass = ($('ncPass') && $('ncPass').value) || '';
        if (pass) sessionStorage.setItem(SS_PASS, pass);
        else sessionStorage.removeItem(SS_PASS);
    } catch {
        /* ignore */
    }
}

function connOpts() {
    return {
        baseUrl: normalizeNextcloudBase(($('ncUrl') && $('ncUrl').value) || ''),
        user: norm(($('ncUser') && $('ncUser').value) || ''),
        appPassword: norm(($('ncPass') && $('ncPass').value) || ''),
        corsProxy: norm(($('ncProxy') && $('ncProxy').value) || '')
    };
}

async function getToken() {
    if (typeof window.ms365AuthAcquireToken === 'function') {
        return window.ms365AuthAcquireToken(SCOPES);
    }
    const G = window.ms365GraphUnifiedGroups;
    if (G && typeof G.getGraphToken === 'function') return G.getGraphToken(SCOPES);
    throw new Error('Anmeldung (MSAL) nicht geladen – bitte oben bei Microsoft anmelden.');
}

function guidLike(s) {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
        String(s || '').trim()
    );
}

async function searchTeams(query) {
    const q = String(query || '').trim();
    if (!q) throw new Error('Suchbegriff fehlt.');
    const token = await getToken();
    const G = window.ms365GraphUnifiedGroups;
    if (!G) throw new Error('Graph-Gruppenmodul fehlt.');
    if (guidLike(q)) {
        const g = await G.fetchGroup(token, q);
        if (!g || !g.id) throw new Error('Gruppe nicht gefunden.');
        return [
            {
                id: g.id,
                displayName: g.displayName || '',
                mail: g.mail || '',
                mailNickname: g.mailNickname || ''
            }
        ];
    }
    const list = await G.searchUnifiedGroups(token, q);
    return (list || []).map(function (g) {
        return {
            id: g.id,
            displayName: g.displayName || '',
            mail: g.mail || '',
            mailNickname: g.mailNickname || ''
        };
    });
}

function currentRel() {
    const c = state.crumbs;
    return c.length ? c[c.length - 1].relPath : '';
}

function setHint(text, ok) {
    const el = $('ncHint');
    if (!el) return;
    el.textContent = text;
    if (el.parentElement) el.parentElement.classList.toggle('tm-perm--ok', !!ok);
}

/**
 * @param {'helper'|'zip'|'folder'|'webdav'} mode
 * @param {boolean} [persist]
 */
function setSourceMode(mode, persist) {
    state.sourceMode = mode;
    document.querySelectorAll('.nc-mode-tab').forEach(function (btn) {
        btn.classList.toggle('is-active', btn.getAttribute('data-mode') === mode);
    });
    const helper = $('ncPaneHelper');
    const zip = $('ncPaneZip');
    const folder = $('ncPaneFolder');
    if (helper) helper.hidden = mode !== 'helper';
    if (zip) zip.hidden = mode !== 'zip';
    if (folder) folder.hidden = mode !== 'folder';
    if (persist !== false) saveConn();
}

function resetSourceBrowser(message) {
    state.connected = false;
    state.vfs = null;
    state.crumbs = [{ relPath: '', name: 'Stamm' }];
    state.items = [];
    state.selected.clear();
    state.channelMaps = [];
    state.lastPlan = null;
    renderCrumb();
    renderBrowser(message);
    renderSelection();
    renderChannelMaps();
    renderPlanSummary(null);
}

function renderCrumb() {
    const el = $('ncCrumb');
    if (!el) return;
    el.replaceChildren();
    state.crumbs.forEach(function (c, idx) {
        if (idx > 0) {
            const sep = document.createElement('span');
            sep.className = 'sep';
            sep.textContent = '/';
            el.appendChild(sep);
        }
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.textContent = c.name;
        btn.addEventListener('click', function () {
            navigateCrumb(idx).catch(function (e) {
                toast(e.message || String(e));
            });
        });
        el.appendChild(btn);
    });
}

function renderBrowser(emptyMsg) {
    const body = $('ncSourceBody');
    if (!body) return;
    body.replaceChildren();
    if (!state.connected) {
        body.innerHTML =
            '<tr><td colspan="4" class="muted">' +
            escapeHtml(
                emptyMsg ||
                    'ZIP oder Ordner wählen – dann erscheint die Struktur hier.'
            ) +
            '</td></tr>';
        return;
    }
    if (!state.items.length) {
        body.innerHTML = '<tr><td colspan="4" class="muted">Ordner ist leer.</td></tr>';
        return;
    }
    state.items.forEach(function (it) {
        const tr = document.createElement('tr');
        const tdCheck = document.createElement('td');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = state.selected.has(it.relPath);
        cb.addEventListener('change', function () {
            if (cb.checked) state.selected.add(it.relPath);
            else state.selected.delete(it.relPath);
            renderSelection();
            refreshChannelMapsFromSelection().catch(function () {});
        });
        tdCheck.appendChild(cb);
        tr.appendChild(tdCheck);

        const tdName = document.createElement('td');
        if (it.isFolder) {
            const open = document.createElement('button');
            open.type = 'button';
            open.className = 'nc-open';
            open.textContent = it.name;
            open.addEventListener('click', function () {
                openFolder(it).catch(function (e) {
                    toast(e.message || String(e));
                });
            });
            tdName.appendChild(document.createTextNode('📁 '));
            tdName.appendChild(open);
        } else {
            tdName.textContent = it.name;
        }
        tr.appendChild(tdName);

        const tdArt = document.createElement('td');
        tdArt.textContent = it.isFolder ? 'Ordner' : formatBytes(it.size);
        tr.appendChild(tdArt);

        const tdPath = document.createElement('td');
        tdPath.className = 'muted';
        tdPath.style.fontSize = '0.85em';
        tdPath.textContent = it.relPath;
        tr.appendChild(tdPath);
        body.appendChild(tr);
    });
}

function renderSelection() {
    const count = $('ncSelCount');
    const list = $('ncSelList');
    if (count) count.textContent = String(state.selected.size);
    if (!list) return;
    list.replaceChildren();
    Array.from(state.selected)
        .sort(function (a, b) {
            return a.localeCompare(b, 'de');
        })
        .forEach(function (p) {
            const li = document.createElement('li');
            li.textContent = p;
            list.appendChild(li);
        });
}

function renderDest() {
    const el = $('ncDestPicked');
    if (!el) return;
    if (!state.dest.groupId) {
        el.className = 'muted';
        el.textContent = 'Noch kein Ziel-Team gewählt.';
        return;
    }
    el.className = '';
    el.innerHTML =
        '<strong>' +
        escapeHtml(state.dest.label) +
        '</strong><br><span class="muted" style="font-size:0.85em;">' +
        escapeHtml(state.dest.groupId) +
        '</span>';
}

function renderChannelMaps() {
    const wrap = $('ncMapBody');
    if (!wrap) return;
    wrap.replaceChildren();
    if (!state.channelMaps.length) {
        wrap.innerHTML =
            '<tr><td colspan="4" class="muted">Keine nummerierten Unterordner (01 …) erkannt.</td></tr>';
        return;
    }
    state.channelMaps.forEach(function (m, idx) {
        const tr = document.createElement('tr');
        const tdEn = document.createElement('td');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = m.enabled !== false;
        cb.addEventListener('change', function () {
            state.channelMaps[idx].enabled = cb.checked;
        });
        tdEn.appendChild(cb);
        tr.appendChild(tdEn);
        const tdSrc = document.createElement('td');
        tdSrc.textContent = m.sourceFolder;
        tr.appendChild(tdSrc);
        const tdCh = document.createElement('td');
        const inp = document.createElement('input');
        inp.type = 'text';
        inp.value = m.channelName;
        inp.addEventListener('change', function () {
            state.channelMaps[idx].channelName = inp.value.trim().slice(0, 50);
        });
        tdCh.appendChild(inp);
        tr.appendChild(tdCh);
        const tdNo = document.createElement('td');
        tdNo.textContent = m.number;
        tr.appendChild(tdNo);
        wrap.appendChild(tr);
    });
}

function renderPlanSummary(plan) {
    const el = $('ncPlanSummary');
    if (!el) return;
    if (!plan) {
        el.textContent = 'Noch kein Dry-Run.';
        return;
    }
    const s = plan.summary;
    el.innerHTML =
        '<strong>' +
        s.fileCount +
        '</strong> Dateien · ' +
        formatBytes(s.totalBytes) +
        ' · Kanäle: <strong>' +
        s.channelCount +
        '</strong> · Stamm: <strong>' +
        s.teamRootCount +
        '</strong>';
}

function setProgress(done, total) {
    const bar = $('ncProgressBar');
    const label = $('ncProgressLabel');
    const pct = total > 0 ? Math.round((done / total) * 100) : 0;
    if (bar) bar.style.width = pct + '%';
    if (label) label.textContent = total ? done + ' / ' + total + ' (' + pct + '%)' : '—';
}

async function loadChildren(rel) {
    if (isLocalSourceMode(state.sourceMode)) {
        if (!state.vfs) return [];
        return listVirtualChildren(state.vfs, rel);
    }
    if (state.sourceMode === 'helper') {
        return helperListChildren({
            base: state.helperBase,
            relPath: rel,
            userForParse: state.helperUser
        });
    }
    return listChildren({ ...connOpts(), relPath: rel });
}

function mountLocalFs(vfs, label) {
    state.vfs = vfs;
    state.connected = true;
    state.crumbs = [{ relPath: '', name: 'Stamm' }];
    state.items = listVirtualChildren(vfs, '');
    state.selected.clear();
    renderCrumb();
    renderBrowser();
    renderSelection();
    setHint(label + ' geladen – Ordner anklicken und auswählen.', true);
    log(label + ': ' + state.items.length + ' Einträge im Stamm');
    toast(label + ' geladen.');
}

async function onZipChosen(file) {
    if (!file) return;
    setHint('ZIP wird gelesen …', false);
    log('ZIP: ' + file.name + ' (' + formatBytes(file.size) + ')');
    const { loadJsZip } = await import('../../shared/vendor-libs.js');
    const JSZip = await loadJsZip();
    const zip = await JSZip.loadAsync(file);
    const vfs = await buildFsFromJsZip(zip);
    mountLocalFs(vfs, 'ZIP „' + file.name + '“');
}

function onFolderChosen(fileList) {
    if (!fileList || !fileList.length) return;
    const vfs = buildFsFromFileList(fileList);
    mountLocalFs(vfs, 'Ordner (' + fileList.length + ' Dateien)');
}

async function connectHelper() {
    saveConn();
    const base = norm(($('ncHelperUrl') && $('ncHelperUrl').value) || DEFAULT_HELPER_BASE);
    state.helperBase = base;
    setHint('Prüfe lokalen Helfer …', false);
    const statusEl = $('ncHelperStatus');
    if (statusEl) statusEl.textContent = 'Verbinde …';
    const health = await helperHealth(base);
    state.helperUser = norm(health.user) || '';
    if ($('ncUser')) $('ncUser').value = state.helperUser;
    if ($('ncUrl') && health.nextcloudUrl) $('ncUrl').value = health.nextcloudUrl;
    log('Helfer OK: ' + base + ' → ' + (health.nextcloudUrl || '') + ' als ' + state.helperUser);
    const items = await helperListChildren({
        base,
        relPath: '',
        userForParse: state.helperUser
    });
    state.vfs = null;
    state.connected = true;
    state.crumbs = [{ relPath: '', name: 'Stamm' }];
    state.items = items;
    state.selected.clear();
    renderCrumb();
    renderBrowser();
    renderSelection();
    if (statusEl) {
        statusEl.textContent =
            'Verbunden: ' + state.helperUser + ' @ ' + (health.nextcloudUrl || base);
    }
    setHint('Helfer verbunden – Nextcloud-Struktur geladen.', true);
    toast('Lokaler Helfer verbunden.');
}

async function openFolder(it) {
    state.crumbs.push({ relPath: it.relPath, name: it.name });
    state.items = await loadChildren(it.relPath);
    renderCrumb();
    renderBrowser();
}

async function navigateCrumb(index) {
    state.crumbs = state.crumbs.slice(0, index + 1);
    state.items = await loadChildren(currentRel());
    renderCrumb();
    renderBrowser();
}

async function refreshChannelMapsFromSelection() {
    const mapOn = !!( $('ncMapChannels') && $('ncMapChannels').checked );
    if (!mapOn) {
        state.channelMaps = [];
        renderChannelMaps();
        return;
    }
    // Nur die angehakten Ordner selbst → ein Kanal; Unterordner (21_22 …) bleiben Ordner darin
    /** @type {Array<{ name: string, isFolder: boolean, relPath: string }>} */
    const samples =
        isLocalSourceMode(state.sourceMode) && state.vfs
            ? folderSamplesFromSelection(state.vfs, state.selected)
            : Array.from(state.selected).map(function (rel) {
                  const name = rel.split('/').filter(Boolean).pop() || rel;
                  return { name: name, isFolder: true, relPath: rel };
              });
    state.channelMaps = suggestChannelMappings(samples);
    renderChannelMaps();
}

function renderHits(hits) {
    const ul = $('ncDestHits');
    if (!ul) return;
    ul.replaceChildren();
    if (!hits.length) {
        ul.hidden = true;
        return;
    }
    ul.hidden = false;
    hits.forEach(function (g) {
        const li = document.createElement('li');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.innerHTML =
            '<strong>' +
            escapeHtml(g.displayName || g.mailNickname || g.id) +
            '</strong><br><span class="muted" style="font-size:0.85em;">' +
            escapeHtml(g.mailNickname || g.mail || g.id) +
            '</span>';
        btn.addEventListener('click', function () {
            pickDest(g).catch(function (e) {
                toast(e.message || String(e));
                log('FEHLER: ' + (e.message || e));
            });
        });
        li.appendChild(btn);
        ul.appendChild(li);
    });
}

async function pickDest(group) {
    const token = await getToken();
    const drive = await getGroupDrive(token, group.id);
    state.dest.groupId = group.id;
    state.dest.label = group.displayName || group.mailNickname || group.id;
    state.dest.driveId = drive.id;
    const hits = $('ncDestHits');
    if (hits) {
        hits.replaceChildren();
        hits.hidden = true;
    }
    renderDest();
    log('Ziel-Team: ' + state.dest.label);
    toast('Ziel-Team geladen.');
}

async function runSearchDest() {
    const q = norm(($('ncDestQuery') && $('ncDestQuery').value) || '');
    const hits = await searchTeams(q);
    if (!hits.length) {
        renderHits([]);
        toast('Keine Treffer.');
        return;
    }
    renderHits(hits);
    if (hits.length === 1) await pickDest(hits[0]);
    else toast(hits.length + ' Treffer – bitte wählen.');
}

async function collectSelectedFiles() {
    /** @type {Array<{ relPath: string, name: string, size: number }>} */
    const files = [];
    const roots = Array.from(state.selected).sort(function (a, b) {
        return a.localeCompare(b, 'de');
    });

    for (const rel of roots) {
        if (isLocalSourceMode(state.sourceMode)) {
            if (!state.vfs) throw new Error('Keine lokale Quelle geladen.');
            const inv = inventoryVirtual(state.vfs, rel);
            inv.files.forEach((f) =>
                files.push({ relPath: f.relPath, name: f.name, size: f.size || 0 })
            );
            continue;
        }
        const known = state.items.find((i) => i.relPath === rel);
        if (known && !known.isFolder) {
            files.push({ relPath: known.relPath, name: known.name, size: known.size || 0 });
            continue;
        }
        log('Inventar: ' + rel);
        const onProgress = function (p) {
            setHint(
                'Inventar „' + rel + '“: ' + p.files + ' Dateien, ' + formatBytes(p.bytes),
                false
            );
        };
        const inv =
            state.sourceMode === 'helper'
                ? await helperInventoryFiles({
                      base: state.helperBase,
                      relPath: rel,
                      userForParse: state.helperUser,
                      onProgress
                  })
                : await inventoryFiles({
                      ...connOpts(),
                      relPath: rel,
                      onProgress
                  });
        inv.files.forEach((f) => files.push(f));
    }
    /** @type {Map<string, { relPath: string, name: string, size: number }>} */
    const map = new Map();
    files.forEach((f) => map.set(f.relPath, f));
    return { files: Array.from(map.values()), roots };
}

function commonSourceRoot(roots) {
    if (!roots.length) return '';
    if (roots.length === 1) return roots[0];
    const split = roots.map((r) => r.split('/').filter(Boolean));
    const minLen = Math.min(...split.map((p) => p.length));
    /** @type {string[]} */
    const common = [];
    for (let i = 0; i < minLen; i++) {
        const part = split[0][i];
        if (split.every((p) => p[i] === part)) common.push(part);
        else break;
    }
    return common.join('/');
}

async function readSourceBlob(relPath) {
    if (isLocalSourceMode(state.sourceMode)) {
        if (!state.vfs) throw new Error('Keine lokale Quelle.');
        return readVirtualBlob(state.vfs, relPath);
    }
    if (state.sourceMode === 'helper') {
        return helperDownloadBlob({ base: state.helperBase, relPath });
    }
    return downloadBlob({ ...connOpts(), relPath });
}

async function runDryRun() {
    const opts = connOpts();
    const v = validateNcMigrateSelection({
        sourceMode: state.sourceMode,
        hasLocalFs: !!(state.vfs && state.connected),
        hasHelper: state.sourceMode === 'helper' && state.connected,
        nextcloudUrl: opts.baseUrl,
        nextcloudUser: opts.user,
        hasPassword: !!opts.appPassword,
        destGroupId: state.dest.groupId,
        selectedRelPaths: Array.from(state.selected)
    });
    if (!v.ok) throw new Error(v.issues.join('; '));

    saveConn();
    await refreshChannelMapsFromSelection();
    const collected = await collectSelectedFiles();
    const sourceRootRel = commonSourceRoot(collected.roots);
    const plan = buildMigrationPlan({
        sourceRootRel,
        files: collected.files,
        channelMaps: state.channelMaps,
        destRootFolder: norm(($('ncDestRoot') && $('ncDestRoot').value) || 'Import aus Nextcloud'),
        mapNumberedToChannels: !!( $('ncMapChannels') && $('ncMapChannels').checked )
    });
    state.lastPlan = plan;
    renderPlanSummary(plan);
    log(
        'Dry-Run: ' +
            plan.summary.fileCount +
            ' Dateien, ' +
            formatBytes(plan.summary.totalBytes)
    );
    plan.items.slice(0, 12).forEach(function (it) {
        log(
            '  → ' +
                it.sourceRelPath +
                ' → [' +
                (it.destKind === 'channel' ? 'Kanal ' + it.channelName : 'Stamm') +
                '] ' +
                (it.destFolderRel ? it.destFolderRel + '/' : '') +
                it.fileName
        );
    });
    if (plan.items.length > 12) log('  … +' + (plan.items.length - 12) + ' weitere');
    setHint(
        'Dry-Run OK: ' + plan.summary.fileCount + ' Dateien (' + formatBytes(plan.summary.totalBytes) + ')',
        true
    );
    toast('Dry-Run fertig – siehe Protokoll.');
    return plan;
}

async function runMigrate() {
    if (state.running) return;
    let plan = state.lastPlan;
    if (!plan) plan = await runDryRun();
    if (!plan || !plan.items.length) {
        toast('Nichts zu übertragen.');
        return;
    }
    const limitEl = $('ncLimit');
    const limit = limitEl && limitEl.value ? Number(limitEl.value) : 0;
    let items = plan.items;
    if (limit > 0) items = items.slice(0, limit);

    const ok = window.confirm(
        items.length +
            ' Datei(en) nach Team „' +
            state.dest.label +
            '“ übertragen?\n\nTab offen lassen.'
    );
    if (!ok) return;

    state.running = true;
    state.abort = false;
    setProgress(0, items.length);
    const token = await getToken();

    /** @type {Map<string, { driveId: string, folderId: string }>} */
    const channelFolderCache = new Map();
    let existingChannels = await listChannels(token, state.dest.groupId).catch(function () {
        return [];
    });

    let done = 0;
    let failed = 0;
    for (const it of items) {
        if (state.abort) break;
        try {
            await withRetry(async function () {
                /** @type {{ driveId: string, folderId: string }} */
                let target;
                if (it.destKind === 'channel' && it.channelName) {
                    let cached = channelFolderCache.get(it.channelName);
                    if (!cached) {
                        const ch = await ensureChannel(
                            token,
                            state.dest.groupId,
                            it.channelName,
                            existingChannels
                        );
                        if (!existingChannels.find((c) => c.id === ch.id)) {
                            existingChannels = existingChannels.concat([ch]);
                        }
                        const ff = await getChannelFilesFolder(token, state.dest.groupId, ch.id);
                        cached = { driveId: ff.driveId, folderId: ff.folderId };
                        channelFolderCache.set(it.channelName, cached);
                        log('Kanal bereit: ' + it.channelName);
                    }
                    const parent = it.destFolderRel
                        ? await ensureChildFolderPath(
                              token,
                              cached.driveId,
                              cached.folderId,
                              it.destFolderRel
                          )
                        : { id: cached.folderId };
                    target = { driveId: cached.driveId, folderId: parent.id };
                } else {
                    const parent = await ensureFolderPath(
                        token,
                        state.dest.driveId,
                        it.destFolderRel || 'Import aus Nextcloud'
                    );
                    target = { driveId: state.dest.driveId, folderId: parent.id };
                }
                const blob = await readSourceBlob(it.sourceRelPath);
                await uploadBlob(token, target.driveId, target.folderId, it.fileName, blob);
            });
            done++;
            log('OK ' + it.sourceRelPath);
        } catch (e) {
            failed++;
            log('FEHLER ' + it.sourceRelPath + ': ' + (e.message || e));
        }
        setProgress(done + failed, items.length);
        setHint('Übertragung ' + (done + failed) + '/' + items.length, false);
    }

    state.running = false;
    setHint('Fertig: ' + done + ' OK, ' + failed + ' Fehler', failed === 0);
    toast('Fertig: ' + done + ' OK, ' + failed + ' Fehler.');
}

function wire() {
    loadSavedConn();
    if (!state.sourceMode || state.sourceMode === 'webdav') setSourceMode('helper', false);
    setHint('Start-Helfer.cmd starten, dann hier „Verbinden“.', false);
    renderBrowser();
    renderDest();
    renderChannelMaps();
    renderPlanSummary(null);
    setProgress(0, 0);

    document.querySelectorAll('.nc-mode-tab').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const mode = /** @type {'helper'|'zip'|'folder'} */ (btn.getAttribute('data-mode'));
            setSourceMode(mode, true);
        });
    });

    $('ncBtnHelperConnect') &&
        $('ncBtnHelperConnect').addEventListener('click', function () {
            connectHelper().catch(function (e) {
                const statusEl = $('ncHelperStatus');
                if (statusEl) statusEl.textContent = e.message || String(e);
                setHint(e.message || String(e), false);
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });

    $('ncZipFile') &&
        $('ncZipFile').addEventListener('change', function () {
            const f = $('ncZipFile').files && $('ncZipFile').files[0];
            const nameEl = $('ncZipName');
            if (nameEl) nameEl.textContent = f ? f.name + ' (' + formatBytes(f.size) + ')' : 'Noch keine Datei gewählt.';
            onZipChosen(f).catch(function (e) {
                toast(e.message || String(e));
                log('FEHLER: ' + (e.message || e));
                setHint(e.message || String(e), false);
            });
        });

    $('ncFolderInput') &&
        $('ncFolderInput').addEventListener('change', function () {
            try {
                const files = $('ncFolderInput').files;
                const nameEl = $('ncFolderName');
                if (nameEl) {
                    nameEl.textContent = files && files.length ? files.length + ' Dateien gewählt' : 'Noch kein Ordner gewählt.';
                }
                onFolderChosen(files);
            } catch (e) {
                toast(e.message || String(e));
            }
        });

    $('ncBtnSearchDest') &&
        $('ncBtnSearchDest').addEventListener('click', function () {
            runSearchDest().catch(function (e) {
                toast(e.message || String(e));
            });
        });
    $('ncDestQuery') &&
        $('ncDestQuery').addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                runSearchDest().catch(function (e) {
                    toast(e.message || String(e));
                });
            }
        });
    $('ncBtnClearSel') &&
        $('ncBtnClearSel').addEventListener('click', function () {
            state.selected.clear();
            renderSelection();
            renderBrowser();
            state.channelMaps = [];
            renderChannelMaps();
        });
    $('ncBtnSelectVisible') &&
        $('ncBtnSelectVisible').addEventListener('click', function () {
            state.items.forEach(function (it) {
                state.selected.add(it.relPath);
            });
            renderSelection();
            renderBrowser();
            refreshChannelMapsFromSelection().catch(function () {});
        });
    $('ncMapChannels') &&
        $('ncMapChannels').addEventListener('change', function () {
            saveConn();
            refreshChannelMapsFromSelection().catch(function () {});
        });
    $('ncBtnRefreshMaps') &&
        $('ncBtnRefreshMaps').addEventListener('click', function () {
            refreshChannelMapsFromSelection()
                .then(function () {
                    toast('Kanal-Vorschläge aktualisiert.');
                })
                .catch(function (e) {
                    toast(e.message || String(e));
                });
        });
    $('ncBtnDryRun') &&
        $('ncBtnDryRun').addEventListener('click', function () {
            runDryRun().catch(function (e) {
                toast(e.message || String(e));
                log('FEHLER: ' + (e.message || e));
            });
        });
    $('ncBtnMigrate') &&
        $('ncBtnMigrate').addEventListener('click', function () {
            runMigrate().catch(function (e) {
                state.running = false;
                toast(e.message || String(e));
                log('FEHLER: ' + (e.message || e));
            });
        });
    $('ncBtnAbort') &&
        $('ncBtnAbort').addEventListener('click', function () {
            state.abort = true;
            toast('Abbruch nach aktueller Datei …');
        });
    ['ncUrl', 'ncUser', 'ncPass', 'ncProxy', 'ncDestRoot'].forEach(function (id) {
        const el = $(id);
        if (el) el.addEventListener('change', saveConn);
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wire);
} else {
    wire();
}
