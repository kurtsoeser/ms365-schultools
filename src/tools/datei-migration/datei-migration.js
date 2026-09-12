/**
 * Datei-Migration Archiv-Team → neues Team (Suche + Ordnerbrowser).
 */
import {
    normalizeDriveItem,
    sortDriveItems,
    buildCopyBody,
    validateMigrationSelection,
    pushBreadcrumb,
    sliceBreadcrumb
} from './datei-migration-logic.js';

function $(id) {
    return document.getElementById(id);
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function log(msg) {
    const el = $('dmLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(msg || '');
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.Read.All',
    'https://graph.microsoft.com/Files.ReadWrite.All',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

/**
 * @typedef {{ id: string, name: string }} Crumb
 * @typedef {{ id: string, name: string, pathLabel: string }} SelectedItem
 */

const state = {
    source: {
        groupId: '',
        label: '',
        driveId: '',
        items: /** @type {ReturnType<typeof normalizeDriveItem>[]} */ ([]),
        crumbs: /** @type {Crumb[]} */ ([{ id: 'root', name: 'Stamm' }])
    },
    dest: {
        groupId: '',
        label: '',
        driveId: '',
        items: /** @type {ReturnType<typeof normalizeDriveItem>[]} */ ([]),
        crumbs: /** @type {Crumb[]} */ ([{ id: 'root', name: 'Stamm' }])
    },
    /** @type {Map<string, SelectedItem>} */
    selected: new Map()
};

async function getToken() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G || typeof G.getGraphToken !== 'function') throw new Error('Graph nicht geladen.');
    return G.getGraphToken(SCOPES);
}

async function graphJson(method, path, token, body) {
    const G = window.ms365GraphUnifiedGroups;
    return G.graphJson(method, path, token, body, 'v1.0');
}

function guidLike(s) {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(s || '').trim());
}

async function resolveGroupDrive(token, groupId) {
    const data = await graphJson('GET', '/groups/' + encodeURIComponent(groupId) + '/drive', token);
    if (!data || !data.id) throw new Error('Keine Dokumentbibliothek für diese Gruppe (ist es ein Team?).');
    return data;
}

async function listChildren(token, driveId, itemId) {
    const path =
        '/drives/' +
        encodeURIComponent(driveId) +
        '/items/' +
        encodeURIComponent(itemId || 'root') +
        '/children?$select=id,name,size,folder,file,webUrl,lastModifiedDateTime&$top=200';
    const data = await graphJson('GET', path, token);
    return sortDriveItems(((data && data.value) || []).map(normalizeDriveItem));
}

/**
 * @param {string} query
 * @returns {Promise<Array<{ id: string, displayName: string, mail: string, mailNickname: string }>>}
 */
async function searchTeams(query) {
    const q = String(query || '').trim();
    if (!q) throw new Error('Suchbegriff fehlt.');
    const token = await getToken();
    const G = window.ms365GraphUnifiedGroups;

    if (guidLike(q)) {
        try {
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
        } catch (e) {
            throw new Error('GUID nicht gefunden: ' + (e.message || e));
        }
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

function currentFolderId(side) {
    const crumbs = state[side].crumbs;
    return crumbs.length ? crumbs[crumbs.length - 1].id : 'root';
}

function pathLabelFor(side, itemName) {
    const names = state[side].crumbs.map(function (c) {
        return c.name;
    });
    names.push(itemName);
    return names.join(' / ');
}

function renderHits(side, hits) {
    const ul = $(side === 'source' ? 'dmSourceHits' : 'dmDestHits');
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
            pickTeam(side, g).catch(function (e) {
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });
        li.appendChild(btn);
        ul.appendChild(li);
    });
}

function renderPicked(side) {
    const el = $(side === 'source' ? 'dmSourcePicked' : 'dmDestPicked');
    const idEl = $(side === 'source' ? 'dmSourceId' : 'dmDestId');
    const s = state[side];
    if (idEl) idEl.value = s.groupId || '';
    if (!el) return;
    if (!s.groupId) {
        el.className = 'dm-picked muted';
        el.textContent = side === 'source' ? 'Noch kein Quell-Team gewählt.' : 'Noch kein Ziel-Team gewählt.';
        return;
    }
    el.className = 'dm-picked';
    el.innerHTML =
        '<strong>' +
        escapeHtml(s.label) +
        '</strong><br><span class="muted" style="font-size:0.85em;">' +
        escapeHtml(s.groupId) +
        '</span>';
}

function renderCrumb(side) {
    const el = $(side === 'source' ? 'dmSourceCrumb' : 'dmDestCrumb');
    if (!el) return;
    el.replaceChildren();
    state[side].crumbs.forEach(function (c, idx) {
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
            navigateToCrumb(side, idx).catch(function (e) {
                toast(e.message || String(e));
            });
        });
        el.appendChild(btn);
    });
    if (side === 'dest') updateDestFolderLabel();
}

function updateDestFolderLabel() {
    const el = $('dmDestFolderLabel');
    if (!el) return;
    const path = state.dest.crumbs
        .map(function (c) {
            return c.name;
        })
        .join(' / ');
    el.innerHTML = '<strong>Kopierziel:</strong> ' + escapeHtml(path || 'Stamm');
}

function renderSelection() {
    const count = $('dmSelectionCount');
    const list = $('dmSelectionList');
    if (count) count.textContent = String(state.selected.size);
    if (!list) return;
    list.replaceChildren();
    state.selected.forEach(function (it) {
        const li = document.createElement('li');
        li.textContent = it.pathLabel;
        list.appendChild(li);
    });
}

function renderBrowser(side) {
    const body = $(side === 'source' ? 'dmSourceBody' : 'dmDestBody');
    if (!body) return;
    body.replaceChildren();
    const items = state[side].items;
    const colSpan = side === 'source' ? 3 : 2;
    if (!state[side].driveId) {
        body.innerHTML =
            '<tr><td colspan="' +
            colSpan +
            '" class="muted">Team wählen und Bibliothek laden.</td></tr>';
        return;
    }
    if (!items.length) {
        body.innerHTML = '<tr><td colspan="' + colSpan + '" class="muted">Ordner ist leer.</td></tr>';
        return;
    }
    items.forEach(function (it) {
        const tr = document.createElement('tr');
        if (side === 'source') {
            const tdCheck = document.createElement('td');
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = it.id;
            cb.checked = state.selected.has(it.id);
            cb.addEventListener('change', function () {
                if (cb.checked) {
                    state.selected.set(it.id, {
                        id: it.id,
                        name: it.name,
                        pathLabel: pathLabelFor('source', it.name)
                    });
                } else {
                    state.selected.delete(it.id);
                }
                renderSelection();
            });
            tdCheck.appendChild(cb);
            tr.appendChild(tdCheck);
        }
        const tdName = document.createElement('td');
        if (it.isFolder) {
            const open = document.createElement('button');
            open.type = 'button';
            open.className = 'dm-open';
            open.textContent = '📁 ' + it.name;
            open.addEventListener('click', function () {
                openFolder(side, it).catch(function (e) {
                    toast(e.message || String(e));
                });
            });
            tdName.appendChild(open);
        } else {
            tdName.textContent = it.name;
        }
        tr.appendChild(tdName);
        const tdArt = document.createElement('td');
        tdArt.textContent = it.isFolder
            ? 'Ordner' + (it.childCount != null ? ' (' + it.childCount + ')' : '')
            : Math.round(it.size / 1024) + ' KB';
        tr.appendChild(tdArt);
        body.appendChild(tr);
    });
}

async function pickTeam(side, group) {
    const token = await getToken();
    log((side === 'source' ? 'Quelle' : 'Ziel') + ': ' + (group.displayName || group.id));
    const drive = await resolveGroupDrive(token, group.id);
    state[side].groupId = group.id;
    state[side].label = group.displayName || group.mailNickname || group.id;
    state[side].driveId = drive.id;
    state[side].crumbs = [{ id: 'root', name: 'Stamm' }];
    state[side].items = await listChildren(token, drive.id, 'root');
    if (side === 'source') {
        state.selected.clear();
        renderSelection();
    }
    const hits = $(side === 'source' ? 'dmSourceHits' : 'dmDestHits');
    if (hits) {
        hits.replaceChildren();
        hits.hidden = true;
    }
    renderPicked(side);
    renderCrumb(side);
    renderBrowser(side);
    toast((side === 'source' ? 'Quelle' : 'Ziel') + ' geladen.');
}

async function openFolder(side, item) {
    if (!item || !item.isFolder) return;
    const token = await getToken();
    state[side].crumbs = pushBreadcrumb(state[side].crumbs, { id: item.id, name: item.name });
    state[side].items = await listChildren(token, state[side].driveId, item.id);
    renderCrumb(side);
    renderBrowser(side);
}

async function navigateToCrumb(side, index) {
    const token = await getToken();
    state[side].crumbs = sliceBreadcrumb(state[side].crumbs, index);
    const folderId = currentFolderId(side);
    state[side].items = await listChildren(token, state[side].driveId, folderId);
    renderCrumb(side);
    renderBrowser(side);
}

async function runSearch(side) {
    const qEl = $(side === 'source' ? 'dmSourceQuery' : 'dmDestQuery');
    const q = String((qEl && qEl.value) || '').trim();
    log('Suche ' + (side === 'source' ? 'Quelle' : 'Ziel') + ': ' + q);
    const hits = await searchTeams(q);
    if (!hits.length) {
        renderHits(side, []);
        toast('Keine Treffer.');
        return;
    }
    renderHits(side, hits);
    if (hits.length === 1) {
        await pickTeam(side, hits[0]);
    } else {
        toast(hits.length + ' Treffer – bitte wählen.');
    }
}

async function runCopy() {
    const itemIds = Array.from(state.selected.keys());
    const v = validateMigrationSelection({
        sourceGroupId: state.source.groupId,
        destGroupId: state.dest.groupId,
        itemIds: itemIds
    });
    if (!v.ok) throw new Error(v.issues.join(', '));
    if (!state.source.driveId) throw new Error('Quell-Bibliothek fehlt.');
    if (!state.dest.driveId) throw new Error('Ziel-Bibliothek fehlt.');

    const destFolder = currentFolderId('dest');
    const destPath = state.dest.crumbs
        .map(function (c) {
            return c.name;
        })
        .join(' / ');
    if (
        !window.confirm(
            itemIds.length +
                ' Element(e) nach „' +
                destPath +
                '“ kopieren?\n\n' +
                v.note
        )
    ) {
        return;
    }

    const token = await getToken();
    let ok = 0;
    let fail = 0;
    for (let i = 0; i < itemIds.length; i++) {
        const id = itemIds[i];
        const sel = state.selected.get(id);
        const label = sel ? sel.name : id;
        try {
            log('Kopiere „' + label + '“ → ' + destPath + ' …');
            const body = buildCopyBody({
                destDriveId: state.dest.driveId,
                destFolderId: destFolder,
                newName: sel && sel.name ? sel.name : undefined
            });
            const url =
                'https://graph.microsoft.com/v1.0/drives/' +
                encodeURIComponent(state.source.driveId) +
                '/items/' +
                encodeURIComponent(id) +
                '/copy';
            const res = await fetch(url, {
                method: 'POST',
                headers: {
                    Authorization: 'Bearer ' + token,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(body)
            });
            if (res.status === 202 || res.ok) {
                ok++;
                log('  → gestartet (asynchron).');
            } else {
                const t = await res.text();
                fail++;
                log('  → FEHLER ' + res.status + ' ' + t.slice(0, 200));
            }
            await new Promise(function (r) {
                setTimeout(r, 200);
            });
        } catch (e) {
            fail++;
            log('  → FEHLER ' + (e.message || e));
        }
    }
    toast('Kopieraufträge: ' + ok + ' ok, ' + fail + ' Fehler');
    log(v.note);
}

function bindSearchEnter(inputId, side) {
    const el = $(inputId);
    if (!el || el.dataset.boundEnter === '1') return;
    el.dataset.boundEnter = '1';
    el.addEventListener('keydown', function (ev) {
        if (ev.key === 'Enter') {
            ev.preventDefault();
            runSearch(side).catch(function (e) {
                toast(e.message || String(e));
            });
        }
    });
}

function boot() {
    const hint = $('dmHint');
    if (hint) {
        hint.textContent =
            'Teams suchen, Ordner öffnen, Quelle anhaken und Zielordner wählen – dann kopieren. Chat/Aufgaben bleiben unberührt.';
    }
    renderPicked('source');
    renderPicked('dest');
    renderCrumb('source');
    renderCrumb('dest');
    renderSelection();

    const map = [
        ['dmBtnSearchSource', function () {
            return runSearch('source');
        }],
        ['dmBtnSearchDest', function () {
            return runSearch('dest');
        }],
        ['dmBtnCopy', runCopy],
        [
            'dmBtnClearSel',
            function () {
                state.selected.clear();
                renderSelection();
                renderBrowser('source');
                return Promise.resolve();
            }
        ]
    ];
    map.forEach(function (pair) {
        const btn = $(pair[0]);
        if (!btn || btn.dataset.bound === '1') return;
        btn.dataset.bound = '1';
        btn.addEventListener('click', function () {
            Promise.resolve(pair[1]()).catch(function (e) {
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });
    });
    bindSearchEnter('dmSourceQuery', 'source');
    bindSearchEnter('dmDestQuery', 'dest');
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
