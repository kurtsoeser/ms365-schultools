/**
 * Datei-Migration Archiv-Team → neues Team.
 */
import {
    normalizeDriveItem,
    sortDriveItems,
    buildCopyBody,
    validateMigrationSelection
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

/** @type {{ sourceItems: ReturnType<typeof normalizeDriveItem>[], sourceDriveId: string, destDriveId: string }} */
const state = { sourceItems: [], sourceDriveId: '', destDriveId: '' };

async function getToken() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G || typeof G.getGraphToken !== 'function') throw new Error('Graph nicht geladen.');
    return G.getGraphToken(SCOPES);
}

async function graphJson(method, path, token, body) {
    const G = window.ms365GraphUnifiedGroups;
    return G.graphJson(method, path, token, body, 'v1.0');
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

function selectedIds() {
    return Array.from(document.querySelectorAll('#dmSourceBody input[type=checkbox]:checked')).map(function (el) {
        return el.value;
    });
}

function renderSource(items) {
    const body = $('dmSourceBody');
    if (!body) return;
    body.replaceChildren();
    if (!items.length) {
        body.innerHTML = '<tr><td colspan="3" class="muted">Keine Einträge (oder Bibliothek leer).</td></tr>';
        return;
    }
    items.forEach(function (it) {
        const tr = document.createElement('tr');
        tr.innerHTML =
            '<td><label><input type="checkbox" value="' +
            escapeHtml(it.id) +
            '"> ' +
            (it.isFolder ? '📁 ' : '') +
            escapeHtml(it.name) +
            '</label></td><td>' +
            (it.isFolder ? 'Ordner' : Math.round(it.size / 1024) + ' KB') +
            '</td><td class="muted">' +
            escapeHtml((it.lastModified || '').slice(0, 19)) +
            '</td>';
        body.appendChild(tr);
    });
}

async function loadSource() {
    const gid = String(($('dmSourceId') && $('dmSourceId').value) || '').trim();
    if (!gid) throw new Error('Quell-Gruppen-ID fehlt.');
    log('Lade Quell-Drive …');
    const token = await getToken();
    const drive = await resolveGroupDrive(token, gid);
    state.sourceDriveId = drive.id;
    state.sourceItems = await listChildren(token, drive.id, 'root');
    renderSource(state.sourceItems);
    log(state.sourceItems.length + ' Einträge in der Wurzel.');
    toast('Quelle geladen.');
}

async function resolveDest() {
    const gid = String(($('dmDestId') && $('dmDestId').value) || '').trim();
    if (!gid) throw new Error('Ziel-Gruppen-ID fehlt.');
    const token = await getToken();
    const drive = await resolveGroupDrive(token, gid);
    state.destDriveId = drive.id;
    log('Ziel-Drive: ' + drive.id);
    toast('Ziel erkannt.');
}

async function runCopy() {
    const sourceGroupId = String(($('dmSourceId') && $('dmSourceId').value) || '').trim();
    const destGroupId = String(($('dmDestId') && $('dmDestId').value) || '').trim();
    const itemIds = selectedIds();
    const v = validateMigrationSelection({ sourceGroupId, destGroupId, itemIds });
    if (!v.ok) throw new Error(v.issues.join(', '));
    if (!state.sourceDriveId) await loadSource();
    if (!state.destDriveId) await resolveDest();

    if (!window.confirm(itemIds.length + ' Element(e) kopieren?\n\n' + v.note)) return;

    const token = await getToken();
    const destFolder = String(($('dmDestFolder') && $('dmDestFolder').value) || 'root').trim() || 'root';
    let ok = 0;
    let fail = 0;
    for (let i = 0; i < itemIds.length; i++) {
        const id = itemIds[i];
        const item = state.sourceItems.find(function (x) {
            return x.id === id;
        });
        const label = item ? item.name : id;
        try {
            log('Kopiere „' + label + '“ …');
            const body = buildCopyBody({
                destDriveId: state.destDriveId,
                destFolderId: destFolder,
                newName: item && item.name ? item.name : undefined
            });
            // Graph copy returns 202 Accepted – graphJson may throw on empty body; use fetch
            const url =
                'https://graph.microsoft.com/v1.0/drives/' +
                encodeURIComponent(state.sourceDriveId) +
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

function boot() {
    const hint = $('dmHint');
    if (hint) {
        hint.textContent =
            'Kopiert Dateien/Ordner der Team-Dokumentbibliothek (Graph copy). Chat, Planner und Aufgaben bleiben unberührt.';
    }
    const b1 = $('dmBtnLoadSource');
    if (b1 && b1.dataset.bound !== '1') {
        b1.dataset.bound = '1';
        b1.addEventListener('click', function () {
            loadSource().catch(function (e) {
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });
    }
    const b2 = $('dmBtnResolveDest');
    if (b2 && b2.dataset.bound !== '1') {
        b2.dataset.bound = '1';
        b2.addEventListener('click', function () {
            resolveDest().catch(function (e) {
                toast(e.message || String(e));
            });
        });
    }
    const b3 = $('dmBtnCopy');
    if (b3 && b3.dataset.bound !== '1') {
        b3.dataset.bound = '1';
        b3.addEventListener('click', function () {
            runCopy().catch(function (e) {
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });
    }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
