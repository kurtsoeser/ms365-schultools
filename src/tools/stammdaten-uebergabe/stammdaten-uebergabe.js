/**
 * Stammdaten → SharePoint Document Library (Upload / Liste / Download).
 */
import {
    DEFAULT_FOLDER,
    CURRENT_FILE,
    buildDriveRelativePath,
    encodeDriveRootPath,
    designHintDe,
    describeRemoteBackup
} from '../../shared/stammdaten-sharepoint-sync-logic.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
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
    const el = $('suLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(msg || '');
}

function clearLog() {
    const el = $('suLog');
    if (el) el.textContent = '';
}

function getG() {
    const G = window.ms365SpoGraph;
    if (!G) throw new Error('SharePoint-Graph-Helfer nicht geladen.');
    return G;
}

function getSiteUrl() {
    const input = $('suSiteUrl');
    let url = input && input.value ? String(input.value).trim() : '';
    if (!url) {
        try {
            const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
            url = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
            if (url && input) input.value = url;
        } catch {
            /* ignore */
        }
    }
    return url;
}

function getFolder() {
    const el = $('suFolder');
    return String((el && el.value) || DEFAULT_FOLDER).trim() || DEFAULT_FOLDER;
}

function rememberSite(url) {
    if (!url || !window.ms365AppDataV2 || typeof window.ms365AppDataV2.patchSetup !== 'function') return;
    try {
        const cur = window.ms365AppDataV2.getSetup() || {};
        if (String(cur.intranetSiteUrl || '').trim() === url) return;
        window.ms365AppDataV2.patchSetup({ intranetSiteUrl: url });
    } catch {
        /* ignore */
    }
}

function saveLocalMeta(meta) {
    try {
        localStorage.setItem('ms365-stammdaten-spo-sync-v1', JSON.stringify(meta || {}));
    } catch {
        /* ignore */
    }
}

function loadLocalMeta() {
    try {
        return JSON.parse(localStorage.getItem('ms365-stammdaten-spo-sync-v1') || '{}') || {};
    } catch {
        return {};
    }
}

function refreshMetaUi() {
    const el = $('suLastMeta');
    if (!el) return;
    const m = loadLocalMeta();
    if (!m || !m.at) {
        el.textContent = 'Noch kein Upload aus diesem Browser.';
        return;
    }
    el.textContent =
        'Zuletzt: ' +
        String(m.at).replace('T', ' ').replace(/\.\d+Z$/, '') +
        (m.fileName ? ' · ' + m.fileName : '') +
        (m.webUrl ? ' · Datei vorhanden' : '');
    const link = $('suFileLink');
    if (link) {
        if (m.webUrl) {
            link.hidden = false;
            link.href = m.webUrl;
        } else {
            link.hidden = true;
            link.removeAttribute('href');
        }
    }
}

async function ensureToken() {
    return getG().getGraphToken(SCOPES);
}

async function resolveSite(token, webUrl) {
    const site = await getG().resolveSiteFromWebUrl(token, webUrl);
    if (!site || !site.id) throw new Error('Site konnte nicht aufgelöst werden.');
    return site;
}

/**
 * @param {string} siteId
 * @param {string} relativePath Ordner/Datei unter root
 * @param {string} jsonText
 */
async function putJsonFile(siteId, relativePath, jsonText, token) {
    const G = getG();
    const enc = encodeDriveRootPath(relativePath);
    const path = G.graphPathSite(siteId) + '/drive/' + enc + '/content';
    const url = G.graphBase('v1.0') + path;
    const res = await fetch(url, {
        method: 'PUT',
        headers: {
            Authorization: 'Bearer ' + token,
            'Content-Type': 'application/json; charset=utf-8'
        },
        body: jsonText
    });
    const text = await res.text();
    let data = null;
    try {
        data = text ? JSON.parse(text) : {};
    } catch {
        data = { raw: text };
    }
    if (!res.ok) {
        const msg =
            data && data.error && data.error.message
                ? data.error.message
                : text || String(res.status);
        throw new Error('Upload fehlgeschlagen: ' + msg);
    }
    return data;
}

async function listFolder(siteId, folder, token) {
    const G = getG();
    const rel = buildDriveRelativePath(folder, '');
    const enc = encodeDriveRootPath(rel.replace(/\/$/, '') || DEFAULT_FOLDER);
    const path = G.graphPathSite(siteId) + '/drive/' + enc + '/children?$select=id,name,size,lastModifiedDateTime,webUrl,file&$orderby=lastModifiedDateTime desc&$top=50';
    try {
        return await G.graphJson('GET', path, token, undefined, 'v1.0');
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e);
        if (/itemNotFound|404|not found/i.test(msg)) {
            return { value: [] };
        }
        throw e;
    }
}

async function downloadItemContent(siteId, itemId, token) {
    const G = getG();
    const path = G.graphPathSite(siteId) + '/drive/items/' + encodeURIComponent(itemId) + '/content';
    const url = G.graphBase('v1.0') + path;
    const res = await fetch(url, {
        method: 'GET',
        headers: { Authorization: 'Bearer ' + token }
    });
    const text = await res.text();
    if (!res.ok) throw new Error('Download fehlgeschlagen: HTTP ' + res.status);
    let obj;
    try {
        obj = JSON.parse(text);
    } catch {
        throw new Error('Datei ist kein gültiges JSON.');
    }
    return obj;
}

function buildPayloadJson() {
    const bb = window.ms365BrowserBackup;
    if (!bb || typeof bb.buildBackup !== 'function') {
        throw new Error('Browser-Backup-Modul fehlt.');
    }
    const payload = bb.buildBackup();
    return { payload: payload, text: JSON.stringify(payload, null, 2), fileName: CURRENT_FILE };
}

async function runUpload() {
    clearLog();
    const webUrl = getSiteUrl();
    if (!webUrl) throw new Error('SharePoint-Website fehlt (Intranet-URL).');
    const folder = getFolder();
    const keepDated = !!($('suKeepDated') && $('suKeepDated').checked);

    log('Baue Browser-Backup …');
    const built = buildPayloadJson();
    const token = await ensureToken();
    log('Löse Site auf …');
    const site = await resolveSite(token, webUrl);
    rememberSite(webUrl);
    log('Site: ' + (site.displayName || site.id));

    const currentPath = buildDriveRelativePath(folder, CURRENT_FILE);
    log('Upload → ' + currentPath);
    const item = await putJsonFile(site.id, currentPath, built.text, token);
    let datedItem = null;
    if (keepDated && window.ms365BrowserBackup && typeof window.ms365BrowserBackup.backupFilename === 'function') {
        const datedName = window.ms365BrowserBackup.backupFilename(new Date());
        const datedPath = buildDriveRelativePath(folder, datedName);
        log('Zusätzlich datiert → ' + datedPath);
        datedItem = await putJsonFile(site.id, datedPath, built.text, token);
    }

    const web = (item && item.webUrl) || (datedItem && datedItem.webUrl) || '';
    saveLocalMeta({
        at: new Date().toISOString(),
        fileName: CURRENT_FILE,
        folder: folder,
        webUrl: web,
        siteUrl: webUrl,
        summary: describeRemoteBackup(built.payload)
    });
    try {
        localStorage.setItem('ms365-last-backup-export-at', new Date().toISOString());
    } catch {
        /* ignore */
    }
    refreshMetaUi();
    log('Fertig. ' + designHintDe());
    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
        window.ms365ActionLog.append({
            tool: 'stammdaten-uebergabe',
            action: 'spo-upload',
            target: webUrl,
            summary: CURRENT_FILE + ' → ' + folder
        });
    }
    toast('Stammdaten auf SharePoint geschrieben.');
    return item;
}

async function runList() {
    clearLog();
    const webUrl = getSiteUrl();
    if (!webUrl) throw new Error('SharePoint-Website fehlt.');
    const folder = getFolder();
    const token = await ensureToken();
    const site = await resolveSite(token, webUrl);
    rememberSite(webUrl);
    log('Liste Ordner „' + folder + '“ …');
    const data = await listFolder(site.id, folder, token);
    const items = (data && data.value) || [];
    const body = $('suRemoteBody');
    if (body) {
        body.replaceChildren();
        if (!items.length) {
            body.innerHTML = '<tr><td colspan="4" class="muted">Ordner leer oder noch nicht angelegt (wird beim Upload erzeugt).</td></tr>';
        } else {
            items.forEach(function (it) {
                if (!it || !it.file) return;
                const tr = document.createElement('tr');
                const when = it.lastModifiedDateTime
                    ? String(it.lastModifiedDateTime).replace('T', ' ').replace(/\.\d+Z$/, ' UTC')
                    : '';
                const size = it.size != null ? Math.round(Number(it.size) / 1024) + ' KB' : '';
                tr.innerHTML =
                    '<td>' +
                    escapeHtml(it.name || '') +
                    '</td><td>' +
                    escapeHtml(when) +
                    '</td><td>' +
                    escapeHtml(size) +
                    '</td><td></td>';
                const td = tr.lastElementChild;
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn btn-sm';
                btn.innerHTML = '<i class="bi bi-download"></i>Laden';
                btn.addEventListener('click', function () {
                    runDownload(it.id, it.name).catch(function (e) {
                        toast(e.message || String(e));
                        log('FEHLER: ' + (e.message || e));
                    });
                });
                td.appendChild(btn);
                if (it.webUrl) {
                    const a = document.createElement('a');
                    a.className = 'btn btn-sm';
                    a.href = it.webUrl;
                    a.target = '_blank';
                    a.rel = 'noopener';
                    a.textContent = 'Öffnen';
                    td.appendChild(document.createTextNode(' '));
                    td.appendChild(a);
                }
                body.appendChild(tr);
            });
        }
    }
    log(items.filter(function (i) { return i && i.file; }).length + ' Datei(en).');
    toast('Ordner gelesen.');
}

async function runDownload(itemId, name) {
    const webUrl = getSiteUrl();
    if (!webUrl) throw new Error('SharePoint-Website fehlt.');
    const token = await ensureToken();
    const site = await resolveSite(token, webUrl);
    log('Lade ' + (name || itemId) + ' …');
    const obj = await downloadItemContent(site.id, itemId, token);
    const bb = window.ms365BrowserBackup;
    if (!bb || typeof bb.isBackupPayload !== 'function') throw new Error('Backup-Modul fehlt.');
    if (!bb.isBackupPayload(obj) && !bb.isLegacyAppDataPayload(obj)) {
        throw new Error('Datei ist kein erkanntes MS365-Browser-Backup.');
    }
    const summary =
        (obj.schoolName || obj.domain || name || 'Backup') +
        (obj.exportedAt ? ' · ' + String(obj.exportedAt).replace('T', ' ').slice(0, 19) : '');
    if (
        !window.confirm(
            'Backup von SharePoint übernehmen und lokale Daten ersetzen?\n\n' +
                summary +
                '\n\nMicrosoft-Anmeldung bleibt erhalten; PIN ggf. neu.'
        )
    ) {
        log('Abgebrochen.');
        return;
    }
    bb.importPayload(obj);
    toast('Backup von SharePoint übernommen.');
    log('Import abgeschlossen. Seite neu laden empfohlen.');
    if (window.confirm('Seite jetzt neu laden?')) window.location.reload();
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function fillHint() {
    const el = $('suDesignHint');
    if (el) el.textContent = designHintDe();
}

function boot() {
    fillHint();
    try {
        const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
        const saved = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
        if (saved && $('suSiteUrl') && !$('suSiteUrl').value) $('suSiteUrl').value = saved;
    } catch {
        /* ignore */
    }
    if ($('suFolder') && !$('suFolder').value) $('suFolder').value = DEFAULT_FOLDER;
    refreshMetaUi();

    const up = $('suBtnUpload');
    if (up && up.dataset.bound !== '1') {
        up.dataset.bound = '1';
        up.addEventListener('click', function () {
            runUpload().catch(function (e) {
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });
    }
    const list = $('suBtnList');
    if (list && list.dataset.bound !== '1') {
        list.dataset.bound = '1';
        list.addEventListener('click', function () {
            runList().catch(function (e) {
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });
    }
    const loadCur = $('suBtnLoadCurrent');
    if (loadCur && loadCur.dataset.bound !== '1') {
        loadCur.dataset.bound = '1';
        loadCur.addEventListener('click', function () {
            (async function () {
                clearLog();
                const webUrl = getSiteUrl();
                if (!webUrl) throw new Error('SharePoint-Website fehlt.');
                const folder = getFolder();
                const token = await ensureToken();
                const site = await resolveSite(token, webUrl);
                const data = await listFolder(site.id, folder, token);
                const items = (data && data.value) || [];
                const cur = items.find(function (i) {
                    return i && i.file && String(i.name || '') === CURRENT_FILE;
                });
                if (!cur) throw new Error('Datei „' + CURRENT_FILE + '“ nicht gefunden – zuerst hochladen.');
                await runDownload(cur.id, cur.name);
            })().catch(function (e) {
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}
