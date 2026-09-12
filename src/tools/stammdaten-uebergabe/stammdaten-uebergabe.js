/**
 * Stammdaten → IT-Dokumentbibliothek (Upload / Rechte / Liste / Download).
 */
import {
    DEFAULT_FOLDER,
    CURRENT_FILE,
    IT_LIBRARY_TITLE,
    buildDriveRelativePath,
    encodeDriveRootPath,
    designHintDe,
    describeRemoteBackup,
    isBroadSiteAudience,
    entraGroupLogonName,
    buildItLibraryPlan,
    SPO_ROLE
} from '../../shared/stammdaten-sharepoint-sync-logic.js';

const SCOPES_GRAPH = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/Group.Read.All'
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

function getLibraryTitle() {
    const el = $('suLibraryTitle');
    return String((el && el.value) || IT_LIBRARY_TITLE).trim() || IT_LIBRARY_TITLE;
}

function loadItMeta() {
    try {
        const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
        if (setup && setup.stammdatenItLibrary && typeof setup.stammdatenItLibrary === 'object') {
            return setup.stammdatenItLibrary;
        }
    } catch {
        /* ignore */
    }
    try {
        return JSON.parse(localStorage.getItem('ms365-stammdaten-it-library-v1') || '{}') || {};
    } catch {
        return {};
    }
}

function saveItMeta(meta) {
    const m = meta || {};
    try {
        localStorage.setItem('ms365-stammdaten-it-library-v1', JSON.stringify(m));
    } catch {
        /* ignore */
    }
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup({ stammdatenItLibrary: m });
        }
    } catch {
        /* ignore */
    }
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
    const it = loadItMeta();
    const m = loadLocalMeta();
    if (el) {
        const bits = [];
        if (it && it.driveId) {
            bits.push('IT-Bibliothek „' + (it.listTitle || IT_LIBRARY_TITLE) + '“');
            if (it.securedAt) bits.push('Rechte gesetzt');
            if (it.itGroupMail || it.itGroupId) bits.push('Gruppe: ' + (it.itGroupMail || it.itGroupId));
        } else {
            bits.push('IT-Bibliothek noch nicht eingerichtet');
        }
        if (m && m.at) {
            bits.push(
                'Upload: ' +
                    String(m.at).replace('T', ' ').replace(/\.\d+Z$/, '') +
                    (m.fileName ? ' · ' + m.fileName : '')
            );
        }
        el.textContent = bits.join(' · ');
    }
    const link = $('suFileLink');
    if (link) {
        const href = (m && m.webUrl) || (it && it.webUrl) || '';
        if (href) {
            link.hidden = false;
            link.href = href;
        } else {
            link.hidden = true;
            link.removeAttribute('href');
        }
    }
    const status = $('suItStatus');
    if (status) {
        status.textContent =
            it && it.driveId
                ? 'Bereit: Drive ' + String(it.driveId).slice(0, 8) + '…'
                : 'Noch keine IT-Bibliothek – bitte unten einrichten.';
    }
}

async function ensureGraphToken() {
    return getG().getGraphToken(SCOPES_GRAPH);
}

async function resolveSite(token, webUrl) {
    const site = await getG().resolveSiteFromWebUrl(token, webUrl);
    if (!site || !site.id) throw new Error('Site konnte nicht aufgelöst werden.');
    return site;
}

async function findListByTitle(token, siteId, listTitle) {
    const G = getG();
    const title = String(listTitle || '').trim();
    const path =
        G.graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await G.graphJson('GET', path, token, undefined, 'v1.0');
    const list = (data && data.value) || [];
    return list[0] || null;
}

async function createDocumentLibrary(token, siteId, title, description) {
    const G = getG();
    return G.graphJson(
        'POST',
        G.graphPathSite(siteId) + '/lists',
        token,
        {
            displayName: title,
            description: description || '',
            list: { template: 'documentLibrary' }
        },
        'v1.0'
    );
}

async function getListDrive(token, siteId, listId) {
    const G = getG();
    return G.graphJson(
        'GET',
        G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/drive',
        token,
        undefined,
        'v1.0'
    );
}

async function resolveGroupId(token, mailOrId) {
    const raw = String(mailOrId || '').trim();
    if (!raw) return '';
    if (/^[0-9a-f-]{36}$/i.test(raw)) return raw;
    const G = getG();
    const esc = raw.replace(/'/g, "''");
    const filter = encodeURIComponent(
        "mail eq '" + esc + "' or mailNickname eq '" + esc + "' or displayName eq '" + esc + "'"
    );
    const data = await G.graphJson(
        'GET',
        '/groups?$filter=' + filter + '&$select=id,displayName,mail,mailNickname&$top=5',
        token,
        undefined,
        'v1.0'
    );
    const g = ((data && data.value) || [])[0];
    return g && g.id ? String(g.id) : '';
}

async function putJsonOnDrive(driveId, relativePath, jsonText, token) {
    const G = getG();
    const enc = encodeDriveRootPath(relativePath);
    const url = G.graphBase('v1.0') + '/drives/' + encodeURIComponent(driveId) + '/' + enc + '/content';
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
            data && data.error && data.error.message ? data.error.message : text || String(res.status);
        throw new Error('Upload fehlgeschlagen: ' + msg);
    }
    return data;
}

async function listDriveFolder(driveId, folder, token) {
    const G = getG();
    const rel = buildDriveRelativePath(folder, '');
    const enc = encodeDriveRootPath(rel.replace(/\/$/, '') || DEFAULT_FOLDER);
    const path =
        '/drives/' +
        encodeURIComponent(driveId) +
        '/' +
        enc +
        '/children?$select=id,name,size,lastModifiedDateTime,webUrl,file&$orderby=lastModifiedDateTime desc&$top=50';
    try {
        return await G.graphJson('GET', path, token, undefined, 'v1.0');
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e);
        if (/itemNotFound|404|not found/i.test(msg)) return { value: [] };
        throw e;
    }
}

async function downloadDriveItem(driveId, itemId, token) {
    const G = getG();
    const url = G.graphBase('v1.0') + '/drives/' + encodeURIComponent(driveId) + '/items/' + encodeURIComponent(itemId) + '/content';
    const res = await fetch(url, { method: 'GET', headers: { Authorization: 'Bearer ' + token } });
    const text = await res.text();
    if (!res.ok) throw new Error('Download fehlgeschlagen: HTTP ' + res.status);
    return JSON.parse(text);
}

function requireDriveId() {
    const it = loadItMeta();
    if (!it || !it.driveId) {
        throw new Error('Bitte zuerst „IT-Bibliothek einrichten“ ausführen.');
    }
    return it;
}

async function secureLibraryWithSpo(siteWebUrl, listTitle, groupObjectId) {
    const G = getG();
    let host = '';
    try {
        host = new URL(siteWebUrl).hostname;
    } catch {
        throw new Error('Ungültige Site-URL für SharePoint-Token.');
    }
    log('Hole SharePoint-Token (Sites.FullControl.All) …');
    const spoScope = 'https://' + host + '/Sites.FullControl.All';
    let spoToken;
    try {
        spoToken = await G.getGraphToken([spoScope]);
    } catch (e) {
        throw new Error(
            'SharePoint-Token fehlgeschlagen (App-Zustimmung für SharePoint / Sites.FullControl.All?). ' +
                (e && e.message ? e.message : e)
        );
    }
    const digest = await G.getSpoRequestDigest(siteWebUrl, spoToken);
    log('Breche Vererbung (kopiert zuerst bestehende Rollen) …');
    await G.spoBreakListInheritance(siteWebUrl, spoToken, digest, listTitle, true);
    const assignments = await G.spoListRoleAssignments(siteWebUrl, spoToken, digest, listTitle);
    let removed = 0;
    for (let i = 0; i < assignments.length; i++) {
        const a = assignments[i];
        const member = a.Member || a.member || {};
        if (!isBroadSiteAudience(member)) continue;
        const pid = member.Id != null ? member.Id : a.PrincipalId;
        try {
            await G.spoRemoveRoleAssignment(siteWebUrl, spoToken, digest, listTitle, pid);
            removed++;
            log('Entfernt: ' + (member.Title || pid));
        } catch (e) {
            log('Hinweis Entfernen ' + (member.Title || pid) + ': ' + (e.message || e));
        }
    }
    log('Breite Rollen entfernt: ' + removed);
    const logon = entraGroupLogonName(groupObjectId);
    log('EnsureUser IT-Gruppe …');
    const principal = await G.spoEnsureUser(siteWebUrl, spoToken, digest, logon);
    await G.spoAddRoleAssignment(siteWebUrl, spoToken, digest, listTitle, principal.id, SPO_ROLE.contribute);
    log('IT-Gruppe berechtigt (Contribute): ' + (principal.title || groupObjectId));
    return { removed: removed, principalId: principal.id };
}

async function runSetupItLibrary() {
    clearLog();
    const webUrl = getSiteUrl();
    if (!webUrl) throw new Error('SharePoint-Website fehlt.');
    const listTitle = getLibraryTitle();
    let groupRaw = String(($('suItGroup') && $('suItGroup').value) || '').trim();
    if (!groupRaw) {
        try {
            const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
            const matched = setup && setup.matched ? setup.matched : {};
            if (matched.verwaltungGroupId) groupRaw = String(matched.verwaltungGroupId);
        } catch {
            /* ignore */
        }
    }
    const plan = buildItLibraryPlan({
        listTitle: listTitle,
        itGroupId: /^[0-9a-f-]{36}$/i.test(groupRaw) ? groupRaw : '',
        itGroupMail: /^[0-9a-f-]{36}$/i.test(groupRaw) ? '' : groupRaw
    });
    if (!plan.ok) throw new Error(plan.issues.join(', '));

    if (
        !window.confirm(
            'IT-Bibliothek „' +
                listTitle +
                '“ anlegen/prüfen und Rechte setzen?\n\n' +
                '• Vererbung wird gebrochen\n' +
                '• Site-Besucher und -Mitglieder verlieren Zugriff\n' +
                '• Site-Besitzer bleiben\n' +
                '• Gewählte Gruppe erhält Mitwirken (Contribute)'
        )
    ) {
        return;
    }

    const token = await ensureGraphToken();
    log('Löse Site auf …');
    const site = await resolveSite(token, webUrl);
    rememberSite(webUrl);

    let list = await findListByTitle(token, site.id, listTitle);
    if (!list) {
        log('Lege Dokumentbibliothek an …');
        list = await createDocumentLibrary(token, site.id, listTitle, plan.description);
        await getG().sleep(1500);
    } else {
        log('Bibliothek existiert bereits.');
    }
    const listId = list.id || list.Id;
    if (!listId) throw new Error('Listen-ID fehlt.');
    const drive = await getListDrive(token, site.id, listId);
    if (!drive || !drive.id) throw new Error('Drive der Bibliothek fehlt.');

    let groupId = plan.itGroupId;
    if (!groupId) {
        log('Suche Gruppe …');
        groupId = await resolveGroupId(token, plan.itGroupMail);
    }
    if (!groupId) throw new Error('Gruppe nicht gefunden: ' + (plan.itGroupMail || plan.itGroupId));

    try {
        await secureLibraryWithSpo(webUrl, listTitle, groupId);
    } catch (e) {
        log('FEHLER Rechte: ' + (e.message || e));
        log(
            'Bibliothek ist angelegt, aber Rechte konnten nicht gesetzt werden (oft CORS oder fehlende SharePoint-Zustimmung). ' +
                'Bitte in SharePoint manuell: Bibliothek → Berechtigungen → Vererbung beenden → Besucher/Mitglieder entfernen → IT-Gruppe Mitwirken.'
        );
        toast('Bibliothek da – Rechte manuell prüfen');
    }

    const meta = {
        listTitle: listTitle,
        listId: String(listId),
        driveId: String(drive.id),
        webUrl: String(list.webUrl || drive.webUrl || ''),
        itGroupId: groupId,
        itGroupMail: plan.itGroupMail || '',
        securedAt: new Date().toISOString(),
        siteUrl: webUrl
    };
    saveItMeta(meta);
    if ($('suItGroup') && groupId && !$('suItGroup').value) $('suItGroup').value = groupId;
    refreshMetaUi();
    log('Fertig. ' + designHintDe());
    toast('IT-Bibliothek eingerichtet.');
    return meta;
}

function buildPayloadJson() {
    const bb = window.ms365BrowserBackup;
    if (!bb || typeof bb.buildBackup !== 'function') throw new Error('Browser-Backup-Modul fehlt.');
    const payload = bb.buildBackup();
    return { payload: payload, text: JSON.stringify(payload, null, 2) };
}

async function runUpload() {
    clearLog();
    const it = requireDriveId();
    const webUrl = getSiteUrl() || it.siteUrl;
    const folder = getFolder();
    const keepDated = !!($('suKeepDated') && $('suKeepDated').checked);
    log('Baue Browser-Backup …');
    const built = buildPayloadJson();
    const token = await ensureGraphToken();
    const currentPath = buildDriveRelativePath(folder, CURRENT_FILE);
    log('Upload → ' + it.listTitle + ' / ' + currentPath);
    const item = await putJsonOnDrive(it.driveId, currentPath, built.text, token);
    if (keepDated && window.ms365BrowserBackup && typeof window.ms365BrowserBackup.backupFilename === 'function') {
        const datedName = window.ms365BrowserBackup.backupFilename(new Date());
        log('Zusätzlich datiert → ' + datedName);
        await putJsonOnDrive(it.driveId, buildDriveRelativePath(folder, datedName), built.text, token);
    }
    saveLocalMeta({
        at: new Date().toISOString(),
        fileName: CURRENT_FILE,
        folder: folder,
        webUrl: (item && item.webUrl) || it.webUrl || '',
        siteUrl: webUrl,
        driveId: it.driveId,
        summary: describeRemoteBackup(built.payload)
    });
    try {
        localStorage.setItem('ms365-last-backup-export-at', new Date().toISOString());
    } catch {
        /* ignore */
    }
    refreshMetaUi();
    log('Fertig.');
    toast('Stammdaten in IT-Bibliothek geschrieben.');
}

async function runList() {
    clearLog();
    const it = requireDriveId();
    const folder = getFolder();
    const token = await ensureGraphToken();
    log('Liste „' + it.listTitle + '“ / ' + folder + ' …');
    const data = await listDriveFolder(it.driveId, folder, token);
    const items = (data && data.value) || [];
    const body = $('suRemoteBody');
    if (body) {
        body.replaceChildren();
        if (!items.length) {
            body.innerHTML = '<tr><td colspan="4" class="muted">Ordner leer (wird beim Upload angelegt).</td></tr>';
        } else {
            items.forEach(function (row) {
                if (!row || !row.file) return;
                const tr = document.createElement('tr');
                const when = row.lastModifiedDateTime
                    ? String(row.lastModifiedDateTime).replace('T', ' ').replace(/\.\d+Z$/, ' UTC')
                    : '';
                const size = row.size != null ? Math.round(Number(row.size) / 1024) + ' KB' : '';
                tr.innerHTML =
                    '<td>' +
                    escapeHtml(row.name || '') +
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
                    runDownload(row.id, row.name).catch(function (e) {
                        toast(e.message || String(e));
                        log('FEHLER: ' + (e.message || e));
                    });
                });
                td.appendChild(btn);
                body.appendChild(tr);
            });
        }
    }
    toast(items.filter(function (i) { return i && i.file; }).length + ' Datei(en)');
}

async function runDownload(itemId, name) {
    const it = requireDriveId();
    const token = await ensureGraphToken();
    log('Lade ' + (name || itemId) + ' …');
    const obj = await downloadDriveItem(it.driveId, itemId, token);
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
            'Backup aus IT-Bibliothek übernehmen und lokale Daten ersetzen?\n\n' + summary
        )
    ) {
        return;
    }
    bb.importPayload(obj);
    toast('Backup übernommen.');
    if (window.confirm('Seite jetzt neu laden?')) window.location.reload();
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function fillDefaults() {
    const hint = $('suDesignHint');
    if (hint) hint.textContent = designHintDe();
    try {
        const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
        const saved = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
        if (saved && $('suSiteUrl') && !$('suSiteUrl').value) $('suSiteUrl').value = saved;
        if ($('suItGroup') && !$('suItGroup').value) {
            const it = (setup && setup.stammdatenItLibrary) || loadItMeta();
            if (it && (it.itGroupMail || it.itGroupId)) {
                $('suItGroup').value = it.itGroupMail || it.itGroupId;
            } else if (setup && setup.matched && setup.matched.verwaltungGroupId) {
                $('suItGroup').value = String(setup.matched.verwaltungGroupId);
            }
        }
    } catch {
        /* ignore */
    }
    if ($('suFolder') && !$('suFolder').value) $('suFolder').value = DEFAULT_FOLDER;
    if ($('suLibraryTitle') && !$('suLibraryTitle').value) $('suLibraryTitle').value = IT_LIBRARY_TITLE;
    refreshMetaUi();
}

function boot() {
    fillDefaults();
    const setupBtn = $('suBtnSetupIt');
    if (setupBtn && setupBtn.dataset.bound !== '1') {
        setupBtn.dataset.bound = '1';
        setupBtn.addEventListener('click', function () {
            runSetupItLibrary().catch(function (e) {
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });
    }
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
                const it = requireDriveId();
                const folder = getFolder();
                const token = await ensureGraphToken();
                const data = await listDriveFolder(it.driveId, folder, token);
                const items = (data && data.value) || [];
                const cur = items.find(function (i) {
                    return i && i.file && String(i.name || '') === CURRENT_FILE;
                });
                if (!cur) throw new Error('Datei „' + CURRENT_FILE + '“ nicht gefunden.');
                await runDownload(cur.id, cur.name);
            })().catch(function (e) {
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });
    }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
