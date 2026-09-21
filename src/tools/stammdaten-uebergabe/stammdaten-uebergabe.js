/**
 * Stammdaten → IT-Dokumentbibliothek (Upload / Rechte / Liste / Download).
 */
import {
    DEFAULT_FOLDER,
    CURRENT_FILE,
    IT_LIBRARY_TITLE,
    designHintDe,
    isBroadSiteAudience,
    entraGroupLogonName,
    buildItLibraryPlan,
    SPO_ROLE
} from '../../shared/stammdaten-sharepoint-sync-logic.js';
import {
    loadItMeta,
    saveItMeta,
    loadLocalSyncMeta,
    listDriveFolder,
    downloadDriveItem,
    uploadCurrentBackup,
    downloadCurrentBackup,
    requireItLibrary
} from '../../shared/stammdaten-sharepoint-sync-api.js';

const SCOPES_GRAPH = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/Group.Read.All'
];

/** Formularwerte über MSAL-Redirect hinweg (Popup-Fallback / Header-Login). */
const FORM_DRAFT_KEY = 'ms365-su-form-draft-v1';
const PENDING_ACTION_KEY = 'ms365-su-pending-action-v1';
const PENDING_MAX_AGE_MS = 30 * 60 * 1000;

function $(id) {
    return document.getElementById(id);
}

function collectFormState() {
    return {
        siteUrl: String(($('suSiteUrl') && $('suSiteUrl').value) || '').trim(),
        libraryTitle: String(($('suLibraryTitle') && $('suLibraryTitle').value) || '').trim(),
        itGroup: String(($('suItGroup') && $('suItGroup').value) || '').trim(),
        folder: String(($('suFolder') && $('suFolder').value) || '').trim(),
        keepDated: !!($('suKeepDated') && $('suKeepDated').checked)
    };
}

function applyFormState(state) {
    if (!state || typeof state !== 'object') return;
    if ($('suSiteUrl') && state.siteUrl) $('suSiteUrl').value = String(state.siteUrl);
    if ($('suLibraryTitle') && state.libraryTitle) $('suLibraryTitle').value = String(state.libraryTitle);
    if ($('suItGroup') && state.itGroup) $('suItGroup').value = String(state.itGroup);
    if ($('suFolder') && state.folder) $('suFolder').value = String(state.folder);
    if ($('suKeepDated') && typeof state.keepDated === 'boolean') {
        $('suKeepDated').checked = state.keepDated;
    }
}

function persistFormDraft() {
    try {
        sessionStorage.setItem(FORM_DRAFT_KEY, JSON.stringify(collectFormState()));
    } catch {
        /* ignore */
    }
}

function restoreFormDraft() {
    try {
        const raw = sessionStorage.getItem(FORM_DRAFT_KEY);
        if (!raw) return false;
        applyFormState(JSON.parse(raw));
        return true;
    } catch {
        return false;
    }
}

function setPendingAction(action) {
    try {
        sessionStorage.setItem(
            PENDING_ACTION_KEY,
            JSON.stringify({
                action: String(action || ''),
                at: Date.now(),
                form: collectFormState()
            })
        );
    } catch {
        /* ignore */
    }
}

function takePendingAction() {
    try {
        const raw = sessionStorage.getItem(PENDING_ACTION_KEY);
        if (!raw) return null;
        sessionStorage.removeItem(PENDING_ACTION_KEY);
        const pending = JSON.parse(raw);
        if (!pending || !pending.action || !pending.at) return null;
        if (Date.now() - Number(pending.at) > PENDING_MAX_AGE_MS) return null;
        return pending;
    } catch {
        try {
            sessionStorage.removeItem(PENDING_ACTION_KEY);
        } catch {
            /* ignore */
        }
        return null;
    }
}

function clearPendingAction() {
    try {
        sessionStorage.removeItem(PENDING_ACTION_KEY);
    } catch {
        /* ignore */
    }
}

function bindFormDraftPersistence() {
    ['suSiteUrl', 'suLibraryTitle', 'suItGroup', 'suFolder', 'suKeepDated'].forEach(function (id) {
        const el = $(id);
        if (!el || el.dataset.draftBound === '1') return;
        el.dataset.draftBound = '1';
        const ev = el.type === 'checkbox' ? 'change' : 'input';
        el.addEventListener(ev, persistFormDraft);
        el.addEventListener('change', persistFormDraft);
    });
    const siteEl = $('suSiteUrl');
    if (siteEl && siteEl.dataset.rememberBound !== '1') {
        siteEl.dataset.rememberBound = '1';
        siteEl.addEventListener('change', function () {
            const url = String(siteEl.value || '').trim();
            if (url) rememberSite(url);
        });
    }
    if (!bindFormDraftPersistence._unloadBound) {
        bindFormDraftPersistence._unloadBound = true;
        window.addEventListener('pagehide', persistFormDraft);
        window.addEventListener('beforeunload', persistFormDraft);
    }
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

function refreshMetaUi() {
    const el = $('suLastMeta');
    const it = loadItMeta();
    const m = loadLocalSyncMeta();
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

/**
 * Graph POST /lists schlägt oft mit Access Denied fehl – dann SharePoint REST (Site-Besitzer).
 */
async function createDocumentLibraryViaSpo(siteWebUrl, title, description) {
    const G = getG();
    let host = '';
    try {
        host = new URL(siteWebUrl).hostname;
    } catch {
        throw new Error('Ungültige Site-URL.');
    }
    const spoScope = 'https://' + host + '/Sites.FullControl.All';
    let spoToken;
    try {
        spoToken = await G.getGraphToken([spoScope]);
    } catch (e) {
        throw new Error(
            'SharePoint-Token fehlt fürs Anlegen (Zustimmung Sites.FullControl.All / Office 365 SharePoint Online?). ' +
                (e && e.message ? e.message : e)
        );
    }
    const digest = await G.getSpoRequestDigest(siteWebUrl, spoToken);
    const created = await G.spoCreateDocumentLibrary(siteWebUrl, spoToken, digest, title, description);
    return { spoToken: spoToken, digest: digest, list: created };
}

function explainAccessDenied(err) {
    const msg = String((err && err.message) || err || '');
    if (!/access denied|AccessDenied|403/i.test(msg)) return msg;
    return (
        msg +
        '\n\nTypische Ursachen:\n' +
        '• Sie sind kein Besitzer der SharePoint-Website (nur Mitglied/Besucher).\n' +
        '• App-Zustimmung fehlt: Sites.ReadWrite.All und Sites.FullControl.All (SharePoint).\n' +
        '• Workaround: Bibliothek in SharePoint manuell anlegen (Dokumentbibliothek „' +
        IT_LIBRARY_TITLE +
        '“), dann hier erneut „einrichten“ – wir verbinden nur und setzen Rechte.'
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

function requireDriveId() {
    return requireItLibrary();
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

async function runSetupItLibrary(opts) {
    clearLog();
    const skipConfirm = !!(opts && opts.skipConfirm);
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
        !skipConfirm &&
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

    persistFormDraft();
    rememberSite(webUrl);
    setPendingAction('setup');

    const token = await ensureGraphToken();
    log('Löse Site auf …');
    const site = await resolveSite(token, webUrl);
    rememberSite(webUrl);

    let list = null;
    try {
        list = await findListByTitle(token, site.id, listTitle);
    } catch (e) {
        log('Hinweis Graph-Suche: ' + (e.message || e));
    }

    if (!list) {
        log('Bibliothek nicht gefunden – lege über SharePoint REST an (zuverlässiger als Graph) …');
        try {
            const spoCreated = await createDocumentLibraryViaSpo(webUrl, listTitle, plan.description);
            log('SPO: Bibliothek angelegt.');
            await getG().sleep(2000);
            list = await findListByTitle(token, site.id, listTitle);
            if (!list && spoCreated.list && (spoCreated.list.Id || spoCreated.list.id)) {
                list = {
                    id: spoCreated.list.Id || spoCreated.list.id,
                    displayName: listTitle,
                    webUrl: ''
                };
            }
        } catch (e1) {
            const spoMsg = String((e1 && e1.message) || e1 || '');
            log('SPO-Anlage: ' + spoMsg);
            if (!/access denied|AccessDenied|403|Zustimmung|FullControl/i.test(spoMsg)) {
                log('Fallback: versuche Graph POST /lists …');
                try {
                    list = await createDocumentLibrary(token, site.id, listTitle, plan.description);
                    await getG().sleep(1500);
                } catch (e2) {
                    throw new Error(explainAccessDenied(e1.message ? e1 : e2));
                }
            } else {
                throw new Error(explainAccessDenied(e1));
            }
        }
        if (!list) {
            try {
                const spoToken = await getG().getGraphToken([
                    'https://' + new URL(webUrl).hostname + '/Sites.FullControl.All'
                ]);
                const digest = await getG().getSpoRequestDigest(webUrl, spoToken);
                const spoList = await getG().spoGetListByTitle(webUrl, spoToken, digest, listTitle);
                if (spoList && (spoList.Id || spoList.id)) {
                    list = { id: spoList.Id || spoList.id, displayName: listTitle, webUrl: '' };
                }
            } catch {
                /* ignore */
            }
        }
        if (!list) {
            throw new Error(
                explainAccessDenied(
                    new Error(
                        'Bibliothek „' +
                            listTitle +
                            '“ konnte nicht angelegt/gefunden werden. Bitte manuell als Dokumentbibliothek anlegen und erneut versuchen.'
                    )
                )
            );
        }
    } else {
        log('Bibliothek existiert bereits (Graph).');
    }
    const listId = list.id || list.Id;
    if (!listId) throw new Error('Listen-ID fehlt.');
    let drive;
    try {
        drive = await getListDrive(token, site.id, listId);
    } catch (e) {
        throw new Error(
            'Drive der Bibliothek nicht lesbar: ' +
                (e.message || e) +
                ' – Sites.ReadWrite.All und Zugriff auf die Site prüfen.'
        );
    }
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
            'Bibliothek ist vorhanden, aber Rechte konnten nicht gesetzt werden (oft CORS oder fehlende SharePoint-Zustimmung). ' +
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
    clearPendingAction();
    persistFormDraft();
    refreshMetaUi();
    log('Fertig. ' + designHintDe());
    toast('IT-Bibliothek eingerichtet.');
    return meta;
}

async function runUpload() {
    clearLog();
    persistFormDraft();
    setPendingAction('upload');
    requireDriveId();
    const webUrl = getSiteUrl();
    const folder = getFolder();
    const keepDated = !!($('suKeepDated') && $('suKeepDated').checked);
    log('Baue Browser-Backup und lade hoch …');
    await uploadCurrentBackup({ folder: folder, keepDated: keepDated, siteUrl: webUrl });
    clearPendingAction();
    refreshMetaUi();
    log('Fertig.');
    toast('Stammdaten in IT-Bibliothek geschrieben.');
}

async function runList() {
    clearLog();
    persistFormDraft();
    setPendingAction('list');
    const it = requireDriveId();
    const folder = getFolder();
    const token = await ensureGraphToken();
    log('Liste „' + it.listTitle + '“ / ' + folder + ' …');
    const data = await listDriveFolder(it.driveId, folder, token);
    clearPendingAction();
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
    restoreFormDraft();
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

function handleActionError(e) {
    const msg = String((e && e.message) || e || '');
    if (/Weiterleitung zur Anmeldung/i.test(msg)) {
        persistFormDraft();
        log('Anmeldung nötig – Eingaben bleiben erhalten. Nach der Rückkehr wird fortgesetzt …');
        toast('Zur Anmeldung – Formular bleibt erhalten.');
        return;
    }
    clearPendingAction();
    log('FEHLER: ' + msg);
    toast(msg);
}

function resumePendingIfAny() {
    let hasPending = false;
    try {
        hasPending = !!sessionStorage.getItem(PENDING_ACTION_KEY);
    } catch {
        return;
    }
    if (!hasPending) return;

    (async function () {
        // Warten, bis MSAL den Redirect-Callback verarbeitet hat.
        for (let i = 0; i < 24; i++) {
            if (typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn()) break;
            await new Promise(function (r) {
                setTimeout(r, 250);
            });
        }
        const pending = takePendingAction();
        if (!pending || !pending.action) return;
        if (pending.form) applyFormState(pending.form);
        persistFormDraft();
        log('Anmeldung abgeschlossen – setze fort: ' + pending.action + ' …');
        try {
            if (pending.action === 'setup') await runSetupItLibrary({ skipConfirm: true });
            else if (pending.action === 'upload') await runUpload();
            else if (pending.action === 'list') await runList();
        } catch (e) {
            handleActionError(e);
        }
    })();
}

function boot() {
    fillDefaults();
    bindFormDraftPersistence();
    if (location.hash === '#setup') {
        const setupEl = document.getElementById('setup');
        if (setupEl && typeof setupEl.scrollIntoView === 'function') {
            setTimeout(function () {
                setupEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }, 80);
        }
    }
    const setupBtn = $('suBtnSetupIt');
    if (setupBtn && setupBtn.dataset.bound !== '1') {
        setupBtn.dataset.bound = '1';
        setupBtn.addEventListener('click', function () {
            runSetupItLibrary().catch(handleActionError);
        });
    }
    const up = $('suBtnUpload');
    if (up && up.dataset.bound !== '1') {
        up.dataset.bound = '1';
        up.addEventListener('click', function () {
            runUpload().catch(handleActionError);
        });
    }
    const list = $('suBtnList');
    if (list && list.dataset.bound !== '1') {
        list.dataset.bound = '1';
        list.addEventListener('click', function () {
            runList().catch(handleActionError);
        });
    }
    const loadCur = $('suBtnLoadCurrent');
    if (loadCur && loadCur.dataset.bound !== '1') {
        loadCur.dataset.bound = '1';
        loadCur.addEventListener('click', function () {
            (async function () {
                clearLog();
                persistFormDraft();
                log('Lade aktuelle Datei …');
                const preview = await downloadCurrentBackup({ folder: getFolder(), apply: false });
                const obj = preview.payload || {};
                const summary =
                    (obj.schoolName || obj.domain || preview.item.name || 'Backup') +
                    (obj.exportedAt ? ' · ' + String(obj.exportedAt).replace('T', ' ').slice(0, 19) : '');
                if (
                    !window.confirm(
                        'Backup aus IT-Bibliothek übernehmen und lokale Daten ersetzen?\n\n' + summary
                    )
                ) {
                    return;
                }
                window.ms365BrowserBackup.importPayload(obj);
                toast('Backup übernommen.');
                if (window.confirm('Seite jetzt neu laden?')) window.location.reload();
            })().catch(handleActionError);
        });
    }
    // Nach MSAL-Redirect: Formular wiederherstellen und Aktion fortsetzen
    setTimeout(resumePendingIfAny, 400);
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
