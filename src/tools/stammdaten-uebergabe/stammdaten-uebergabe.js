import { escapeHtml } from '../../shared/utils/strings.js';
/**
 * Stammdaten → IT-Dokumentbibliothek (Upload / Rechte / Liste / Download).
 */
import {
    DEFAULT_FOLDER,
    IT_LIBRARY_TITLE,
    designHintDe,
    designHintBulletsDe,
    isBroadSiteAudience,
    entraGroupLogonName,
    buildItLibraryPlan,
    SPO_ROLE,
    resolveItLibraryBrowserHref,
    isItLibraryConfigured,
    normalizeDriveListFolder,
    sortDriveBrowserItemsByColumn,
    isLikelyImportableBackupFileName
} from '../../shared/stammdaten-sharepoint-sync-logic.js';
import {
    loadItMeta,
    saveItMeta,
    loadLocalSyncMeta,
    writeItLibraryFormDraft,
    listDriveFolder,
    downloadDriveItem,
    downloadDriveItemBlob,
    uploadCurrentBackup,
    downloadCurrentBackup,
    requireItLibrary
} from '../../shared/stammdaten-sharepoint-sync-api.js';

const SCOPES_GRAPH = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/Group.Read.All'
];

/** Formularwerte dauerhaft (localStorage) + kurzfristig für MSAL-Redirect (sessionStorage). */
const FORM_DRAFT_KEY = 'ms365-su-form-draft-v1';
const PENDING_ACTION_KEY = 'ms365-su-pending-action-v1';
const PENDING_MAX_AGE_MS = 30 * 60 * 1000;
const SU_WIZARD_STEP_COUNT = 2;
const SU_WIZARD_STEP_STORAGE_KEY = 'ms365-su-wizard-step-v1';
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Relativer Pfad in der IT-Bibliothek (leer = Root). */
let suBrowseFolder = '';
/** @type {null | Promise<void>} */
let suBrowseInFlight = null;
/** @type {Array<object>} */
let suBrowseItems = [];
/** @type {{ key: 'name'|'modified'|'size', dir: 1|-1 }} */
let suBrowseSort = { key: 'name', dir: 1 };

/** @type {null | (() => ReturnType<typeof collectFormStateFromDom>)} */
let externalFormGetter = null;
/** @type {null | ((msg: string) => void)} */
let externalLogFn = null;

function $(id) {
    return document.getElementById(id);
}

function collectFormStateFromDom() {
    return {
        siteUrl: String(($('suSiteUrl') && $('suSiteUrl').value) || '').trim(),
        libraryTitle: String(($('suLibraryTitle') && $('suLibraryTitle').value) || '').trim(),
        itGroup: String(($('suItGroup') && $('suItGroup').value) || '').trim(),
        folder: String(($('suFolder') && $('suFolder').value) || '').trim(),
        keepDated: !!($('suKeepDated') && $('suKeepDated').checked)
    };
}

export function configureItLibrarySetupUi(opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    externalFormGetter = typeof o.getFormState === 'function' ? o.getFormState : null;
    externalLogFn = typeof o.log === 'function' ? o.log : null;
}

export function collectFormState() {
    if (externalFormGetter) {
        const s = externalFormGetter();
        if (s && typeof s === 'object') {
            return {
                siteUrl: String(s.siteUrl || '').trim(),
                libraryTitle: String(s.libraryTitle || IT_LIBRARY_TITLE).trim() || IT_LIBRARY_TITLE,
                itGroup: String(s.itGroup || '').trim(),
                folder: String(s.folder || DEFAULT_FOLDER).trim() || DEFAULT_FOLDER,
                keepDated: typeof s.keepDated === 'boolean' ? s.keepDated : true
            };
        }
    }
    return collectFormStateFromDom();
}

function applyFormState(state) {
    if (!state || typeof state !== 'object') return;
    if ($('suSiteUrl') && state.siteUrl) $('suSiteUrl').value = String(state.siteUrl);
    if ($('suLibraryTitle') && state.libraryTitle) $('suLibraryTitle').value = String(state.libraryTitle);
    if ($('suItGroup') && state.itGroup) $('suItGroup').value = String(state.itGroup);
    if ($('suFolder') && state.folder) $('suFolder').value = String(state.folder);
    if ($('suKeepDated')) {
        $('suKeepDated').checked = typeof state.keepDated === 'boolean' ? state.keepDated : true;
    }
}

function mergeItMetaPrefs(state) {
    if (!state || typeof state !== 'object') return;
    const cur = loadItMeta() || {};
    const next = Object.assign({}, cur);
    if (state.siteUrl) next.siteUrl = state.siteUrl;
    if (state.libraryTitle) next.listTitle = state.libraryTitle;
    if (state.itGroup) {
        if (GUID_RE.test(state.itGroup)) {
            next.itGroupId = state.itGroup;
            if (!next.itGroupMail) next.itGroupMail = '';
        } else {
            next.itGroupMail = state.itGroup;
        }
    }
    const changed =
        String(cur.siteUrl || '') !== String(next.siteUrl || '') ||
        String(cur.listTitle || '') !== String(next.listTitle || '') ||
        String(cur.itGroupId || '') !== String(next.itGroupId || '') ||
        String(cur.itGroupMail || '') !== String(next.itGroupMail || '');
    if (changed) saveItMeta(next);
}

function persistFormDraft(opts) {
    const state = collectFormState();
    const raw = JSON.stringify(state);
    try {
        localStorage.setItem(FORM_DRAFT_KEY, raw);
    } catch {
        /* ignore */
    }
    try {
        sessionStorage.setItem(FORM_DRAFT_KEY, raw);
    } catch {
        /* ignore */
    }
    // Stammdaten/Setup nur bei change/pagehide mergen – nicht bei jedem Tastendruck.
    if (opts && opts.mergeSetup) {
        if (state.siteUrl) rememberSite(state.siteUrl);
        mergeItMetaPrefs(state);
    }
}

function restoreFormDraft() {
    try {
        const localRaw = localStorage.getItem(FORM_DRAFT_KEY);
        if (localRaw) {
            applyFormState(JSON.parse(localRaw));
            return true;
        }
    } catch {
        /* ignore */
    }
    try {
        const sessionRaw = sessionStorage.getItem(FORM_DRAFT_KEY);
        if (!sessionRaw) return false;
        applyFormState(JSON.parse(sessionRaw));
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
        el.addEventListener(ev, function () {
            persistFormDraft();
        });
        el.addEventListener('change', function () {
            persistFormDraft({ mergeSetup: true });
        });
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
        window.addEventListener('pagehide', function () {
            persistFormDraft({ mergeSetup: true });
        });
        window.addEventListener('beforeunload', function () {
            persistFormDraft({ mergeSetup: true });
        });
    }
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function log(msg) {
    if (externalLogFn) {
        externalLogFn(String(msg || ''));
        return;
    }
    const el = $('suLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(msg || '');
    el.scrollTop = el.scrollHeight;
}

function clearLog() {
    if (externalLogFn) return;
    const el = $('suLog');
    if (el) el.textContent = '';
}

function getG() {
    const G = window.ms365SpoGraph;
    if (!G) throw new Error('SharePoint-Graph-Helfer nicht geladen.');
    return G;
}

function getSiteUrl() {
    const fromState = collectFormState().siteUrl;
    if (fromState) return fromState;
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
    return collectFormState().folder || DEFAULT_FOLDER;
}

function getLibraryTitle() {
    return collectFormState().libraryTitle || IT_LIBRARY_TITLE;
}

function rememberSite(url) {
    const u = String(url || '')
        .trim()
        .replace(/\/$/, '');
    if (!u) return;
    try {
        writeItLibraryFormDraft({ siteUrl: u });
    } catch {
        /* ignore */
    }
    try {
        if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.getSetup !== 'function') return;
        const setup = window.ms365AppDataV2.getSetup() || {};
        const patch = {};
        if (!String(setup.intranetSiteUrl || '').trim()) patch.intranetSiteUrl = u;
        if (!String(setup.schoolIntranetSiteUrl || '').trim()) patch.schoolIntranetSiteUrl = u;
        if (Object.keys(patch).length && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup(patch);
        }
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
        const href = resolveItLibraryBrowserHref(it, m);
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
                : 'Noch keine IT-Bibliothek – bitte in Schritt 1 einrichten.';
    }
    refreshSuWizardGlance();
}

function loadSuWizardStep() {
    try {
        const n = parseInt(localStorage.getItem(SU_WIZARD_STEP_STORAGE_KEY) || '0', 10);
        if (n >= 1 && n <= SU_WIZARD_STEP_COUNT) return n;
    } catch {
        /* ignore */
    }
    return isItLibraryConfigured(loadItMeta()) ? 2 : 1;
}

function saveSuWizardStep(n) {
    try {
        localStorage.setItem(SU_WIZARD_STEP_STORAGE_KEY, String(n));
    } catch {
        /* ignore */
    }
}

function updateSuPhaseHint(step) {
    const el = $('suPhaseHint');
    if (!el) return;
    const it = loadItMeta();
    const m = loadLocalSyncMeta();
    if (step === 1) {
        el.textContent = isItLibraryConfigured(it)
            ? 'Schritt 1: IT-Bibliothek ist eingerichtet – optional prüfen oder weiter zu Backup.'
            : 'Schritt 1: IT-Bibliothek einrichten (einmalig).';
        return;
    }
    if (!isItLibraryConfigured(it)) {
        el.textContent = 'Schritt 2: Zuerst IT-Bibliothek in Schritt 1 einrichten, dann Backup hochladen oder laden.';
        return;
    }
    if (m && m.at) {
        const when = String(m.at).replace('T', ' ').replace(/\.\d+Z$/, '').slice(0, 19);
        el.textContent = 'Schritt 2: Backup verwalten – zuletzt ' + (m.direction === 'pull' ? 'geladen' : 'gesichert') + ' ' + when + '.';
    } else {
        el.textContent = 'Schritt 2: Browser-Backup in die IT-Bibliothek hochladen oder von dort laden.';
    }
}

function refreshSuWizardGlance() {
    const it = loadItMeta();
    const m = loadLocalSyncMeta();
    const libOk = isItLibraryConfigured(it);
    const v1 = $('suGlanceValue1');
    const v2 = $('suGlanceValue2');
    const tab1 = $('suWizardTab1');
    const tab2 = document.querySelector('[data-su-wizard-step="2"]');
    if (v1) v1.textContent = libOk ? 'Eingerichtet' : 'Offen';
    if (tab1) {
        tab1.classList.toggle('is-ok', libOk);
        tab1.classList.toggle('is-warn', !libOk);
    }
    let step2Label = 'Offen';
    if (m && m.at) {
        step2Label = m.direction === 'pull' ? 'Geladen' : 'Gesichert';
    } else if (libOk) {
        step2Label = 'Bereit';
    }
    if (v2) v2.textContent = step2Label;
    if (tab2) {
        const step2Ok = !!(m && m.at);
        tab2.classList.toggle('is-ok', step2Ok);
        tab2.classList.toggle('is-warn', !step2Ok);
    }
}

function showSuWizardStep(n) {
    const step = Math.max(1, Math.min(SU_WIZARD_STEP_COUNT, parseInt(n, 10) || 1));
    saveSuWizardStep(step);
    for (let i = 1; i <= SU_WIZARD_STEP_COUNT; i++) {
        const panel = $('suWizardStep' + i);
        if (!panel) continue;
        const on = i === step;
        panel.hidden = !on;
        panel.setAttribute('aria-hidden', on ? 'false' : 'true');
    }
    document.querySelectorAll('#suWizardGlance [data-su-wizard-step]').forEach(function (btn) {
        const sn = parseInt(btn.getAttribute('data-su-wizard-step'), 10);
        const on = sn === step;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-selected', on ? 'true' : 'false');
        btn.setAttribute('tabindex', on ? '0' : '-1');
    });
    const back = $('suWizardBack');
    const next = $('suWizardNext');
    if (back) back.disabled = step <= 1;
    if (next) {
        next.textContent = step >= SU_WIZARD_STEP_COUNT ? 'Fertig' : 'Weiter zu Backup';
        next.setAttribute('aria-label', step >= SU_WIZARD_STEP_COUNT ? 'Assistent schließen' : 'Nächster Schritt');
    }
    updateSuPhaseHint(step);
    refreshSuWizardGlance();
    if (step === 2 && isItLibraryConfigured(loadItMeta())) {
        refreshLibraryBrowse({ silent: true }).catch(function () {
            /* Statuszeile zeigt Fehler */
        });
    }
}

function wireSuWizard() {
    if (!$('suWizardGlance')) return;
    document.querySelectorAll('[data-su-wizard-step]').forEach(function (btn) {
        if (btn.dataset.suWizardBound === '1') return;
        btn.dataset.suWizardBound = '1';
        btn.addEventListener('click', function () {
            showSuWizardStep(btn.getAttribute('data-su-wizard-step'));
        });
    });
    const back = $('suWizardBack');
    const next = $('suWizardNext');
    if (back && back.dataset.suWizardBound !== '1') {
        back.dataset.suWizardBound = '1';
        back.addEventListener('click', function () {
            showSuWizardStep(Math.max(1, loadSuWizardStep() - 1));
        });
    }
    if (next && next.dataset.suWizardBound !== '1') {
        next.dataset.suWizardBound = '1';
        next.addEventListener('click', function () {
            const cur = loadSuWizardStep();
            if (cur >= SU_WIZARD_STEP_COUNT) {
                toast('Backup-Schritt – Upload oder Laden oben.');
                return;
            }
            showSuWizardStep(cur + 1);
        });
    }
}

function resolveInitialSuWizardStep() {
    if (location.hash === '#setup') return 1;
    try {
        const raw = sessionStorage.getItem(PENDING_ACTION_KEY);
        if (raw) {
            const pending = JSON.parse(raw);
            if (pending && pending.action === 'setup') return 1;
            if (pending && (pending.action === 'upload' || pending.action === 'list')) return 2;
        }
    } catch {
        /* ignore */
    }
    return loadSuWizardStep();
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

export async function runSetupItLibrary(opts) {
    clearLog();
    const skipConfirm = !!(opts && opts.skipConfirm);
    const webUrl = getSiteUrl();
    if (!webUrl) throw new Error('SharePoint-Website fehlt.');
    const listTitle = getLibraryTitle();
    let groupRaw = String(collectFormState().itGroup || '').trim();
    if (!groupRaw) {
        try {
            const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
            const matched = setup && setup.matched ? setup.matched : {};
            if (matched.schulleitungGroupId) groupRaw = String(matched.schulleitungGroupId);
            else if (matched.verwaltungGroupId) groupRaw = String(matched.verwaltungGroupId);
        } catch {
            /* ignore */
        }
    }
    const plan = buildItLibraryPlan({
        listTitle: listTitle,
        itGroupId: GUID_RE.test(groupRaw) ? groupRaw : '',
        itGroupMail: GUID_RE.test(groupRaw) ? '' : groupRaw
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

    persistFormDraft({ mergeSetup: true });
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
    persistFormDraft({ mergeSetup: true });
    refreshMetaUi();
    log('Fertig. ' + designHintDe());
    toast('IT-Bibliothek eingerichtet.');
    showSuWizardStep(2);
    return meta;
}

async function runUpload() {
    clearLog();
    persistFormDraft({ mergeSetup: true });
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
    await refreshLibraryBrowse({ silent: true });
}

function formatBrowsePathLabel(folderPath) {
    const p = normalizeDriveListFolder(folderPath);
    return p || 'Bibliotheks-Root';
}

function renderBrowseCrumb(folderPath) {
    const nav = $('suBrowseCrumb');
    if (!nav) return;
    nav.replaceChildren();
    const rootBtn = document.createElement('button');
    rootBtn.type = 'button';
    rootBtn.textContent = 'IT-Bibliothek';
    if (!normalizeDriveListFolder(folderPath)) {
        rootBtn.setAttribute('aria-current', 'location');
    }
    rootBtn.addEventListener('click', function () {
        suBrowseFolder = '';
        refreshLibraryBrowse().catch(handleActionError);
    });
    nav.appendChild(rootBtn);
    const parts = normalizeDriveListFolder(folderPath).split('/').filter(Boolean);
    let acc = '';
    parts.forEach(function (seg, idx) {
        const sep = document.createElement('span');
        sep.className = 'su-browse__sep';
        sep.textContent = '/';
        sep.setAttribute('aria-hidden', 'true');
        nav.appendChild(sep);
        acc = acc ? acc + '/' + seg : seg;
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.textContent = seg;
        const pathHere = acc;
        if (idx === parts.length - 1) {
            btn.setAttribute('aria-current', 'location');
        } else {
            btn.addEventListener('click', function () {
                suBrowseFolder = pathHere;
                refreshLibraryBrowse().catch(handleActionError);
            });
        }
        nav.appendChild(btn);
    });
}

function formatItemSize(row) {
    if (!row || row.size == null) return row && row.folder ? '—' : '';
    const n = Number(row.size);
    if (!Number.isFinite(n)) return '';
    if (n < 1024) return n + ' B';
    if (n < 1024 * 1024) return Math.round(n / 1024) + ' KB';
    return (n / (1024 * 1024)).toFixed(1) + ' MB';
}

function triggerBrowserFileDownload(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName || 'download';
    a.rel = 'noopener';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () {
        URL.revokeObjectURL(url);
    }, 4000);
}

function updateBrowseSortIndicators() {
    const table = $('suBrowseTable');
    if (!table) return;
    table.querySelectorAll('th.is-sortable[data-su-sort]').forEach(function (th) {
        const label = th.dataset.sortLabel || th.textContent.replace(/[▲▼]\s*$/, '').trim();
        th.dataset.sortLabel = label;
        th.textContent = label;
        const key = th.getAttribute('data-su-sort');
        if (key === suBrowseSort.key) {
            th.setAttribute('aria-sort', suBrowseSort.dir === -1 ? 'descending' : 'ascending');
            const ind = document.createElement('span');
            ind.className = 'su-sort-ind';
            ind.setAttribute('aria-hidden', 'true');
            ind.textContent = suBrowseSort.dir === -1 ? '▼' : '▲';
            th.appendChild(ind);
        } else {
            th.setAttribute('aria-sort', 'none');
        }
    });
}

function applyBrowseColumnSort(key) {
    const k = key === 'modified' || key === 'size' ? key : 'name';
    if (suBrowseSort.key === k) {
        suBrowseSort.dir = suBrowseSort.dir === 1 ? -1 : 1;
    } else {
        suBrowseSort = { key: k, dir: 1 };
    }
    renderLibraryBrowseRows(suBrowseItems);
}

function wireSuBrowseTableSort() {
    const table = $('suBrowseTable');
    if (!table || table.dataset.sortBound === '1') return;
    table.dataset.sortBound = '1';
    table.querySelectorAll('th.is-sortable[data-su-sort]').forEach(function (th) {
        th.setAttribute('tabindex', '0');
        th.addEventListener('click', function () {
            applyBrowseColumnSort(th.getAttribute('data-su-sort'));
        });
        th.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter' || ev.key === ' ') {
                ev.preventDefault();
                applyBrowseColumnSort(th.getAttribute('data-su-sort'));
            }
        });
    });
}

function renderLibraryBrowseRows(items) {
    const body = $('suRemoteBody');
    if (!body) return;
    suBrowseItems = items || [];
    body.replaceChildren();
    const sorted = sortDriveBrowserItemsByColumn(suBrowseItems, suBrowseSort.key, suBrowseSort.dir);
    updateBrowseSortIndicators();
    if (!sorted.length) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.colSpan = 4;
        td.className = 'muted';
        td.textContent = 'Dieser Ordner ist leer.';
        tr.appendChild(td);
        body.appendChild(tr);
        return;
    }
    sorted.forEach(function (row) {
        if (!row || !row.id) return;
        const tr = document.createElement('tr');
        const isFolder = !!(row.folder && !row.file);
        if (isFolder) tr.classList.add('su-browse__row-folder');
        const when = row.lastModifiedDateTime
            ? String(row.lastModifiedDateTime).replace('T', ' ').replace(/\.\d+Z$/, '')
            : '';
        const nameCell = document.createElement('td');
        if (isFolder) {
            nameCell.innerHTML = '<i class="bi bi-folder2" aria-hidden="true"></i> ' + escapeHtml(row.name || '');
        } else {
            nameCell.innerHTML = '<i class="bi bi-file-earmark" aria-hidden="true"></i> ' + escapeHtml(row.name || '');
        }
        const whenCell = document.createElement('td');
        whenCell.textContent = when;
        const sizeCell = document.createElement('td');
        sizeCell.textContent = formatItemSize(row);
        const actionCell = document.createElement('td');
        const actions = document.createElement('div');
        actions.className = 'su-browse__actions';
        if (isFolder) {
            const openBtn = document.createElement('button');
            openBtn.type = 'button';
            openBtn.className = 'btn btn-sm';
            openBtn.innerHTML = '<i class="bi bi-folder2-open"></i>Öffnen';
            const childPath = normalizeDriveListFolder(suBrowseFolder)
                ? normalizeDriveListFolder(suBrowseFolder) + '/' + String(row.name || '')
                : String(row.name || '');
            openBtn.addEventListener('click', function () {
                suBrowseFolder = childPath;
                refreshLibraryBrowse().catch(handleActionError);
            });
            actions.appendChild(openBtn);
        } else {
            const dlBtn = document.createElement('button');
            dlBtn.type = 'button';
            dlBtn.className = 'btn btn-sm';
            dlBtn.innerHTML = '<i class="bi bi-download"></i>Herunterladen';
            dlBtn.addEventListener('click', function () {
                runDownloadFile(row.id, row.name).catch(handleActionError);
            });
            actions.appendChild(dlBtn);
            if (isLikelyImportableBackupFileName(row.name, { folder: suBrowseFolder })) {
                const impBtn = document.createElement('button');
                impBtn.type = 'button';
                impBtn.className = 'btn btn-sm alt';
                impBtn.innerHTML = '<i class="bi bi-box-arrow-in-down"></i>Backup übernehmen';
                impBtn.addEventListener('click', function () {
                    runDownload(row.id, row.name).catch(handleActionError);
                });
                actions.appendChild(impBtn);
            }
        }
        actionCell.appendChild(actions);
        tr.appendChild(nameCell);
        tr.appendChild(whenCell);
        tr.appendChild(sizeCell);
        tr.appendChild(actionCell);
        body.appendChild(tr);
    });
}

/**
 * @param {{ silent?: boolean, folder?: string }} [opts]
 */
async function refreshLibraryBrowse(opts) {
    const options = opts || {};
    if (suBrowseInFlight) return suBrowseInFlight;
    const statusEl = $('suBrowseStatus');
    const setStatus = function (msg) {
        if (statusEl) statusEl.textContent = msg || '';
    };
    suBrowseInFlight = (async function () {
        let it;
        try {
            it = requireDriveId();
        } catch (e) {
            renderBrowseCrumb('');
            renderLibraryBrowseRows([]);
            setStatus('IT-Bibliothek noch nicht eingerichtet.');
            if (!options.silent) toast((e && e.message) || String(e));
            return;
        }
        if (typeof options.folder === 'string') {
            suBrowseFolder = normalizeDriveListFolder(options.folder);
        }
        renderBrowseCrumb(suBrowseFolder);
        setStatus('Lade …');
        const token = await ensureGraphToken();
        const data = await listDriveFolder(it.driveId, suBrowseFolder, token);
        const items = (data && data.value) || [];
        renderLibraryBrowseRows(items);
        const folders = items.filter(function (i) {
            return i && i.folder;
        }).length;
        const files = items.filter(function (i) {
            return i && i.file;
        }).length;
        setStatus(
            formatBrowsePathLabel(suBrowseFolder) +
                ' · ' +
                folders +
                ' Ordner, ' +
                files +
                ' Datei' +
                (files === 1 ? '' : 'en')
        );
        if (!options.silent) {
            log('Inhalt: ' + formatBrowsePathLabel(suBrowseFolder) + ' (' + folders + ' Ordner, ' + files + ' Dateien).');
        }
    })()
        .catch(function (e) {
            setStatus('Fehler beim Laden.');
            if (!options.silent) {
                log('FEHLER: ' + (e && e.message ? e.message : e));
                toast(e && e.message ? e.message : String(e));
            }
            throw e;
        })
        .finally(function () {
            suBrowseInFlight = null;
        });
    return suBrowseInFlight;
}

async function runList() {
    persistFormDraft({ mergeSetup: true });
    setPendingAction('list');
    try {
        await refreshLibraryBrowse();
        clearPendingAction();
    } catch (e) {
        clearPendingAction();
        throw e;
    }
}

async function runDownloadFile(itemId, name) {
    const it = requireDriveId();
    const token = await ensureGraphToken();
    log('Lade Datei herunter: ' + (name || itemId) + ' …');
    const res = await downloadDriveItemBlob(it.driveId, itemId, token);
    triggerBrowserFileDownload(res.blob, name || 'datei');
    toast('Download gestartet.');
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

function fillDesignHintUi() {
    const list = $('suDesignHintList');
    if (list) {
        list.replaceChildren();
        designHintBulletsDe().forEach(function (line) {
            const li = document.createElement('li');
            li.textContent = line;
            list.appendChild(li);
        });
    }
}

function fillDefaults() {
    fillDesignHintUi();
    restoreFormDraft();
    try {
        const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
        const it = loadItMeta() || (setup && setup.stammdatenItLibrary) || {};
        const siteFromSetup = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
        const siteFromIt = it.siteUrl ? String(it.siteUrl).trim() : '';
        if ($('suSiteUrl') && !$('suSiteUrl').value) {
            $('suSiteUrl').value = siteFromIt || siteFromSetup;
        }
        if ($('suLibraryTitle')) {
            const savedTitle = it.listTitle ? String(it.listTitle).trim() : '';
            if (savedTitle) $('suLibraryTitle').value = savedTitle;
            else if (!$('suLibraryTitle').value) $('suLibraryTitle').value = IT_LIBRARY_TITLE;
        }
        if ($('suItGroup') && !$('suItGroup').value) {
            if (it.itGroupMail || it.itGroupId) {
                $('suItGroup').value = it.itGroupMail || it.itGroupId;
            } else if (setup && setup.matched && setup.matched.schulleitungGroupId) {
                $('suItGroup').value = String(setup.matched.schulleitungGroupId);
            } else if (setup && setup.matched && setup.matched.verwaltungGroupId) {
                $('suItGroup').value = String(setup.matched.verwaltungGroupId);
            }
        }
    } catch {
        /* ignore */
    }
    if ($('suFolder') && !$('suFolder').value) $('suFolder').value = DEFAULT_FOLDER;
    if ($('suLibraryTitle') && !$('suLibraryTitle').value) $('suLibraryTitle').value = IT_LIBRARY_TITLE;
    persistFormDraft({ mergeSetup: true });
    refreshMetaUi();
}

function handleActionError(e) {
    const msg = String((e && e.message) || e || '');
    if (/Weiterleitung zur Anmeldung/i.test(msg)) {
        persistFormDraft({ mergeSetup: true });
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
        persistFormDraft({ mergeSetup: true });
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
    if (!document.getElementById('suBtnSetupIt')) return;
    fillDefaults();
    bindFormDraftPersistence();
    wireSuWizard();
    wireSuBrowseTableSort();
    showSuWizardStep(resolveInitialSuWizardStep());
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
                persistFormDraft({ mergeSetup: true });
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
