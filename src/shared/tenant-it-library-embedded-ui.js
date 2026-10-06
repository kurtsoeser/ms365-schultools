/**
 * Schulregister (Stammdaten): IT-Sicherungsbibliothek – eingebettet, sync mit Stammdaten-Übergabe.
 */
import { DEFAULT_FOLDER, IT_LIBRARY_TITLE, isItLibraryConfigured } from './stammdaten-sharepoint-sync-logic.js';
import {
    loadItMeta,
    readItLibraryFormDraft,
    writeItLibraryFormDraft
} from './stammdaten-sharepoint-sync-api.js';
import {
    readIntranetSiteUrl,
    readSchoolIntranetSiteUrl,
    writeIntranetSiteUrl,
    writeSchoolIntranetSiteUrl
} from './stammdaten-intranet-listen-ui.js';
import { configureItLibrarySetupUi, runSetupItLibrary } from '../tools/stammdaten-uebergabe/stammdaten-uebergabe.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function el(id) {
    return document.getElementById(id);
}

function setStatus(text, tone) {
    const status = el('tenantIntranetSiteStatus');
    if (!status) return;
    status.textContent = text || '';
    status.dataset.tone = tone || '';
}

function setItStatus(text, tone) {
    const status = el('tenantItLibraryStatus');
    if (!status) return;
    status.textContent = text || '';
    status.dataset.tone = tone || '';
}

function appendItLog(msg) {
    const logEl = el('tenantItLibraryLog');
    if (!logEl) return;
    logEl.textContent += (logEl.textContent ? '\n' : '') + String(msg || '');
    logEl.scrollTop = logEl.scrollHeight;
}

function clearItLog() {
    const logEl = el('tenantItLibraryLog');
    if (logEl) logEl.textContent = '';
}

export function readItLibrarySiteUrl() {
    const input = el('tenantItLibrarySiteUrl');
    if (input && String(input.value || '').trim()) {
        return String(input.value).trim();
    }
    const draft = readItLibraryFormDraft();
    if (draft.siteUrl) return String(draft.siteUrl).trim();
    const meta = loadItMeta() || {};
    if (meta.siteUrl) return String(meta.siteUrl).trim();
    return '';
}

export function writeItLibrarySiteUrl(url) {
    const u = String(url || '').trim();
    const input = el('tenantItLibrarySiteUrl');
    if (input && input.value !== u) input.value = u;
    const draft = readItLibraryFormDraft();
    writeItLibraryFormDraft({
        siteUrl: u,
        libraryTitle: draft.libraryTitle,
        itGroup: draft.itGroup,
        folder: draft.folder,
        keepDated: draft.keepDated
    });
}

function readTenantItFormState() {
    const draft = readItLibraryFormDraft();
    const titleEl = el('tenantItLibraryTitle');
    const groupEl = el('tenantItLibraryGroup');
    return {
        siteUrl: readItLibrarySiteUrl(),
        libraryTitle:
            String((titleEl && titleEl.value) || draft.libraryTitle || IT_LIBRARY_TITLE).trim() ||
            IT_LIBRARY_TITLE,
        itGroup: String((groupEl && groupEl.value) || draft.itGroup || '').trim(),
        folder: String(draft.folder || DEFAULT_FOLDER).trim() || DEFAULT_FOLDER,
        keepDated: typeof draft.keepDated === 'boolean' ? draft.keepDated : true
    };
}

function persistTenantItDraft() {
    const state = readTenantItFormState();
    writeItLibraryFormDraft(state);
}

function loadItLibrarySiteIntoField() {
    const input = el('tenantItLibrarySiteUrl');
    if (!input || String(input.value || '').trim()) return;
    let saved = readItLibrarySiteUrl();
    if (!saved) {
        saved = readIntranetSiteUrl();
    }
    if (saved) input.value = saved;
}

function fillItLibraryFields() {
    loadItLibrarySiteIntoField();
    const draft = readItLibraryFormDraft();
    const meta = loadItMeta() || {};
    const titleEl = el('tenantItLibraryTitle');
    const groupEl = el('tenantItLibraryGroup');
    if (titleEl && !String(titleEl.value || '').trim()) {
        titleEl.value =
            String(meta.listTitle || draft.libraryTitle || IT_LIBRARY_TITLE).trim() || IT_LIBRARY_TITLE;
    }
    if (groupEl && !String(groupEl.value || '').trim()) {
        const g = meta.itGroupMail || meta.itGroupId || draft.itGroup || '';
        if (g) groupEl.value = String(g);
        else {
            try {
                const setup =
                    window.ms365AppDataV2 && window.ms365AppDataV2.getSetup
                        ? window.ms365AppDataV2.getSetup()
                        : null;
                const matched = setup && setup.matched ? setup.matched : {};
                if (matched.verwaltungGroupId) groupEl.value = String(matched.verwaltungGroupId);
            } catch {
                /* ignore */
            }
        }
    }
    refreshItLibraryStatus();
}

function refreshItLibraryStatus() {
    const meta = loadItMeta() || {};
    if (isItLibraryConfigured(meta)) {
        const bits = ['IT-Bibliothek „' + (meta.listTitle || IT_LIBRARY_TITLE) + '“ verknüpft'];
        if (meta.securedAt) bits.push('Rechte gesetzt');
        if (meta.siteUrl) bits.push(meta.siteUrl);
        setItStatus(bits.join(' · '), 'ok');
    } else {
        setItStatus(
            'Noch nicht eingerichtet – Bibliothek anlegen oder auf der Übergabe-Seite fortsetzen.',
            'muted'
        );
    }
}

async function verifySharePointSite(webUrl, statusElId, onSuccess) {
    const status = el(statusElId);
    function setLocal(text, tone) {
        if (!status) return;
        status.textContent = text || '';
        status.dataset.tone = tone || '';
    }
    if (!webUrl) {
        setLocal('Bitte zuerst eine SharePoint-Webadresse eintragen.', 'warn');
        return;
    }
    const G = window.ms365SpoGraph;
    if (!G || typeof G.getGraphToken !== 'function' || typeof G.resolveSiteFromWebUrl !== 'function') {
        setLocal('SharePoint-Hilfen nicht geladen – Seite neu laden.', 'error');
        return;
    }
    setLocal('Prüfe Website …', 'pending');
    try {
        const token = await G.getGraphToken(SCOPES);
        const site = await G.resolveSiteFromWebUrl(token, webUrl);
        const name = site && site.displayName ? String(site.displayName).trim() : '';
        const id = site && site.id ? String(site.id) : '';
        if (!id) throw new Error('Site konnte nicht aufgelöst werden.');
        if (typeof onSuccess === 'function') onSuccess(webUrl);
        setLocal(
            (name ? 'Website gefunden: „' + name + '“' : 'Website erkannt.') +
                (site.webUrl ? ' · ' + site.webUrl : ''),
            'ok'
        );
    } catch (e) {
        setLocal('Nicht erreichbar oder keine Berechtigung: ' + (e && e.message ? e.message : String(e)), 'error');
    }
}

export async function verifyTenantIntranetSite() {
    const webUrl = readIntranetSiteUrl();
    await verifySharePointSite(webUrl, 'tenantIntranetSiteStatus', function (url) {
        writeIntranetSiteUrl(url);
    });
}

export async function verifyTenantItLibrarySite() {
    const webUrl = readItLibrarySiteUrl();
    await verifySharePointSite(webUrl, 'tenantItLibrarySiteStatus', function (url) {
        writeItLibrarySiteUrl(url);
        persistTenantItDraft();
    });
}

export async function verifyTenantSchoolIntranetSite() {
    const webUrl = readSchoolIntranetSiteUrl();
    await verifySharePointSite(webUrl, 'tenantSchoolIntranetSiteStatus', function (url) {
        writeSchoolIntranetSiteUrl(url);
    });
}

function bindItLibrarySiteField() {
    const input = el('tenantItLibrarySiteUrl');
    if (!input || input.dataset.itSiteBound === '1') return;
    input.dataset.itSiteBound = '1';
    let timer = null;
    input.addEventListener('input', function () {
        if (timer) clearTimeout(timer);
        timer = setTimeout(function () {
            writeItLibrarySiteUrl(String(input.value || '').trim());
            persistTenantItDraft();
        }, 400);
    });
    input.addEventListener('change', function () {
        writeItLibrarySiteUrl(String(input.value || '').trim());
        persistTenantItDraft();
    });
}

function bindItLibraryFields() {
    ['tenantItLibraryTitle', 'tenantItLibraryGroup'].forEach(function (id) {
        const input = el(id);
        if (!input || input.dataset.itDraftBound === '1') return;
        input.dataset.itDraftBound = '1';
        input.addEventListener('input', function () {
            persistTenantItDraft();
        });
        input.addEventListener('change', function () {
            persistTenantItDraft();
        });
    });
}

function bindItLibrarySetup() {
    const btn = el('tenantItLibrarySetup');
    if (!btn || btn.dataset.bound === '1') return;
    btn.dataset.bound = '1';
    configureItLibrarySetupUi({
        getFormState: readTenantItFormState,
        log: function (msg) {
            appendItLog(msg);
        }
    });
    btn.addEventListener('click', function () {
        persistTenantItDraft();
        clearItLog();
        runSetupItLibrary()
            .then(function () {
                refreshItLibraryStatus();
            })
            .catch(function (e) {
                const msg = e && e.message ? e.message : String(e);
                appendItLog('FEHLER: ' + msg);
                toast(msg);
            });
    });
}

function bindSiteVerify() {
    const btn = el('tenantIntranetSiteVerify');
    if (btn && btn.dataset.bound !== '1') {
        btn.dataset.bound = '1';
        btn.addEventListener('click', function () {
            verifyTenantIntranetSite().catch(function (e) {
                toast(e && e.message ? e.message : String(e));
            });
        });
    }
    const hubBtn = el('tenantSchoolIntranetSiteVerify');
    if (hubBtn && hubBtn.dataset.bound !== '1') {
        hubBtn.dataset.bound = '1';
        hubBtn.addEventListener('click', function () {
            verifyTenantSchoolIntranetSite().catch(function (e) {
                toast(e && e.message ? e.message : String(e));
            });
        });
    }
    const itSiteBtn = el('tenantItLibrarySiteVerify');
    if (itSiteBtn && itSiteBtn.dataset.bound !== '1') {
        itSiteBtn.dataset.bound = '1';
        itSiteBtn.addEventListener('click', function () {
            verifyTenantItLibrarySite().catch(function (e) {
                toast(e && e.message ? e.message : String(e));
            });
        });
    }
}

function loadSchoolIntranetSiteIntoField() {
    const input = el('tenantSchoolIntranetSiteUrl');
    if (!input || String(input.value || '').trim()) return;
    const saved = readSchoolIntranetSiteUrl();
    if (saved) input.value = saved;
}

function bindSchoolIntranetSiteField() {
    const input = el('tenantSchoolIntranetSiteUrl');
    if (!input || input.dataset.schoolIntranetBound === '1') return;
    input.dataset.schoolIntranetBound = '1';
    let timer = null;
    input.addEventListener('input', function () {
        if (timer) clearTimeout(timer);
        timer = setTimeout(function () {
            writeSchoolIntranetSiteUrl(String(input.value || '').trim());
        }, 400);
    });
    input.addEventListener('change', function () {
        writeSchoolIntranetSiteUrl(String(input.value || '').trim());
    });
}

export function mountTenantItLibraryEmbeddedUi() {
    loadSchoolIntranetSiteIntoField();
    bindSchoolIntranetSiteField();
    fillItLibraryFields();
    bindItLibrarySiteField();
    bindItLibraryFields();
    bindItLibrarySetup();
    bindSiteVerify();
    persistTenantItDraft();
}
