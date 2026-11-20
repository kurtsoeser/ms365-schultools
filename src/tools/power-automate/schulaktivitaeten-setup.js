/**
 * Schulaktivitäten: SharePoint-Liste prüfen + Power-Automate-Paket (Approvals Direktion) parametrieren.
 */
import { createSchulaktivitaetenLists } from '../sharepoint/sharepoint-liste-schulaktivitaeten.js';
import { LIST_TITLES } from '../schulaktivitaeten-planer/schulaktivitaeten-planer-schema.js';
import { computeAktSetupGlance, effectiveAktFlowAccount } from './schulaktivitaeten-setup-status.js';
import {
    AKT_FLOW_ASSET_ID,
    AKT_FLOW_SOURCE,
    applyAktFlowParamsToDefinition,
    assertAktFlowDefinitionMatchesSchool,
    normalizeSharePointListGuid,
    replaceAllLiteral
} from './schulaktivitaeten-setup-package.js';
import { SPO_ROLE } from '../../shared/stammdaten-sharepoint-sync-logic.js';

const G = () => {
    const api = window.ms365SpoGraph;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
};

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

const STORAGE_KEY = 'ms365-akt-setup-v1';
const STEP_STORAGE_KEY = 'ms365-akt-setup-step-v1';
const FLOW_DONE_KEY = 'ms365-pa-done-schulaktivitaeten';
const SETUP_STEP_COUNT = 5;
const TEMPLATE_BASE = '../assets/power-automate/schulaktivitaeten';
const FLOW_ASSET_ID = AKT_FLOW_ASSET_ID;
/** Platzhalter in assets – nur für Prefill-Fallback / Manifest-Replace. */
const SOURCE = AKT_FLOW_SOURCE;

function $(id) {
    return document.getElementById(id);
}

function log(msg) {
    const el = $('aktLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + msg;
    el.scrollTop = el.scrollHeight;
    const details = el.closest('details');
    if (details) details.open = true;
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function loadCfg() {
    try {
        return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}') || {};
    } catch {
        return {};
    }
}

function saveCfg(cfg) {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
    } catch {
        /* ignore */
    }
    if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.notifyLocalDataChanged === 'function') {
        window.ms365BrowserBackup.notifyLocalDataChanged('akt-setup');
    }
}

function readForm() {
    return {
        siteUrl: String(($('aktSiteUrl') && $('aktSiteUrl').value) || '')
            .trim()
            .replace(/\/$/, ''),
        listName:
            String(($('aktListName') && $('aktListName').value) || '').trim() || LIST_TITLES.aktivitaeten,
        listId: String(($('aktListId') && $('aktListId').value) || '').trim(),
        emailDirektion: String(($('aktEmailDirektion') && $('aktEmailDirektion').value) || '')
            .trim()
            .toLowerCase(),
        emailMailbox: String(($('aktEmailMailbox') && $('aktEmailMailbox').value) || '')
            .trim()
            .toLowerCase(),
        flowServiceAccount: String(($('aktFlowServiceAccount') && $('aktFlowServiceAccount').value) || '')
            .trim()
            .toLowerCase(),
        flowDisplayName:
            String(($('aktFlowName') && $('aktFlowName').value) || '').trim() ||
            'Schulaktivitäten - Genehmigung Direktion',
        mailAsTechnikUser: !($('aktMailAsTechnikUser') && !$('aktMailAsTechnikUser').checked)
    };
}

function writeForm(cfg) {
    if (!cfg) return;
    if ($('aktSiteUrl') && cfg.siteUrl) $('aktSiteUrl').value = cfg.siteUrl;
    if ($('aktListName') && cfg.listName) $('aktListName').value = cfg.listName;
    if ($('aktListId') && cfg.listId) $('aktListId').value = cfg.listId;
    if ($('aktEmailDirektion') && cfg.emailDirektion) $('aktEmailDirektion').value = cfg.emailDirektion;
    if ($('aktEmailMailbox') && cfg.emailMailbox) $('aktEmailMailbox').value = cfg.emailMailbox;
    if ($('aktFlowServiceAccount') && cfg.flowServiceAccount) {
        $('aktFlowServiceAccount').value = cfg.flowServiceAccount;
    }
    if ($('aktFlowName') && cfg.flowDisplayName) $('aktFlowName').value = cfg.flowDisplayName;
    if ($('aktMailAsTechnikUser')) {
        $('aktMailAsTechnikUser').checked = cfg.mailAsTechnikUser !== false;
    }
}

function persistFromForm() {
    const cfg = readForm();
    saveCfg(cfg);
    refreshGlance();
    return cfg;
}

function convertSharedMailboxToUserSend(definitionRoot, useUserMailboxSend) {
    if (!useUserMailboxSend || !definitionRoot) return;
    function patchAction(action) {
        if (
            !action ||
            !action.inputs ||
            !action.inputs.host ||
            action.inputs.host.operationId !== 'SharedMailboxSendEmailV2'
        ) {
            return;
        }
        const params = action.inputs.parameters || {};
        delete params['emailMessage/MailboxAddress'];
        action.inputs.parameters = params;
        action.inputs.host.operationId = 'SendEmailV2';
    }
    function walkActions(actions) {
        if (!actions || typeof actions !== 'object') return;
        Object.keys(actions).forEach((key) => {
            const action = actions[key];
            patchAction(action);
            if (action && action.actions) walkActions(action.actions);
            if (action && action.else && action.else.actions) walkActions(action.else.actions);
        });
    }
    walkActions(definitionRoot.actions);
}

function cryptoRandomGuid() {
    if (window.crypto && typeof window.crypto.randomUUID === 'function') {
        return window.crypto.randomUUID();
    }
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
        const r = (Math.random() * 16) | 0;
        const v = c === 'x' ? r : (r & 0x3) | 0x8;
        return v.toString(16);
    });
}

function applyParamsToDefinition(rawDef, cfg) {
    return applyAktFlowParamsToDefinition(
        rawDef,
        {
            ...cfg,
            listId: String(cfg.listId || '').trim(),
            listName: String(cfg.listName || '').trim() || LIST_TITLES.aktivitaeten
        },
        {
            flowAccount: effectiveAktFlowAccount(cfg),
            newFlowId: cryptoRandomGuid(),
            convertSharedMailboxToUserSend
        }
    );
}

async function fetchText(path) {
    const res = await fetch(path, { cache: 'no-store' });
    if (!res.ok) throw new Error('Vorlage nicht geladen: ' + path + ' (' + res.status + ')');
    return await res.text();
}

function downloadBlob(blob, filename) {
    const a = document.createElement('a');
    const url = URL.createObjectURL(blob);
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
}

async function buildPackageZip(cfg) {
    if (typeof JSZip === 'undefined') throw new Error('JSZip fehlt – Seite neu laden.');
    if (!cfg.siteUrl) throw new Error('SharePoint-Website fehlt.');
    const siteUrl = String(cfg.siteUrl || '')
        .trim()
        .replace(/\/+$/, '');
    if (!/\/sites\//i.test(siteUrl) && !/\/teams\//i.test(siteUrl)) {
        throw new Error(
            'Site-URL wirkt unvollständig (kein /sites/… beziehungsweise /teams/…). Erwartet z. B. https://….sharepoint.com/sites/MS365-Schultools'
        );
    }
    if (!cfg.listId) {
        throw new Error(
            'Listen-ID fehlt – zuerst „Listen anlegen/prüfen“. Power Automate braucht die echte Listen-GUID (nicht nur den Namen).'
        );
    }
    const flowAccount = effectiveAktFlowAccount(cfg);
    if (!flowAccount) {
        throw new Error('Technik-Konto fehlt – in Schritt 2 eintragen.');
    }

    const defPath = TEMPLATE_BASE + '/Microsoft.Flow/flows/' + FLOW_ASSET_ID + '/definition.json';
    const rootManifestPath = TEMPLATE_BASE + '/manifest.json';
    const flowsManifestPath = TEMPLATE_BASE + '/Microsoft.Flow/flows/manifest.json';
    const apisMapPath = TEMPLATE_BASE + '/Microsoft.Flow/flows/' + FLOW_ASSET_ID + '/apisMap.json';
    const connectionsMapPath =
        TEMPLATE_BASE + '/Microsoft.Flow/flows/' + FLOW_ASSET_ID + '/connectionsMap.json';

    log('Lade Flow-Vorlage …');
    const [defRaw, rootManifestRaw, flowsManifestRaw, apisMapRaw, connectionsMapRaw] = await Promise.all([
        fetchText(defPath),
        fetchText(rootManifestPath),
        fetchText(flowsManifestPath),
        fetchText(apisMapPath),
        fetchText(connectionsMapPath)
    ]);

    const definition = applyParamsToDefinition(defRaw, cfg);
    assertAktFlowDefinitionMatchesSchool(definition, {
        siteUrl,
        listId: String(cfg.listId || '').trim()
    });
    let rootManifest = JSON.parse(rootManifestRaw);
    rootManifest.details = rootManifest.details || {};
    rootManifest.details.displayName = cfg.flowDisplayName;
    rootManifest.details.description =
        'Schulaktivitäten: Trigger Liste, Microsoft Approvals an die Direktion, Status-Mail an die Lehrkraft. Multitenant: Site/Liste werden beim Download aus dem Setup eingesetzt.';
    rootManifest.details.createdTime = new Date().toISOString();
    rootManifest.details.sourceEnvironment = '';

    Object.keys(rootManifest.resources || {}).forEach((key) => {
        const r = rootManifest.resources[key];
        if (!r || !r.details) return;
        // Connections wie Freistellung: Existing (New hängt PA bei „Zugehörige Ressourcen“)
        if (r.type === 'Microsoft.PowerApps/apis/connections') {
            r.suggestedCreationType = 'Existing';
            r.creationType = 'Existing';
            if (String(r.details.displayName || '').indexOf('@') !== -1) {
                r.details.displayName = flowAccount;
            }
        }
        if (r.type === 'Microsoft.Flow/flows') {
            r.details.displayName = cfg.flowDisplayName;
            r.suggestedCreationType = 'New';
        }
    });

    let rootStr = JSON.stringify(rootManifest);
    rootStr = replaceAllLiteral(rootStr, SOURCE.connectionOwner, flowAccount);
    rootManifest = JSON.parse(rootStr);

    const zip = new JSZip();
    zip.file('manifest.json', JSON.stringify(rootManifest, null, 2));
    zip.file('Microsoft.Flow/flows/manifest.json', flowsManifestRaw);
    const folder = zip.folder('Microsoft.Flow/flows/' + FLOW_ASSET_ID);
    folder.file('definition.json', JSON.stringify(definition, null, 2));
    folder.file('apisMap.json', apisMapRaw);
    folder.file('connectionsMap.json', connectionsMapRaw);

    log('Erzeuge ZIP …');
    const blob = await zip.generateAsync({ type: 'blob' });
    const stamp = new Date().toISOString().replace(/[:.]/g, '').slice(0, 15);
    const filename = 'Schulaktivitaeten_' + stamp + '.zip';
    downloadBlob(blob, filename);
    log('Paket heruntergeladen: ' + filename);
    log('Dieses ZIP gilt nur für die eingestellte Schule:');
    log('  Site: ' + siteUrl);
    log('  Liste: ' + (cfg.listName || LIST_TITLES.aktivitaeten) + ' (' + cfg.listId + ')');
    log('Andere Schule / andere Site → Setup erneut öffnen und ZIP neu erzeugen.');
    log('Import (wichtig – sonst GetTable 403):');
    log('  1) Bei Power Automate als ' + flowAccount + ' anmelden.');
    log('  2) Package (Legacy) → Flow „Als neu erstellen“.');
    log('  3) Bei zugehörigen Ressourcen SharePoint/Approvals/Outlook auswählen bzw. Connection mit diesem Konto anlegen.');
    log('  4) SharePoint-Connection muss das Konto sein, das fullControl auf der Liste hat.');
    log('Flow einschalten und mit Testantrag (Status beantragt) prüfen.');
    return filename;
}

async function ensureList(cfg) {
    const write = (m) => log(m);
    write('Prüfe / lege Schulaktivitäten-Listen an …');
    const res = await createSchulaktivitaetenLists(cfg.siteUrl, write);
    const listId = res && res.lists && res.lists.aktivitaeten ? String(res.lists.aktivitaeten.id) : '';
    if (!listId) throw new Error('Listen-ID für Schulaktivitaeten fehlt.');
    if ($('aktListId')) $('aktListId').value = listId;
    if ($('aktListName')) $('aktListName').value = LIST_TITLES.aktivitaeten;
    const next = { ...cfg, listId, listName: LIST_TITLES.aktivitaeten };
    saveCfg(next);
    write('Liste bereit: ' + listId);
    if (res.lists.aktivitaeten.webUrl) write('URL: ' + res.lists.aktivitaeten.webUrl);
    refreshGlance();
    return { listId, webUrl: res.lists.aktivitaeten.webUrl || '' };
}

/**
 * SPO-Token + Digest für Listenrechte / GUID-Auflösung.
 * @param {string} siteWebUrl
 */
async function acquireSpoWriteContext(siteWebUrl) {
    const api = G();
    const url = String(siteWebUrl || '')
        .trim()
        .replace(/\/+$/, '');
    let host = '';
    try {
        host = new URL(url).hostname;
    } catch {
        throw new Error('Site-URL ungültig.');
    }
    const scopes = ['https://' + host + '/AllSites.Write', 'https://' + host + '/AllSites.Read'];
    let spoToken = '';
    let lastErr = '';
    for (let i = 0; i < scopes.length; i++) {
        try {
            if (typeof window.ms365AuthAcquireTokenSilent === 'function') {
                try {
                    spoToken = await window.ms365AuthAcquireTokenSilent([scopes[i]]);
                } catch {
                    spoToken = await api.getGraphToken([scopes[i]]);
                }
            } else {
                spoToken = await api.getGraphToken([scopes[i]]);
            }
            if (spoToken) break;
        } catch (e) {
            lastErr = e && e.message ? String(e.message) : String(e);
        }
    }
    if (!spoToken) throw new Error('SharePoint-Token fehlt (' + (lastErr || 'AllSites.Write') + ')');
    const digest = await api.getSpoRequestDigest(url, spoToken);
    return { api, siteWebUrl: url, spoToken, digest };
}

/**
 * Listen-GUID wie Power Automate sie erwartet (SPO REST Id, nicht nur Graph).
 * @param {string} siteWebUrl
 * @param {string} listTitle
 * @returns {Promise<string>}
 */
async function resolveListGuidViaSpoRest(siteWebUrl, listTitle) {
    const title = String(listTitle || LIST_TITLES.aktivitaeten).trim() || LIST_TITLES.aktivitaeten;
    const ctx = await acquireSpoWriteContext(siteWebUrl);
    const esc = title.replace(/'/g, "''");
    const res = await ctx.api.spoRestFetch(
        ctx.siteWebUrl,
        ctx.spoToken,
        ctx.digest,
        'GET',
        "/_api/web/lists/getbytitle('" + esc + "')?$select=Id,Title"
    );
    if (!res.ok) {
        throw new Error(
            'Liste „' + title + '“ per SPO nicht lesbar (HTTP ' + res.status + '). Zuerst Listen anlegen/prüfen.'
        );
    }
    const raw = (res.data && (res.data.Id || res.data.id)) || '';
    const listId = normalizeSharePointListGuid(raw);
    if (!listId) throw new Error('SPO lieferte keine Listen-GUID für „' + title + '“.');
    return listId;
}

/**
 * Nach Vererbungsbruch: Technik-/Import-Konto braucht Vollzugriff, sonst GetTable 403 beim ZIP-Import.
 * @param {string} siteWebUrl
 * @param {string[]} emails
 * @param {string[]} listTitles
 * @param {(m: string) => void} write
 */
async function grantFlowImportAccountsFullControl(siteWebUrl, emails, listTitles, write) {
    const uniq = [];
    const seen = new Set();
    (emails || []).forEach((e) => {
        const em = String(e || '')
            .trim()
            .toLowerCase();
        if (!em || !em.includes('@') || seen.has(em)) return;
        seen.add(em);
        uniq.push(em);
    });
    if (!uniq.length) {
        write('  ! Kein Flow-/Import-Konto – GetTable 403 beim Import möglich.');
        return;
    }
    const ctx = await acquireSpoWriteContext(siteWebUrl);
    const titles = (listTitles || []).map((t) => String(t || '').trim()).filter(Boolean);
    for (let t = 0; t < titles.length; t++) {
        const listTitle = titles[t];
        for (let i = 0; i < uniq.length; i++) {
            const em = uniq[i];
            try {
                const logon =
                    typeof ctx.api.membershipLogonName === 'function'
                        ? ctx.api.membershipLogonName(em)
                        : 'i:0#.f|membership|' + em;
                const principal = await ctx.api.spoEnsureUser(
                    ctx.siteWebUrl,
                    ctx.spoToken,
                    ctx.digest,
                    logon
                );
                await ctx.api.spoAddRoleAssignment(
                    ctx.siteWebUrl,
                    ctx.spoToken,
                    ctx.digest,
                    listTitle,
                    principal.id,
                    SPO_ROLE.fullControl
                );
                write('  + Import-Konto ' + em + ' → fullControl auf „' + listTitle + '“');
            } catch (e) {
                write(
                    '  ! Konnte ' +
                        em +
                        ' auf „' +
                        listTitle +
                        '“ nicht berechtigen: ' +
                        (e && e.message ? e.message : e)
                );
            }
        }
    }
}

async function currentSignedInEmail() {
    try {
        if (typeof window.ms365AuthGetAccount === 'function') {
            const a = window.ms365AuthGetAccount();
            const em = String((a && (a.username || a.mail || a.email)) || '')
                .trim()
                .toLowerCase();
            if (em.includes('@')) return em;
        }
    } catch {
        /* ignore */
    }
    try {
        const tok = await G().getGraphToken(['https://graph.microsoft.com/User.Read']);
        const me = await G().graphJson('GET', '/me?$select=mail,userPrincipalName', tok, undefined, 'v1.0');
        const em = String((me && (me.mail || me.userPrincipalName)) || '')
            .trim()
            .toLowerCase();
        if (em.includes('@')) return em;
    } catch {
        /* ignore */
    }
    return '';
}

async function resolveListIdIfNeeded(cfg) {
    const title = cfg.listName || LIST_TITLES.aktivitaeten;
    // Bevorzugt SPO-REST-GUID (exakt wie Power Automate GetTable)
    try {
        const listId = await resolveListGuidViaSpoRest(cfg.siteUrl, title);
        if ($('aktListId')) $('aktListId').value = listId;
        const next = { ...cfg, listId, listName: title };
        saveCfg(next);
        return next;
    } catch (spoErr) {
        log(
            'SPO-GUID-Auflösung: ' +
                (spoErr && spoErr.message ? spoErr.message : spoErr) +
                ' – versuche Graph …'
        );
    }
    const token = await G().getGraphToken(SCOPES);
    const site = await G().resolveSiteFromWebUrl(token, cfg.siteUrl);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const path =
        G().graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName';
    const data = await G().graphJson('GET', path, token, undefined, 'v1.0');
    const list = ((data && data.value) || [])[0];
    if (!list) throw new Error('Liste „' + title + '“ nicht gefunden – zuerst anlegen.');
    const listId = normalizeSharePointListGuid(list.id) || String(list.id);
    if ($('aktListId')) $('aktListId').value = listId;
    const next = { ...cfg, listId, listName: title };
    saveCfg(next);
    return next;
}

function flowImportedFlag() {
    try {
        return localStorage.getItem(FLOW_DONE_KEY) === '1';
    } catch {
        return false;
    }
}

function onboardingProgress() {
    try {
        if (window.ms365PaOnboarding && typeof window.ms365PaOnboarding.progress === 'function') {
            return window.ms365PaOnboarding.progress();
        }
    } catch {
        /* ignore */
    }
    return { done: 0, total: 0 };
}

function computeGlance() {
    const cfg = readForm();
    const ob = onboardingProgress();
    return computeAktSetupGlance(cfg, {
        flowImported: flowImportedFlag(),
        onboardingDone: ob.done,
        onboardingTotal: ob.total
    });
}

function refreshImportAccountHint() {
    const el = $('aktImportAccountHint');
    if (!el) return;
    el.textContent = effectiveAktFlowAccount(readForm()) || 'Technik-Konto aus Schritt 2';
}

function refreshGlance() {
    refreshImportAccountHint();
    const g = computeGlance();
    const host = $('aktSetupGlance');
    if (!host) return;
    host.querySelectorAll('[data-akt-glance]').forEach((card) => {
        const key = card.getAttribute('data-akt-glance');
        let ok = false;
        let warn = false;
        if (key === 'prep') {
            if (g.prepOk === null) {
                ok = false;
                warn = false;
            } else {
                ok = g.prepOk;
                warn = !ok;
            }
        } else if (key === 'emails') {
            ok = g.emailsOk;
            warn = !ok;
        } else if (key === 'list') {
            ok = g.listOk;
            warn = !ok;
        } else if (key === 'roles') {
            ok =
                typeof window.ms365AktRolesConfigured === 'function'
                    ? !!window.ms365AktRolesConfigured()
                    : !!String(readForm().emailDirektion || '').trim();
            warn = !ok;
        } else if (key === 'flow') {
            ok = g.flowOk;
            warn = !ok;
        }
        card.classList.toggle('is-ok', !!ok);
        card.classList.toggle('is-warn', !!warn && !ok);
        const val = card.querySelector('[data-akt-glance-value]');
        if (val) {
            if (key === 'prep' && g.prepOk === null) val.textContent = 'Optional';
            else val.textContent = ok ? 'Erledigt' : 'Offen';
        }
    });
}

function updatePhaseHint(step) {
    const el = $('aktPhaseHint');
    if (!el) return;
    const g = computeGlance();
    if (step === 1) el.textContent = 'Schritt 1: Schule vorbereiten (Checkliste).';
    else if (step === 2) {
        el.textContent = g.emailsOk
            ? 'Schritt 2: Konten gesetzt – weiter zur Liste.'
            : 'Schritt 2: Technik-Konto, Direktion und Absender eintragen (Suchen-Button).';
    } else if (step === 3) {
        el.textContent = g.listOk
            ? 'Schritt 3: Liste bereit – Rollen kurz prüfen.'
            : 'Schritt 3: Website wählen und Listen anlegen / prüfen.';
    } else if (step === 4) {
        const rolesOk =
            typeof window.ms365AktRolesConfigured === 'function' && window.ms365AktRolesConfigured();
        el.textContent = rolesOk
            ? 'Schritt 4: Genehmigung & Einreicher gesetzt – weiter zum Flow-Paket.'
            : 'Schritt 4: Entra-Gruppen / Personen für Genehmigung und Einreicher suchen und speichern.';
    } else {
        el.textContent = g.flowOk
            ? 'Schritt 5: Flow markiert – im Planer testen.'
            : 'Schritt 5: ZIP laden und in Power Automate (Legacy) importieren.';
    }
}

function loadSetupStep() {
    try {
        const n = parseInt(localStorage.getItem(STEP_STORAGE_KEY) || '1', 10);
        if (n >= 1 && n <= SETUP_STEP_COUNT) return n;
    } catch {
        /* ignore */
    }
    return 1;
}

function saveSetupStep(n) {
    try {
        localStorage.setItem(STEP_STORAGE_KEY, String(n));
    } catch {
        /* ignore */
    }
}

function showSetupStep(n) {
    const step = Math.max(1, Math.min(SETUP_STEP_COUNT, parseInt(n, 10) || 1));
    saveSetupStep(step);
    for (let i = 1; i <= SETUP_STEP_COUNT; i++) {
        const panel = $('aktSetupStep' + i);
        if (!panel) continue;
        const on = i === step;
        panel.hidden = !on;
        panel.setAttribute('aria-hidden', on ? 'false' : 'true');
    }
    document.querySelectorAll('#aktSetupGlance [data-akt-setup-step]').forEach((btn) => {
        const sn = parseInt(btn.getAttribute('data-akt-setup-step'), 10);
        const on = sn === step;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    const back = $('aktSetupBack');
    const next = $('aktSetupNext');
    if (back) back.disabled = step <= 1;
    if (next) next.textContent = step >= SETUP_STEP_COUNT ? 'Fertig' : 'Weiter';
    updatePhaseHint(step);
    refreshGlance();
}

function wireWizard() {
    document.querySelectorAll('[data-akt-setup-step]').forEach((btn) => {
        btn.addEventListener('click', () => showSetupStep(btn.getAttribute('data-akt-setup-step')));
    });
    const back = $('aktSetupBack');
    const next = $('aktSetupNext');
    if (back) {
        back.addEventListener('click', () => showSetupStep(Math.max(1, loadSetupStep() - 1)));
    }
    if (next) {
        next.addEventListener('click', () => {
            const cur = loadSetupStep();
            if (cur >= SETUP_STEP_COUNT) {
                toast('Setup abgeschlossen – Planer öffnen und Testantrag stellen.');
                return;
            }
            if (cur === 1 || cur === 2) persistFromForm();
            showSetupStep(cur + 1);
        });
    }
    showSetupStep(loadSetupStep());
}

async function onEnsureList() {
    try {
        const cfg = persistFromForm();
        if (!cfg.siteUrl) throw new Error('Bitte SharePoint-Website eintragen.');
        log('— Liste —');
        await ensureList(cfg);
        toast('Listen bereit.');
    } catch (e) {
        log('Fehler: ' + (e && e.message ? e.message : e));
        toast(e && e.message ? e.message : String(e));
    }
}

async function onBuildPackage() {
    try {
        let cfg = persistFromForm();
        log('— Flow-Paket —');
        if (!cfg.siteUrl) throw new Error('Bitte SharePoint-Website eintragen.');
        const flowAccount = effectiveAktFlowAccount(cfg);
        if (!flowAccount) throw new Error('Technik-Konto fehlt – in Schritt 2 eintragen.');

        // Immer frische SPO-Listen-GUID (GetTable)
        try {
            cfg = await resolveListIdIfNeeded({ ...cfg, listId: '' });
        } catch (resolveErr) {
            log(
                'Liste nicht gefunden – lege an: ' +
                    (resolveErr && resolveErr.message ? resolveErr.message : resolveErr)
            );
            await ensureList(cfg);
            cfg = readForm();
            cfg = await resolveListIdIfNeeded({ ...cfg, listId: '' });
        }
        if (!cfg.listId) throw new Error('Listen-ID fehlt nach Auflösung.');
        log('Verwende SPO-Listen-GUID: ' + cfg.listId);

        // Nach „breite Site-Rollen entfernt“: Import-Konto sonst GetTable 403
        log('Berechtige Flow-/Import-Konto auf den Paket-Listen …');
        const signedIn = await currentSignedInEmail();
        await grantFlowImportAccountsFullControl(
            cfg.siteUrl,
            [flowAccount, signedIn],
            [LIST_TITLES.aktivitaeten, LIST_TITLES.regelwerk],
            log
        );

        await buildPackageZip(cfg);
        toast('Flow-Paket heruntergeladen – Import als ' + flowAccount + ' mit neuen Connections.');
    } catch (e) {
        log('Fehler: ' + (e && e.message ? e.message : e));
        toast(e && e.message ? e.message : String(e));
    }
}

function prefillSiteFromApp() {
    const input = $('aktSiteUrl');
    if (!input || String(input.value || '').trim()) return;
    const candidates = [];
    try {
        const local = String(localStorage.getItem('ms365-akt-planer-site-v1') || '').trim();
        if (local) candidates.push(local);
    } catch {
        /* ignore */
    }
    try {
        const setup =
            window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup()
                : null;
        if (setup) {
            const planerSites =
                setup.planerSiteUrls && typeof setup.planerSiteUrls === 'object'
                    ? setup.planerSiteUrls
                    : {};
            candidates.push(
                planerSites.aktivitaeten,
                setup.planerDefaultSiteUrl,
                setup.intranetSiteUrl
            );
        }
    } catch {
        /* ignore */
    }
    candidates.push(SOURCE.siteUrl);
    for (let i = 0; i < candidates.length; i++) {
        const u = String(candidates[i] || '')
            .trim()
            .replace(/\/+$/, '');
        if (u && (/\/sites\//i.test(u) || /\/teams\//i.test(u))) {
            input.value = u;
            return;
        }
    }
}

function wire() {
    const cfg = loadCfg();
    if (!cfg.flowServiceAccount && cfg.emailMailbox) {
        cfg.flowServiceAccount = String(cfg.emailMailbox).trim().toLowerCase();
    }
    writeForm(cfg);
    prefillSiteFromApp();
    wireWizard();
    refreshGlance();

    const btnList = $('aktBtnList');
    const btnPkg = $('aktBtnPackage');
    const btnSave = $('aktBtnSave');
    if (btnList) btnList.addEventListener('click', () => onEnsureList());
    if (btnPkg) btnPkg.addEventListener('click', () => onBuildPackage());
    if (btnSave) {
        btnSave.addEventListener('click', () => {
            persistFromForm();
            toast('Konten in diesem Browser gespeichert.');
        });
    }

    const doneBox = $('aktDone');
    if (doneBox) {
        doneBox.checked = flowImportedFlag();
        doneBox.addEventListener('change', () => {
            try {
                if (doneBox.checked) localStorage.setItem(FLOW_DONE_KEY, '1');
                else localStorage.removeItem(FLOW_DONE_KEY);
            } catch {
                /* ignore */
            }
            refreshGlance();
            updatePhaseHint(loadSetupStep());
        });
    }

    [
        'aktSiteUrl',
        'aktListName',
        'aktListId',
        'aktEmailDirektion',
        'aktEmailMailbox',
        'aktFlowServiceAccount',
        'aktFlowName',
        'aktMailAsTechnikUser'
    ].forEach((id) => {
        const el = $(id);
        if (!el) return;
        el.addEventListener('change', () => {
            persistFromForm();
            refreshGlance();
        });
        el.addEventListener('input', refreshGlance);
    });

    const syncMailboxBtn = $('aktFlowAccountUseMailbox');
    if (syncMailboxBtn) {
        syncMailboxBtn.addEventListener('click', () => {
            const mb = String(($('aktEmailMailbox') && $('aktEmailMailbox').value) || '')
                .trim()
                .toLowerCase();
            if (!mb) {
                toast('Zuerst Absender-Postfach eintragen.');
                return;
            }
            if ($('aktFlowServiceAccount')) $('aktFlowServiceAccount').value = mb;
            persistFromForm();
            refreshGlance();
        });
    }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', wire);
else wire();

window.ms365AktSetup = {
    buildPackageZip,
    ensureList,
    loadSetupStep,
    showSetupStep,
    refreshGlance
};
