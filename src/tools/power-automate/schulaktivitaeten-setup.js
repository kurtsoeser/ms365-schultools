/**
 * Schulaktivitäten: SharePoint-Liste prüfen + Power-Automate-Paket (Approvals Direktion) parametrieren.
 */
import { createSchulaktivitaetenLists } from '../sharepoint/sharepoint-liste-schulaktivitaeten.js';
import { LIST_TITLES } from '../schulaktivitaeten-planer/schulaktivitaeten-planer-schema.js';
import { computeAktSetupGlance, effectiveAktFlowAccount } from './schulaktivitaeten-setup-status.js';

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
const FLOW_ASSET_ID = 'a7f3c91e-2b4d-4e8a-9c1f-6d5e8b0a4f27';

/** Platzhalter in der Flow-Vorlage – werden beim Paketbau ersetzt. */
const SOURCE = {
    siteUrl: 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools',
    listId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
    emailDirektion: 'direktor@ms365.schule',
    emailMailbox: 'automate@ms365.schule',
    connectionOwner: 'kurt@kurtsoeser.at'
};

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

function replaceAll(haystack, needle, replacement) {
    if (!needle) return haystack;
    return String(haystack).split(needle).join(replacement);
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
    let s = typeof rawDef === 'string' ? rawDef : JSON.stringify(rawDef);
    s = replaceAll(s, SOURCE.siteUrl, cfg.siteUrl);
    s = replaceAll(s, SOURCE.listId, cfg.listId);
    s = replaceAll(s, SOURCE.emailDirektion, cfg.emailDirektion);
    s = replaceAll(s, SOURCE.emailMailbox, cfg.emailMailbox);
    const flowAccount = effectiveAktFlowAccount(cfg);
    if (flowAccount) s = replaceAll(s, SOURCE.connectionOwner, flowAccount);

    const obj = JSON.parse(s);
    if (obj.properties) {
        obj.properties.displayName = cfg.flowDisplayName;
        const newId = cryptoRandomGuid();
        obj.name = newId;
        obj.id = '/providers/Microsoft.Flow/flows/' + newId;
        if (obj.properties.definition) {
            convertSharedMailboxToUserSend(obj.properties.definition, cfg.mailAsTechnikUser !== false);
        }
    }
    return obj;
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
    if (!cfg.listId) throw new Error('Listen-ID fehlt – zuerst Liste anlegen oder ID eintragen.');
    if (!cfg.siteUrl) throw new Error('SharePoint-Website fehlt.');
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
    let rootManifest = JSON.parse(rootManifestRaw);
    rootManifest.details = rootManifest.details || {};
    rootManifest.details.displayName = cfg.flowDisplayName;
    rootManifest.details.description =
        'Schulaktivitäten: Trigger Liste, Microsoft Approvals an die Direktion, Status-Mail an die Lehrkraft.';
    rootManifest.details.createdTime = new Date().toISOString();
    rootManifest.details.sourceEnvironment = '';

    Object.keys(rootManifest.resources || {}).forEach((key) => {
        const r = rootManifest.resources[key];
        if (!r || !r.details) return;
        if (r.type === 'Microsoft.PowerApps/apis/connections') {
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
    rootStr = replaceAll(rootStr, SOURCE.connectionOwner, flowAccount);
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
    log('Nächster Schritt: make.powerautomate.com → Meine Flows → Importieren → Package (Legacy).');
    log(
        'Wichtig: In Power Automate mit ' +
            flowAccount +
            ' anmelden und Connections (SharePoint, Approvals, Outlook) verbinden.'
    );
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

async function resolveListIdIfNeeded(cfg) {
    if (cfg.listId) return cfg;
    const token = await G().getGraphToken(SCOPES);
    const site = await G().resolveSiteFromWebUrl(token, cfg.siteUrl);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const title = cfg.listName || LIST_TITLES.aktivitaeten;
    const path =
        G().graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName';
    const data = await G().graphJson('GET', path, token, undefined, 'v1.0');
    const list = ((data && data.value) || [])[0];
    if (!list) throw new Error('Liste „' + title + '“ nicht gefunden – zuerst anlegen.');
    const listId = String(list.id);
    if ($('aktListId')) $('aktListId').value = listId;
    const next = { ...cfg, listId };
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
            ok = !!String(readForm().emailDirektion || '').trim();
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
            : 'Schritt 2: Technik-Konto, Direktion (Approver) und Absender eintragen.';
    } else if (step === 3) {
        el.textContent = g.listOk
            ? 'Schritt 3: Liste bereit – Rollen kurz prüfen.'
            : 'Schritt 3: Website wählen und Listen anlegen / prüfen.';
    } else if (step === 4) {
        el.textContent =
            'Schritt 4: Genehmiger = Direktion aus Schritt 2; Planer-Rollen im Planer (Demo Lehrer/Admin).';
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
        if (!cfg.listId) {
            log('Keine Listen-ID – löse Liste auf bzw. lege an …');
            try {
                cfg = await resolveListIdIfNeeded(cfg);
            } catch {
                await ensureList(cfg);
                cfg = readForm();
            }
        }
        log('— Flow-Paket —');
        await buildPackageZip(cfg);
        toast('Flow-Paket heruntergeladen.');
    } catch (e) {
        log('Fehler: ' + (e && e.message ? e.message : e));
        toast(e && e.message ? e.message : String(e));
    }
}

function prefillSiteFromApp() {
    const input = $('aktSiteUrl');
    if (!input || String(input.value || '').trim()) return;
    try {
        const local = String(localStorage.getItem('ms365-akt-planer-site-v1') || '').trim();
        if (local) {
            input.value = local;
            return;
        }
    } catch {
        /* ignore */
    }
    try {
        const setup =
            window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup()
                : null;
        if (setup && setup.intranetSiteUrl) input.value = String(setup.intranetSiteUrl).trim();
    } catch {
        /* ignore */
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
