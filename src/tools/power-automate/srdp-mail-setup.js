/**
 * sRDP: Bestätigungsmail-Flow (Shared Mailbox) parametrieren und als Legacy-ZIP laden.
 */
import { listTitleForYear, STORAGE_KEY as SRDP_WIZARD_KEY } from '../srdp-anmeldung/srdp-anmeldung-schema.js';

const G = () => {
    const api = window.ms365SpoGraph;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
};

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

const STORAGE_KEY = 'ms365-srdp-mail-setup-v1';
const FLOW_DONE_KEY = 'ms365-pa-done-srdp-mail';
const TEMPLATE_BASE = '../assets/power-automate/srdp-anmeldung';
const FLOW_ASSET_ID = 'c9e2a4b1-7f3d-4c8e-9a2b-5d6e8f0a1b2c';

/** Platzhalter in der Flow-Vorlage – werden beim Paketbau ersetzt. */
const SOURCE = {
    siteUrl: 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools',
    listId: 'bbbbbbbb-cccc-dddd-eeee-ffffffffffff',
    emailMailbox: 'matura@ms365.schule',
    emailBcc: 'verwaltung@ms365.schule',
    connectionOwner: 'kurt@kurtsoeser.at'
};

function $(id) {
    return document.getElementById(id);
}

function log(msg) {
    const el = $('srdpMailLog');
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
}

function loadWizardHints() {
    try {
        const raw = localStorage.getItem(SRDP_WIZARD_KEY);
        if (!raw) return {};
        const o = JSON.parse(raw);
        return o && typeof o === 'object' ? o : {};
    } catch {
        return {};
    }
}

function readForm() {
    return {
        siteUrl: String(($('srdpMailSiteUrl') && $('srdpMailSiteUrl').value) || '')
            .trim()
            .replace(/\/$/, ''),
        listName: String(($('srdpMailListName') && $('srdpMailListName').value) || '').trim(),
        listId: String(($('srdpMailListId') && $('srdpMailListId').value) || '').trim(),
        emailMailbox: String(($('srdpMailMailbox') && $('srdpMailMailbox').value) || '')
            .trim()
            .toLowerCase(),
        emailBcc: String(($('srdpMailBcc') && $('srdpMailBcc').value) || '')
            .trim()
            .toLowerCase(),
        flowServiceAccount: String(($('srdpMailFlowAccount') && $('srdpMailFlowAccount').value) || '')
            .trim()
            .toLowerCase(),
        flowDisplayName:
            String(($('srdpMailFlowName') && $('srdpMailFlowName').value) || '').trim() ||
            'sRDP-Anmeldung - Bestätigungsmail',
        mailAsTechnikUser: !($('srdpMailAsTechnik') && !$('srdpMailAsTechnik').checked)
    };
}

function writeForm(cfg) {
    if (!cfg) return;
    if ($('srdpMailSiteUrl') && cfg.siteUrl) $('srdpMailSiteUrl').value = cfg.siteUrl;
    if ($('srdpMailListName') && cfg.listName) $('srdpMailListName').value = cfg.listName;
    if ($('srdpMailListId') && cfg.listId) $('srdpMailListId').value = cfg.listId;
    if ($('srdpMailMailbox') && cfg.emailMailbox) $('srdpMailMailbox').value = cfg.emailMailbox;
    if ($('srdpMailBcc') && cfg.emailBcc) $('srdpMailBcc').value = cfg.emailBcc;
    if ($('srdpMailFlowAccount') && cfg.flowServiceAccount)
        $('srdpMailFlowAccount').value = cfg.flowServiceAccount;
    if ($('srdpMailFlowName') && cfg.flowDisplayName) $('srdpMailFlowName').value = cfg.flowDisplayName;
    if ($('srdpMailAsTechnik')) $('srdpMailAsTechnik').checked = cfg.mailAsTechnikUser !== false;
}

function persistFromForm() {
    const cfg = readForm();
    saveCfg(cfg);
    return cfg;
}

function replaceAll(haystack, needle, replacement) {
    if (!needle) return haystack;
    return String(haystack).split(needle).join(replacement == null ? '' : replacement);
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
    s = replaceAll(s, SOURCE.emailMailbox, cfg.emailMailbox || SOURCE.emailMailbox);
    s = replaceAll(s, SOURCE.emailBcc, cfg.emailBcc || '');
    const flowAccount = cfg.flowServiceAccount || cfg.emailMailbox || '';
    if (flowAccount) s = replaceAll(s, SOURCE.connectionOwner, flowAccount);

    const obj = JSON.parse(s);
    if (obj.properties) {
        obj.properties.displayName = cfg.flowDisplayName;
        const newId = cryptoRandomGuid();
        obj.name = newId;
        obj.id = '/providers/Microsoft.Flow/flows/' + newId;
        if (obj.properties.definition) {
            convertSharedMailboxToUserSend(obj.properties.definition, cfg.mailAsTechnikUser === true);
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

async function findListByTitle(token, siteId, listTitle) {
    const title = String(listTitle || '').trim();
    const path =
        G().graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await G().graphJson('GET', path, token, undefined, 'v1.0');
    return ((data && data.value) || [])[0] || null;
}

async function resolveListId(cfg) {
    if (cfg.listId) return cfg;
    if (!cfg.siteUrl) throw new Error('SharePoint-Website fehlt.');
    if (!cfg.listName) throw new Error('Listenname fehlt.');
    const token = await G().getGraphToken(SCOPES);
    const site = await G().resolveSiteFromWebUrl(token, cfg.siteUrl);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const list = await findListByTitle(token, siteId, cfg.listName);
    if (!list) {
        throw new Error(
            'Liste „' + cfg.listName + '“ nicht gefunden. Zuerst unter sRDP-Liste einrichten anlegen.'
        );
    }
    const next = { ...cfg, listId: String(list.id) };
    if ($('srdpMailListId')) $('srdpMailListId').value = next.listId;
    saveCfg(next);
    log('Liste gefunden: ' + list.displayName + ' (' + next.listId + ')');
    if (list.webUrl) log('URL: ' + list.webUrl);
    return next;
}

async function buildPackageZip(cfgIn) {
    if (typeof JSZip === 'undefined') throw new Error('JSZip fehlt – Seite neu laden.');
    let cfg = persistFromForm();
    if (!cfg.emailMailbox && !cfg.mailAsTechnikUser) {
        throw new Error('Freigegebenes Postfach (Absender) fehlt.');
    }
    if (!cfg.flowServiceAccount) {
        throw new Error('Technik-Konto für den Import fehlt.');
    }
    cfg = await resolveListId(cfg);
    if (!cfg.listId) throw new Error('Listen-ID fehlt.');

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
        'Bei neuer sRDP-/Matura-Anmeldung: Bestätigungsmail mit Zusammenfassung aus freigegebenem Postfach.';
    rootManifest.details.createdTime = new Date().toISOString();
    rootManifest.details.sourceEnvironment = '';

    Object.keys(rootManifest.resources || {}).forEach((key) => {
        const r = rootManifest.resources[key];
        if (!r || !r.details) return;
        if (r.type === 'Microsoft.PowerApps/apis/connections') {
            if (String(r.details.displayName || '').indexOf('@') !== -1) {
                r.details.displayName = cfg.flowServiceAccount;
            }
        }
        if (r.type === 'Microsoft.Flow/flows') {
            r.details.displayName = cfg.flowDisplayName;
            r.suggestedCreationType = 'New';
        }
    });

    let rootStr = JSON.stringify(rootManifest);
    rootStr = replaceAll(rootStr, SOURCE.connectionOwner, cfg.flowServiceAccount);
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
    const filename = 'sRDP_Bestaetigungsmail_' + stamp + '.zip';
    downloadBlob(blob, filename);
    log('Paket heruntergeladen: ' + filename);
    log('Import: make.powerautomate.com → Meine Flows → Importieren → Package (Legacy).');
    log(
        'Mit ' +
            cfg.flowServiceAccount +
            ' anmelden; Connections SharePoint + Outlook verbinden.' +
            (cfg.mailAsTechnikUser
                ? ' Mail geht als Technik-Konto.'
                : ' Absender = freigegebenes Postfach (' + cfg.emailMailbox + ') – „Senden als“ nötig.')
    );
    log('Test: eine Anmeldung in der App oder Liste anlegen → Mail prüfen.');
    return filename;
}

function prefillFromWizard() {
    const wiz = loadWizardHints();
    const cfg = loadCfg();
    const jahr = String(wiz.jahr || wiz.terminJahr || '2026').trim();
    const profileId = wiz.profileId || 'hak';
    const defaults = {
        siteUrl: cfg.siteUrl || wiz.siteUrl || '',
        listName: cfg.listName || listTitleForYear(jahr, profileId),
        listId: cfg.listId || '',
        emailMailbox: cfg.emailMailbox || '',
        emailBcc: cfg.emailBcc || '',
        flowServiceAccount: cfg.flowServiceAccount || '',
        flowDisplayName: cfg.flowDisplayName || 'sRDP-Anmeldung - Bestätigungsmail',
        mailAsTechnikUser: cfg.mailAsTechnikUser === true
    };
    writeForm(defaults);
    if ($('srdpMailDone')) {
        try {
            $('srdpMailDone').checked = localStorage.getItem(FLOW_DONE_KEY) === '1';
        } catch {
            /* ignore */
        }
    }
}

function bind() {
    prefillFromWizard();

    const saveBtn = $('srdpMailBtnSave');
    if (saveBtn) {
        saveBtn.addEventListener('click', () => {
            persistFromForm();
            toast('Einstellungen gespeichert.');
            log('Gespeichert.');
        });
    }

    const resolveBtn = $('srdpMailBtnResolve');
    if (resolveBtn) {
        resolveBtn.addEventListener('click', async () => {
            try {
                persistFromForm();
                await resolveListId(readForm());
                toast('Liste gefunden.');
            } catch (e) {
                toast((e && e.message) || String(e));
                log('Fehler: ' + ((e && e.message) || e));
            }
        });
    }

    const zipBtn = $('srdpMailBtnZip');
    if (zipBtn) {
        zipBtn.addEventListener('click', async () => {
            try {
                await buildPackageZip();
                toast('Flow-Paket heruntergeladen.');
            } catch (e) {
                toast((e && e.message) || String(e));
                log('Fehler: ' + ((e && e.message) || e));
            }
        });
    }

    const useMb = $('srdpMailAccountUseMailbox');
    if (useMb) {
        useMb.addEventListener('click', () => {
            const mb = String(($('srdpMailMailbox') && $('srdpMailMailbox').value) || '').trim();
            if ($('srdpMailFlowAccount') && mb) $('srdpMailFlowAccount').value = mb;
        });
    }

    const done = $('srdpMailDone');
    if (done) {
        done.addEventListener('change', () => {
            try {
                if (done.checked) localStorage.setItem(FLOW_DONE_KEY, '1');
                else localStorage.removeItem(FLOW_DONE_KEY);
            } catch {
                /* ignore */
            }
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bind);
} else {
    bind();
}
