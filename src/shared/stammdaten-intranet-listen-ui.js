/**
 * Einbettung: Stammdaten → SharePoint-Intranet-Listen anlegen/abgleichen.
 * Mount: <div data-ms365-intranet-listen-mount data-mount-prefix="tsdSpo"></div>
 */
import {
    applyStammdatenPackagePermissions,
    loadPermissionsConfig
} from '../tools/sharepoint/stammdaten-liste-permissions.js';
import {
    buildStammdatenGroupFields,
    initEmbeddedPermissionsUi,
    readPermissionsFromPickers,
    persistPickersToStorage
} from '../tools/sharepoint/stammdaten-permissions-ui.js';
import {
    DEFAULT_INTRANET_LIST_TITLES,
    resolveIntranetListTitle
} from './intranet-list-title-logic.js';
import {
    refreshIntranetListOpenLinks,
    saveIntranetListLink,
    showIntranetSyncToast
} from './tenant-intranet-list-links.js';

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function confirmAsync(msg) {
    if (typeof window.ms365AppDialogConfirm === 'function') {
        return window.ms365AppDialogConfirm(msg, { title: 'SharePoint-Listen', confirmLabel: 'Fortfahren', cancelLabel: 'Abbrechen' });
    }
    return Promise.resolve(window.confirm(msg));
}

function api() {
    const a = window.ms365SpoStammdatenListen;
    if (!a || typeof a.syncSelectedLists !== 'function') {
        throw new Error('Listen-Synchronisation nicht geladen (sharepoint-liste-stammdaten.js).');
    }
    return a;
}

function el(id) {
    return document.getElementById(id);
}

function readRunOpts(p) {
    return readIntranetRunOpts(p);
}

function collectListOpts(p) {
    return buildListOptsForKinds(
        ['schueler', 'faecher', 'fachgruppen', 'arges', 'klassen', 'lehrer'].filter(function (kind) {
            const map = {
                schueler: 'WantSchueler',
                faecher: 'WantFaecher',
                fachgruppen: 'WantFachgruppen',
                arges: 'WantArge',
                klassen: 'WantKlassen',
                lehrer: 'WantLehrer'
            };
            const id = map[kind];
            return el(p + id) && el(p + id).checked;
        }),
        p
    );
}

function logTo(p, msg) {
    const logEl = el(p + 'Log');
    if (!logEl) return;
    logEl.textContent += (logEl.textContent ? '\n' : '') + msg;
    logEl.scrollTop = logEl.scrollHeight;
}

export const DEFAULT_LIST_TITLES = DEFAULT_INTRANET_LIST_TITLES;
export { resolveIntranetListTitle };

const INTRANET_LIST_KINDS = ['schueler', 'faecher', 'fachgruppen', 'arges', 'klassen', 'lehrer'];

const INTRANET_LIST_TITLE_SUFFIX = {
    schueler: 'SchuelerName',
    faecher: 'FaecherName',
    fachgruppen: 'FachgruppenName',
    arges: 'ArgeName',
    klassen: 'KlassenName',
    lehrer: 'LehrerName'
};

const TENANT_INTRANET_LIST_INPUT_IDS = {
    schueler: 'tenantIntranetListSchueler',
    faecher: 'tenantIntranetListFaecher',
    fachgruppen: 'tenantIntranetListFachgruppen',
    arges: 'tenantIntranetListArges',
    klassen: 'tenantIntranetListKlassen',
    lehrer: 'tenantIntranetListLehrer'
};

function readIntranetListTitlesFromSetup() {
    try {
        const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
        const t = setup && setup.intranetListTitles && typeof setup.intranetListTitles === 'object' ? setup.intranetListTitles : {};
        return t;
    } catch {
        return {};
    }
}

function storedIntranetListTitle(kind) {
    const tenantId = TENANT_INTRANET_LIST_INPUT_IDS[kind];
    const tenantEl = tenantId ? el(tenantId) : null;
    if (tenantEl) {
        const tv = String(tenantEl.value || '').trim();
        if (tv) return tv;
    }
    const fromSetup = readIntranetListTitlesFromSetup();
    return String(fromSetup[kind] != null ? fromSetup[kind] : '').trim();
}

export function readIntranetListTitlesForUi() {
    const fromSetup = readIntranetListTitlesFromSetup();
    const out = {};
    INTRANET_LIST_KINDS.forEach(function (kind) {
        const tenantId = TENANT_INTRANET_LIST_INPUT_IDS[kind];
        const tenantEl = tenantId ? el(tenantId) : null;
        if (tenantEl && String(tenantEl.value || '').trim()) {
            out[kind] = String(tenantEl.value).trim();
        } else {
            out[kind] = String(fromSetup[kind] != null ? fromSetup[kind] : '').trim();
        }
    });
    return out;
}

export function writeIntranetListTitles(partial) {
    const patch = {};
    INTRANET_LIST_KINDS.forEach(function (kind) {
        if (!partial || partial[kind] === undefined) return;
        const v = String(partial[kind] != null ? partial[kind] : '').trim();
        patch[kind] = v;
        const tenantId = TENANT_INTRANET_LIST_INPUT_IDS[kind];
        const tenantEl = tenantId ? el(tenantId) : null;
        if (tenantEl && tenantEl.value !== v) tenantEl.value = v;
        const suffix = INTRANET_LIST_TITLE_SUFFIX[kind];
        document.querySelectorAll('[data-intranet-prefix]').forEach(function (mount) {
            const p = mount.dataset.intranetPrefix || 'tsdSpo';
            const mountEl = suffix ? el(p + suffix) : null;
            if (mountEl && mountEl.type === 'hidden') {
                mountEl.value = resolveIntranetListTitle(kind, v);
            }
        });
    });
    if (!Object.keys(patch).length) return;
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup({ intranetListTitles: patch });
        }
    } catch {
        /* ignore */
    }
}

export function restoreIntranetListTitlesToTenantFields() {
    const fromSetup = readIntranetListTitlesFromSetup();
    INTRANET_LIST_KINDS.forEach(function (kind) {
        const tenantId = TENANT_INTRANET_LIST_INPUT_IDS[kind];
        const tenantEl = tenantId ? el(tenantId) : null;
        if (!tenantEl || String(tenantEl.value || '').trim()) return;
        const v = String(fromSetup[kind] != null ? fromSetup[kind] : '').trim();
        if (v) tenantEl.value = v;
    });
}

export function syncIntranetListTitlesToMounts() {
    document.querySelectorAll('[data-intranet-prefix]').forEach(function (mount) {
        const p = mount.dataset.intranetPrefix || 'tsdSpo';
        INTRANET_LIST_KINDS.forEach(function (kind) {
            const suffix = INTRANET_LIST_TITLE_SUFFIX[kind];
            const mountEl = suffix ? el(p + suffix) : null;
            if (!mountEl) return;
            const stored = storedIntranetListTitle(kind);
            if (mountEl.type === 'hidden') {
                mountEl.value = resolveIntranetListTitle(kind, stored);
            } else if (!String(mountEl.value || '').trim() && stored) {
                mountEl.value = stored;
            }
        });
    });
}

const INTRANET_KIND_LABELS = {
    schueler: 'Schülerinnen',
    faecher: 'Fächer',
    fachgruppen: 'Fachgruppen',
    arges: 'ARGEs',
    klassen: 'Klassen',
    lehrer: 'Lehrerinnen'
};

export function readIntranetSiteUrl() {
    const tenant = el('tenantIntranetSiteUrl');
    if (tenant && String(tenant.value || '').trim()) {
        return String(tenant.value).trim();
    }
    const mounts = document.querySelectorAll('[data-intranet-prefix]');
    for (let i = 0; i < mounts.length; i++) {
        const p = mounts[i].dataset.intranetPrefix || 'tsdSpo';
        const urlEl = el(p + 'SiteUrl');
        if (urlEl && String(urlEl.value || '').trim()) {
            return String(urlEl.value).trim();
        }
    }
    try {
        const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
        const saved = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
        return saved || '';
    } catch {
        return '';
    }
}

export function readSchoolIntranetSiteUrl() {
    const hub = el('tenantSchoolIntranetSiteUrl');
    if (hub && String(hub.value || '').trim()) {
        return String(hub.value).trim();
    }
    try {
        const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup ? window.ms365AppDataV2.getSetup() : null;
        if (setup && setup.schoolIntranetSiteUrl) {
            return String(setup.schoolIntranetSiteUrl).trim();
        }
        if (setup && setup.intranetSiteUrl) {
            return String(setup.intranetSiteUrl).trim();
        }
    } catch {
        /* ignore */
    }
    return '';
}

export function writeSchoolIntranetSiteUrl(url) {
    const u = String(url || '').trim();
    const hub = el('tenantSchoolIntranetSiteUrl');
    if (hub && hub.value !== u) hub.value = u;
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup({ schoolIntranetSiteUrl: u });
        }
    } catch {
        /* ignore */
    }
}

export function writeIntranetSiteUrl(url) {
    const u = String(url || '').trim();
    const tenant = el('tenantIntranetSiteUrl');
    if (tenant && tenant.value !== u) tenant.value = u;
    document.querySelectorAll('[data-intranet-prefix]').forEach(function (mount) {
        const p = mount.dataset.intranetPrefix || 'tsdSpo';
        const urlEl = el(p + 'SiteUrl');
        if (urlEl && urlEl.value !== u) urlEl.value = u;
    });
    if (!u) return;
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup({ intranetSiteUrl: u });
        }
    } catch {
        /* ignore */
    }
}

function kindFromListTitleSuffix(suffix) {
    const s = String(suffix || '').trim();
    let kind = '';
    INTRANET_LIST_KINDS.forEach(function (k) {
        if (INTRANET_LIST_TITLE_SUFFIX[k] === s) kind = k;
    });
    return kind;
}

function readListTitle(p, suffix, fallback) {
    const input = el(p + suffix);
    const v = input ? String(input.value || '').trim() : '';
    if (v) return v;
    const kind = kindFromListTitleSuffix(suffix);
    if (kind) return resolveIntranetListTitle(kind, storedIntranetListTitle(kind));
    return fallback;
}

export function buildListOptsForKinds(kinds, prefix) {
    const p = prefix || 'tsdSpo';
    const set = new Set((kinds || []).map(function (k) {
        return String(k || '').trim();
    }));
    return {
        schueler: set.has('schueler'),
        faecher: set.has('faecher'),
        fachgruppen: set.has('fachgruppen'),
        arges: set.has('arges'),
        klassen: set.has('klassen'),
        lehrer: set.has('lehrer'),
        schuelerTitle: readListTitle(p, 'SchuelerName', DEFAULT_LIST_TITLES.schueler),
        faecherTitle: readListTitle(p, 'FaecherName', DEFAULT_LIST_TITLES.faecher),
        fachgruppenTitle: readListTitle(p, 'FachgruppenName', DEFAULT_LIST_TITLES.fachgruppen),
        argesTitle: readListTitle(p, 'ArgeName', DEFAULT_LIST_TITLES.arges),
        klassenTitle: readListTitle(p, 'KlassenName', DEFAULT_LIST_TITLES.klassen),
        lehrerTitle: readListTitle(p, 'LehrerName', DEFAULT_LIST_TITLES.lehrer)
    };
}

export function readIntranetRunOpts(prefix) {
    const p = prefix || 'tsdSpo';
    const alwaysNew = el(p + 'AlwaysNew');
    const removeOrphans = el(p + 'RemoveOrphans');
    const klassenPersonen = el(p + 'KlassenPersonen');
    const lehrerPersonen = el(p + 'LehrerPersonen');
    return {
        syncMode: !alwaysNew || !alwaysNew.checked,
        removeOrphans: !removeOrphans || removeOrphans.checked,
        klassenPersonen: !klassenPersonen || klassenPersonen.checked,
        lehrerPersonen: !lehrerPersonen || lehrerPersonen.checked
    };
}

function fillSiteUrl(p) {
    const urlEl = el(p + 'SiteUrl');
    if (!urlEl) return;
    if (String(urlEl.value || '').trim()) return;
    const saved = readIntranetSiteUrl();
    if (saved) urlEl.value = saved;
}

function renderPanel(mount) {
    const p = mount.getAttribute('data-mount-prefix') || 'tsdSpo';
    const compact = mount.getAttribute('data-compact') === 'true';
    mount.dataset.intranetPrefix = p;
    mount.innerHTML =
        '<div class="stammdaten-intranet-listen" style="border:1px solid var(--border);border-radius:12px;padding:14px 16px;background:var(--card);">' +
        '<div style="display:flex;flex-wrap:wrap;gap:10px;align-items:flex-start;justify-content:space-between;margin-bottom:10px;">' +
        '<div><h3 style="margin:0 0 4px;font-size:1.05em;"><i class="bi bi-database"></i> Intranet-Listen aus Stammdaten</h3>' +
        '<p style="margin:0;color:var(--muted);font-size:0.9em;line-height:1.4;max-width:52ch;">Listen auf der Schulwebsite anlegen oder abgleichen – ohne separates Werkzeug. ' +
        '<a href="tools/sharepoint-liste-stammdaten.html">Erweitertes Tool</a></p></div></div>' +
        '<label style="display:block;margin-bottom:8px;"><span style="font-weight:700;font-size:0.9em;">SharePoint-Website</span>' +
        '<input type="url" id="' +
        p +
        'SiteUrl" placeholder="https://…/sites/Intranet" autocomplete="off" spellcheck="false" style="width:100%;margin-top:4px;"></label>' +
        '<div style="display:grid;gap:6px;margin:10px 0;font-size:0.88em;">' +
        checkbox(p, 'WantSchueler', 'Schülerinnen', true) +
        checkbox(p, 'WantFaecher', 'Fächer', true) +
        checkbox(p, 'WantFachgruppen', 'Fachgruppen', true) +
        checkbox(p, 'WantArge', 'ARGEs', true) +
        checkbox(p, 'WantKlassen', 'Klassen', true) +
        checkbox(p, 'WantLehrer', 'Lehrerinnen', true) +
        checkbox(p, 'KlassenPersonen', 'Klassen: Schüler als Personenfeld (M365)', true) +
        checkbox(p, 'LehrerPersonen', 'Lehrer: Personenfeld Lehrkraft (M365)', true) +
        '</div>' +
        (compact
            ? ''
            : '<details style="margin:8px 0;"><summary style="cursor:pointer;font-weight:700;">Listenname je Typ</summary>' +
              nameField(p, 'SchuelerName', 'Schülerinnen') +
              nameField(p, 'FaecherName', 'Fächer') +
              nameField(p, 'FachgruppenName', 'Fachgruppen') +
              nameField(p, 'ArgeName', 'ARGEs') +
              nameField(p, 'KlassenName', 'Klassen') +
              nameField(p, 'LehrerName', 'Lehrerinnen') +
              '</details>') +
        // Abgleich/Berechtigungen bewusst nicht im Register-Stammdaten-Tab; siehe docs/intranet-stammdaten-listen-notizen.md
        '<details style="margin:8px 0;"><summary style="cursor:pointer;font-weight:700;">Abgleich &amp; Berechtigungen</summary>' +
        '<div style="margin-top:8px;display:grid;gap:6px;font-size:0.88em;">' +
        checkbox(p, 'AlwaysNew', 'Immer neue Listen anlegen (statt Abgleich)', false) +
        checkbox(p, 'RemoveOrphans', 'Verwaiste Zeilen entfernen', true) +
        checkbox(p, 'SkipPerms', 'Berechtigungen überspringen', false) +
        '</div>' +
        groupPickerRow(p, 'GroupAdmin', 'Verwaltung / Admin') +
        groupPickerRow(p, 'GroupLehrer', 'Lehrkräfte') +
        groupPickerRow(p, 'GroupSchueler', 'Schüler (Sammelgruppe)') +
        '<p style="margin:8px 0 0;font-size:0.82em;color:var(--muted);">Rollen wie Schularbeiten-Planer: z. B. Klassen/Fächer für Schüler nur Lesen, Schülerinnen-Stammliste ohne Schüler-Gruppe.</p>' +
        '</details>' +
        '<div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:12px;">' +
        '<button type="button" class="btn btn-success" id="' +
        p +
        'BtnSync"><i class="bi bi-arrow-repeat"></i>Listen abgleichen</button>' +
        '<button type="button" class="btn" id="' +
        p +
        'BtnPerms"><i class="bi bi-shield-lock"></i>Nur Berechtigungen</button>' +
        '</div>' +
        '<pre id="' +
        p +
        'Log" class="tm-log" style="margin-top:10px;max-height:160px;overflow:auto;font-size:0.8em;" aria-live="polite"></pre>' +
        '</div>';

    if (compact) {
        INTRANET_LIST_KINDS.forEach(function (kind) {
            const suffix = INTRANET_LIST_TITLE_SUFFIX[kind];
            const hidden = document.createElement('input');
            hidden.type = 'hidden';
            hidden.id = p + suffix;
            hidden.value = resolveIntranetListTitle(kind, storedIntranetListTitle(kind));
            mount.querySelector('.stammdaten-intranet-listen').appendChild(hidden);
        });
    } else {
        INTRANET_LIST_KINDS.forEach(function (kind) {
            const suffix = INTRANET_LIST_TITLE_SUFFIX[kind];
            const mountEl = suffix ? el(p + suffix) : null;
            if (!mountEl) return;
            const stored = storedIntranetListTitle(kind);
            if (stored && !String(mountEl.value || '').trim()) mountEl.value = stored;
        });
    }

    fillSiteUrl(p);
    syncIntranetListTitlesToMounts();
    const urlEl = el(p + 'SiteUrl');
    if (urlEl && urlEl.dataset.intranetUrlBound !== '1') {
        urlEl.dataset.intranetUrlBound = '1';
        urlEl.addEventListener('change', function () {
            writeIntranetSiteUrl(String(urlEl.value || '').trim());
        });
    }
    initEmbeddedPermissionsUi(p, p + 'SkipPerms');

    const syncBtn = el(p + 'BtnSync');
    if (syncBtn) {
        syncBtn.addEventListener('click', function () {
            runSync(p).catch(function (e) {
                toast(e && e.message ? e.message : String(e));
            });
        });
    }
    const permsBtn = el(p + 'BtnPerms');
    if (permsBtn) {
        permsBtn.addEventListener('click', function () {
            runPermsOnly(p).catch(function (e) {
                toast(e && e.message ? e.message : String(e));
            });
        });
    }
}

function checkbox(p, id, label, checked) {
    return (
        '<label style="display:flex;gap:8px;align-items:center;"><input type="checkbox" id="' +
        p +
        id +
        '"' +
        (checked ? ' checked' : '') +
        '> ' +
        label +
        '</label>'
    );
}

function nameField(p, id, label) {
    const defaults = {
        SchuelerName: 'Schülerinnen',
        FaecherName: 'Fächer',
        FachgruppenName: 'Fachgruppen',
        ArgeName: 'ARGEs',
        KlassenName: 'Klassen',
        LehrerName: 'Lehrerinnen'
    };
    return (
        '<label style="display:block;margin-top:6px;"><span style="font-size:0.85em;">' +
        label +
        '</span><input type="text" id="' +
        p +
        id +
        '" value="' +
        (defaults[id] || '') +
        '" maxlength="200" style="width:100%;margin-top:2px;"></label>'
    );
}

function groupPickerRow(p, id, label) {
    return (
        '<div style="margin-top:10px;"><label style="font-size:0.85em;font-weight:700;">' +
        label +
        '</label>' +
        '<div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:4px;">' +
        '<input type="text" id="' +
        p +
        id +
        '" readonly placeholder="Gruppe wählen …" style="flex:1;min-width:180px;">' +
        '<input type="hidden" id="' +
        p +
        id +
        'Id">' +
        '<button type="button" class="btn btn-sm" id="' +
        p +
        id +
        'Pick"><i class="bi bi-search"></i></button>' +
        '<button type="button" class="btn btn-sm alt" id="' +
        p +
        id +
        'Clear"><i class="bi bi-x-lg"></i></button>' +
        '</div></div>'
    );
}

async function applyPerms(webUrl, p, listOpts, write) {
    const defs = buildStammdatenGroupFields(p);
    const hasPickers = defs.some(function (f) {
        return el(f.labelInputId);
    });
    let perms;
    if (hasPickers) {
        perms = readPermissionsFromPickers(defs, p + 'SkipPerms');
        persistPickersToStorage(defs, p + 'SkipPerms');
    } else {
        perms = loadPermissionsConfig();
        const skipEl = el(p + 'SkipPerms');
        if (skipEl && skipEl.checked) perms.skipPerms = true;
    }
    return await applyStammdatenPackagePermissions(webUrl, perms, write, listOpts);
}

export async function runQuickIntranetSync(kind, prefix) {
    const k = String(kind || '').trim();
    if (!INTRANET_KIND_LABELS[k]) {
        throw new Error('Unbekannter Listentyp: ' + k);
    }
    const p = prefix || 'tsdSpo';
    const webUrl = readIntranetSiteUrl();
    if (!webUrl) {
        toast('Bitte unter Stammdaten die SharePoint-Webseite für Intranet-Listen eintragen.');
        return;
    }
    const listOpts = buildListOptsForKinds([k], p);
    const listTitle = listOpts[k + 'Title'] || DEFAULT_LIST_TITLES[k];
    const ok = await confirmAsync(
        'SharePoint-Liste „' +
            listTitle +
            '“ aus den lokalen Stammdaten abgleichen?\n\n' +
            webUrl +
            '\n\nZuerst „Speichern“, wenn Sie gerade Änderungen gemacht haben.'
    );
    if (!ok) return;

    const write = function (msg) {
        logTo(p, msg);
    };
    const runOpts = readIntranetRunOpts(p);
    const results = await api().syncSelectedLists(webUrl, listOpts, write, runOpts);

    const skipPerms = el(p + 'SkipPerms') && el(p + 'SkipPerms').checked;
    if (!skipPerms) {
        try {
            await applyPerms(webUrl, p, listOpts, write);
        } catch (e) {
            write('Berechtigungen: ' + (e && e.message ? e.message : String(e)));
        }
    }

    writeIntranetSiteUrl(webUrl);

    const entry = results && results[k] ? results[k] : null;
    const listUrl = entry && entry.webUrl ? String(entry.webUrl).trim() : '';
    const count = entry && entry.count != null ? entry.count : null;
    return { kind: k, listTitle: listTitle, webUrl: listUrl, count: count, siteUrl: webUrl };
}

async function runSync(p) {
    const webUrl = readIntranetSiteUrl();
    if (!webUrl) {
        toast('Bitte die SharePoint-Website eintragen (Stammdaten oder hier).');
        return;
    }
    const listOpts = collectListOpts(p);
    if (
        !listOpts.schueler &&
        !listOpts.faecher &&
        !listOpts.fachgruppen &&
        !listOpts.arges &&
        !listOpts.klassen &&
        !listOpts.lehrer
    ) {
        toast('Mindestens eine Liste auswählen.');
        return;
    }
    const ok = await confirmAsync(
        'Ausgewählte Listen auf der Website abgleichen?\n\n' + webUrl + '\n\nDaten aus den lokalen Stammdaten in diesem Browser.'
    );
    if (!ok) return;

    const logEl = el(p + 'Log');
    if (logEl) logEl.textContent = '';
    const write = function (msg) {
        logTo(p, msg);
    };
    const runOpts = readRunOpts(p);

    const results = await api().syncSelectedLists(webUrl, listOpts, write, runOpts);

    const skipPerms = el(p + 'SkipPerms') && el(p + 'SkipPerms').checked;
    if (!skipPerms) {
        try {
            await applyPerms(webUrl, p, listOpts, write);
        } catch (e) {
            write('Berechtigungen: ' + (e && e.message ? e.message : String(e)));
        }
    } else {
        write('Berechtigungen übersprungen.');
    }

    writeIntranetSiteUrl(webUrl);

    let synced = 0;
    let totalRows = 0;
    Object.keys(results || {}).forEach(function (kind) {
        const row = results[kind];
        if (!row) return;
        synced += 1;
        if (row.count != null) totalRows += Number(row.count) || 0;
        if (row.webUrl) {
            const titleKey = kind + 'Title';
            saveIntranetListLink(kind, {
                url: row.webUrl,
                title: listOpts[titleKey] || DEFAULT_LIST_TITLES[kind] || kind,
                count: row.count
            });
        }
    });
    refreshIntranetListOpenLinks();
    if (synced === 1) {
        const onlyKind = Object.keys(results)[0];
        const only = results[onlyKind];
        showIntranetSyncToast(
            listOpts[onlyKind + 'Title'] || DEFAULT_LIST_TITLES[onlyKind],
            only && only.count,
            !!(only && only.webUrl)
        );
    } else {
        showIntranetSyncToast(
            synced + ' Listen',
            totalRows || null,
            Object.keys(results || {}).some(function (k) {
                return results[k] && results[k].webUrl;
            })
        );
    }
}

async function runPermsOnly(p) {
    const webUrl = readIntranetSiteUrl();
    if (!webUrl) {
        toast('Bitte die SharePoint-Website eintragen.');
        return;
    }
    const logEl = el(p + 'Log');
    if (logEl) logEl.textContent = '';
    await applyPerms(webUrl, p, collectListOpts(p), function (msg) {
        logTo(p, msg);
    });
    toast('Berechtigungen angewendet.');
}

function boot() {
    document.querySelectorAll('[data-ms365-intranet-listen-mount]').forEach(function (mount) {
        if (mount.dataset.intranetMounted === '1') return;
        mount.dataset.intranetMounted = '1';
        renderPanel(mount);
    });
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();

window.ms365StammdatenIntranetListenUi = {
    runSync: runSync,
    runPermsOnly: runPermsOnly,
    runQuickIntranetSync: runQuickIntranetSync,
    readIntranetSiteUrl: readIntranetSiteUrl,
    writeIntranetSiteUrl: writeIntranetSiteUrl
};
