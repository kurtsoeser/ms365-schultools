/**
 * SharePoint: sRDP-Anmeldeliste anlegen (Spalten, Ansichten, Rechte).
 */
import {
    STORAGE_KEY,
    PROFILE_IDS,
    getProfile,
    listTitleForYear,
    listTitleCandidates,
    buildAnmeldungColumns,
    toGraphColumnBody,
    defaultViewsForProfile,
    requiredColumnNamesForProfile
} from '../srdp-anmeldung/srdp-anmeldung-schema.js';
import {
    filterAbschlussklassen,
    suggestLfsChoices,
    teacherChoiceLabels,
    parseChoiceLines,
    buildItemTitle,
    buildPruefplanKurz
} from '../srdp-anmeldung/srdp-anmeldung-logic.js';
import {
    periodenForTerminJahr,
    formatGermanDate,
    refreshPruefungstermineFromWeb,
    bfkIsoFromColumn,
    normalizeIsoDate
} from '../srdp-anmeldung/srdp-pruefungstermine.js';
import {
    buildFieldFormatSpecs,
    buildClientFormCustomFormatterString
} from '../srdp-anmeldung/srdp-anmeldung-formatting.js';
import {
    entraGroupLogonName,
    isBroadSiteAudience,
    SPO_ROLE
} from '../../shared/stammdaten-sharepoint-sync-logic.js';
import {
    mountSrdpEntraPermHost,
    readSrdpPermissionsFromPickers,
    persistSrdpPickersToStorage,
    loadEffectiveSrdpPermissionsConfig,
    normalizeSrdpPermissionsConfig,
    formatSrdpAudienceSummary,
    saveSrdpPermissionsConfig
} from './sharepoint-liste-srdp-permissions-ui.js';
import { patchSchoolToolSitesSetup } from '../../shared/school-tool-sites.js';
import { defaultStammdatenSpoListTitles, lookupClassM365GroupFromAppData } from '../../shared/school-stammdaten-spo-lists.js';
import { pickEntraGroupsMulti } from '../../shared/entra-group-picker.js';
import { abschlussarbeitenListTitleForYear } from '../srdp-anmeldung/srdp-abschlussarbeiten-schema.js';
import {
    ensureAbschlussarbeitenList,
    loadAbschlussarbeiten,
    replaceAbschlussarbeiten
} from '../srdp-anmeldung/srdp-abschlussarbeiten-graph.js';
import { parseAbaPortalRows } from '../srdp-anmeldung/srdp-abschlussarbeiten-parse.js';

const SCOPES_GRAPH = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/Group.Read.All',
    'https://graph.microsoft.com/User.Read.All'
];

const SRS_WIZARD_STEP_COUNT = 6;
const SRS_WIZARD_STEP_STORAGE_KEY = 'ms365-srdp-wizard-step-v2';
const SRS_WIZARD_STEP_LABELS = [
    'Website & Jahr',
    'Klassen',
    'Optionen',
    'Abschlussarbeiten',
    'Rechte',
    'Anlegen'
];
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function G() {
    const api = window.ms365SpoGraph;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

function $(id) {
    return document.getElementById(id);
}

function log(msg) {
    const el = $('srsLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + msg;
    el.scrollTop = el.scrollHeight;
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function loadWizardState() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return {};
        const o = JSON.parse(raw);
        return o && typeof o === 'object' ? o : {};
    } catch {
        return {};
    }
}

function saveWizardState(partial) {
    const next = Object.assign({}, loadWizardState(), partial || {});
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

function loadStammdaten() {
    if (typeof window.ms365TenantSettingsLoad === 'function') {
        try {
            return window.ms365TenantSettingsLoad() || {};
        } catch {
            /* ignore */
        }
    }
    return {};
}

async function ensureToken(scopes) {
    return await G().getGraphToken(scopes || SCOPES_GRAPH);
}

function listNameKey(s) {
    return String(s || '')
        .trim()
        .toLowerCase()
        .replace(/[\s_\-]+/g, '');
}

async function findListByTitle(token, siteId, listTitle) {
    const title = String(listTitle || '').trim();
    if (!title) return null;
    const want = title.toLowerCase();
    const wantKey = listNameKey(title);

    try {
        const path =
            G().graphPathSite(siteId) +
            '/lists?$filter=' +
            encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
            '&$select=id,displayName,webUrl,name';
        const data = await G().graphJson('GET', path, token, undefined, 'v1.0');
        const hit = ((data && data.value) || [])[0] || null;
        if (hit && hit.id) return hit;
    } catch {
        /* Filter oft leer – Fallbacks */
    }

    try {
        const direct = await G().graphJson(
            'GET',
            G().graphPathSite(siteId) +
                '/lists/' +
                encodeURIComponent(title) +
                '?$select=id,displayName,webUrl,name',
            token,
            undefined,
            'v1.0'
        );
        if (direct && direct.id) return direct;
    } catch {
        /* DisplayName-Pfad */
    }

    let path = G().graphPathSite(siteId) + '/lists?$select=id,displayName,webUrl,name&$top=200';
    while (path) {
        const data = await G().graphJson(
            'GET',
            path.indexOf('http') === 0 ? path : path,
            token,
            undefined,
            'v1.0'
        );
        const batch = (data && data.value) || [];
        for (let i = 0; i < batch.length; i++) {
            const row = batch[i];
            const dn = String(row.displayName || '').trim().toLowerCase();
            const nm = String(row.name || '').trim().toLowerCase();
            if (dn === want || nm === want) return row;
            if (listNameKey(row.displayName) === wantKey || listNameKey(row.name) === wantKey) {
                return row;
            }
        }
        path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
    }
    return null;
}

async function addMissingColumns(siteId, listId, token, defs, write) {
    const colsPath = G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns?$top=200';
    const colsData = await G().graphJson('GET', colsPath, token, undefined, 'v1.0');
    const existing = new Set(
        ((colsData && colsData.value) || [])
            .map((c) => String((c && c.name) || '').trim())
            .filter(Boolean)
    );
    const base = G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
    let added = 0;
    for (let i = 0; i < defs.length; i++) {
        const def = defs[i];
        if (existing.has(def.name)) continue;
        await G().graphJson('POST', base, token, toGraphColumnBody(def), 'v1.0');
        added++;
        write('  + Spalte ' + def.name);
        await G().sleep(120);
    }
    return added;
}

function sameChoiceList(a, b) {
    const left = (Array.isArray(a) ? a : []).map((s) => String(s || '').trim()).filter(Boolean);
    const right = (Array.isArray(b) ? b : []).map((s) => String(s || '').trim()).filter(Boolean);
    if (left.length !== right.length) return false;
    return left.every((s, i) => s === right[i]);
}

/** Bestehende Choice-Spalten auf die Werte aus Schritt 3 setzen. */
async function syncExistingChoiceColumns(siteId, listId, token, defs, write) {
    const colsPath =
        G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns?$top=200';
    const colsData = await G().graphJson('GET', colsPath, token, undefined, 'v1.0');
    const byName = new Map();
    ((colsData && colsData.value) || []).forEach((col) => {
        const name = String((col && col.name) || '').trim();
        if (name) byName.set(name, col);
    });
    const base = G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns/';
    let updated = 0;
    for (let i = 0; i < defs.length; i++) {
        const def = defs[i];
        const choices =
            def && def.choice && Array.isArray(def.choice.choices)
                ? def.choice.choices.map((s) => String(s || '').trim()).filter(Boolean)
                : [];
        if (!def || !def.name || !choices.length) continue;
        const col = byName.get(def.name);
        if (!col || !col.id) continue;
        const current = col.choice && Array.isArray(col.choice.choices) ? col.choice.choices : [];
        if (sameChoiceList(current, choices)) continue;
        try {
            await G().graphJson(
                'PATCH',
                base + encodeURIComponent(col.id),
                token,
                {
                    choice: {
                        allowTextEntry: !!(def.choice && def.choice.allowTextEntry),
                        choices
                    }
                },
                'v1.0'
            );
            updated += 1;
            write('  · Optionen geschrieben: ' + def.displayName);
            await G().sleep(120);
        } catch (e) {
            write('  ! Optionen ' + def.displayName + ': ' + (e && e.message ? e.message : e));
        }
    }
    return updated;
}

async function resolveGroupId(token, mailOrId) {
    const raw = String(mailOrId || '').trim();
    if (!raw) return '';
    if (/^[0-9a-f-]{36}$/i.test(raw)) return raw;
    const esc = raw.replace(/'/g, "''");
    const filter = encodeURIComponent(
        "mail eq '" + esc + "' or mailNickname eq '" + esc + "' or displayName eq '" + esc + "'"
    );
    const data = await G().graphJson(
        'GET',
        '/groups?$filter=' + filter + '&$select=id,displayName,mail,mailNickname&$top=5',
        token,
        undefined,
        'v1.0'
    );
    const g = ((data && data.value) || [])[0];
    return g && g.id ? String(g.id) : '';
}

/**
 * @param {string} webUrl
 * @param {{
 *   profileId?: string,
 *   terminJahr: string|number,
 *   klassen: string[],
 *   lehrer: string[],
 *   lfs: string[],
 *   wahlfaecher: string[],
 *   seminare: string[],
 *   groupAdmin?: string,
 *   groupAdminId?: string,
 *   groupLehrer?: string,
 *   groupLehrerId?: string,
 *   groupKandidaten?: string,
 *   groupKandidatenId?: string,
 *   adminGroups?: Array<{ groupId?: string, groupLabel?: string }>,
 *   lehrerGroups?: Array<{ groupId?: string, groupLabel?: string }>,
 *   kandidatenGroups?: Array<{ groupId?: string, groupLabel?: string }>,
 *   adminUsers?: Array<{ id?: string, displayName?: string, mail?: string }>,
 *   lehrerUsers?: Array<{ id?: string, displayName?: string, mail?: string }>,
 *   kandidatenUsers?: Array<{ id?: string, displayName?: string, mail?: string }>,
 *   applyPermissions?: boolean
 * }} opts
 * @param {(msg: string) => void} [logFn]
 */
async function createSrdpList(webUrl, opts, logFn) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');
    const o = opts || {};
    const profile = getProfile(o.profileId);
    const terminJahr = String(o.terminJahr || '').trim();
    const title = listTitleForYear(terminJahr, profile.id);

    const token = await ensureToken();
    write('Löse Website auf …');
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
    write('Site: ' + (site.displayName || siteId));
    write('Schulform: ' + profile.label + ' · Liste: „' + title + '"');

    const columnDefs = buildAnmeldungColumns({
        profileId: profile.id,
        klassen: o.klassen,
        lehrer: o.lehrer,
        lfs: o.lfs,
        wahlfaecher: o.wahlfaecher,
        seminare: o.seminare,
        terminJahr
    });

    let list = await findListByTitleCandidates(token, siteId, terminJahr, profile.id);
    if (!list || !list.id) {
        write('Erstelle Liste …');
        const created = await G().graphJson(
            'POST',
            G().graphPathSite(siteId) + '/lists',
            token,
            {
                displayName: title,
                description:
                    'Anmeldung ' +
                    profile.examShort +
                    ' ' +
                    profile.label +
                    ' Haupttermin ' +
                    terminJahr +
                    '. Formular mit Varianten-Prüfplan; Anzeige-Name aus Nachname/Vorname.',
                list: { template: 'genericList' }
            },
            'v1.0'
        );
        const listId = created && created.id ? String(created.id) : '';
        if (!listId) throw new Error('Listen-ID fehlt.');
        list = { id: listId, displayName: title, webUrl: created.webUrl || '' };
        write('Liste angelegt: ' + (list.webUrl || listId));
    } else {
        write('Liste bereits vorhanden („' + (list.displayName || title) + '") – ergänze fehlende Spalten.');
    }

    write('Prüfe Spalten …');
    const added = await addMissingColumns(siteId, list.id, token, columnDefs, write);
    write(added ? added + ' Spalte(n) ergänzt.' : 'Spalten vollständig.');
    write('Schreibe LFS, Wahlfach und Seminar in die Liste …');
    const choiceUpdates = await syncExistingChoiceColumns(siteId, list.id, token, columnDefs, write);
    write(
        choiceUpdates
            ? choiceUpdates + ' Auswahlfeld(er) aktualisiert.'
            : 'Auswahlfelder sind bereits auf diesem Stand.'
    );
    if (normalizeIsoDate(o.bfkDatum)) {
        write('Schreibe BFK schriftlich …');
        await upsertBfkColumn(siteId, list.id, token, o.bfkDatum, write);
    }

    let host = '';
    try {
        host = new URL(url).hostname;
    } catch {
        host = '';
    }
    if (!host) throw new Error('SharePoint-Host aus URL nicht lesbar.');

    write('SharePoint-Token (Ansichten / Rechte) …');
    const spoScope = 'https://' + host + '/AllSites.FullControl';
    let spoToken;
    try {
        spoToken = await G().getGraphToken([spoScope]);
    } catch (e) {
        throw new Error(
            'SharePoint-Token fehlgeschlagen (AllSites.FullControl?). ' +
                (e && e.message ? e.message : e)
        );
    }
    const digest = await G().getSpoRequestDigest(url, spoToken);
    const listDisplayName = list.displayName || title;

    try {
        await G().spoPatchTitleField(url, spoToken, digest, listDisplayName, {
            displayName: 'Name (aus Nachname, Vorname)',
            description:
                'Wird in Ansichten aus Nachname und Vorname angezeigt; im Formular ausgeblendet. Optional „Titel nachziehen“ im Wizard.',
            required: false
        });
        write('Title-Feld: optional, im Formular ausgeblendet.');
    } catch (e) {
        write('Hinweis Title-Feld: ' + (e && e.message ? e.message : e));
    }

    write('Formular & Spaltenformatierung …');
    try {
        await applyFormAndColumnFormats(url, spoToken, digest, listDisplayName, terminJahr, profile.id, write);
    } catch (e) {
        write('Hinweis Formatierung: ' + (e && e.message ? e.message : e));
    }

    write('Ansichten …');
    const views = defaultViewsForProfile(profile.id);
    for (let i = 0; i < views.length; i++) {
        const v = views[i];
        try {
            const r = await G().spoEnsureListView(url, spoToken, digest, listDisplayName, v);
            write((r.created ? '  + Ansicht ' : '  · vorhanden ') + v.title);
        } catch (e) {
            write('  ! Ansicht „' + v.title + '": ' + (e && e.message ? e.message : e));
        }
        await G().sleep(150);
    }

    if (o.applyPermissions !== false) {
        write('Berechtigungen …');
        try {
            await applyListPermissions(url, spoToken, digest, listDisplayName, token, o, write);
        } catch (e) {
            write('Hinweis Rechte: ' + (e && e.message ? e.message : e));
        }
    }

    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
        window.ms365ActionLog.append({
            tool: 'sharepoint',
            action: 'create-srdp-list',
            target: url,
            summary: title + ' (' + profile.label + ')'
        });
    }

    write('Fertig. Prüfplan-Beispiel Variante 1: ' + buildPruefplanKurz(1, profile.id));
    return { siteId, list, title, webUrl: list.webUrl || '', profileId: profile.id };
}

async function findListByTitleCandidates(token, siteId, terminJahr, profileId) {
    const titles = listTitleCandidates(terminJahr, profileId);
    for (let i = 0; i < titles.length; i++) {
        const hit = await findListByTitle(token, siteId, titles[i]);
        if (hit && hit.id) return hit;
    }
    return null;
}

/**
 * Column Formatting, Defaults, Formular-Sichtbarkeit, ClientFormCustomFormatter.
 * ConditionalShowFormula wird nicht im Field-MERGE mitgeschickt (SPO 400 auf typed Fields);
 * Sichtbarkeit läuft über spoSetFieldConditionalShowFormula / setshowin*form.
 */
async function applyFormAndColumnFormats(siteWebUrl, spoToken, digest, listTitle, terminJahr, profileId, write) {
    const specs = buildFieldFormatSpecs(terminJahr, profileId);
    for (let i = 0; i < specs.length; i++) {
        const s = specs[i];
        /** @type {Record<string, unknown>} */
        const body = {};
        if (s.customFormatter) body.CustomFormatter = JSON.stringify(s.customFormatter);
        if (s.required === false) body.Required = false;
        if (s.required === true) body.Required = true;
        if (s.displayName) body.Title = s.displayName;
        if (s.defaultValue != null && s.defaultValue !== '') body.DefaultValue = String(s.defaultValue);

        let fieldOk = true;
        if (Object.keys(body).length) {
            try {
                await G().spoPatchListField(siteWebUrl, spoToken, digest, listTitle, s.internalName, body);
            } catch (e) {
                fieldOk = false;
                write('  ! Feld ' + s.internalName + ': ' + (e && e.message ? e.message : e));
            }
        }

        if (s.conditionalShowFormula != null && typeof G().spoSetFieldConditionalShowFormula === 'function') {
            try {
                const r = await G().spoSetFieldConditionalShowFormula(
                    siteWebUrl,
                    spoToken,
                    digest,
                    listTitle,
                    s.internalName,
                    s.conditionalShowFormula
                );
                const via = r && r.via ? String(r.via) : '';
                write(
                    '  · Feld ' +
                        s.internalName +
                        (via === 'show-in-forms'
                            ? ' (im Formular ausgeblendet)'
                            : via === 'verbose-sp-field'
                              ? ' (bedingte Formel)'
                              : '')
                );
            } catch (e) {
                // Bedingte Formeln (z. B. Seminar) sind per REST oft nicht setzbar –
                // Spaltenformat/Defaults sollen trotzdem greifen; Soft-Hinweis reicht.
                const raw = String((e && e.message) || e || '');
                const short =
                    raw.length > 160 ? raw.slice(0, 157).replace(/\s+\S*$/, '') + '…' : raw;
                write(
                    '  · Feld ' +
                        s.internalName +
                        (fieldOk ? '' : ' (Patch fehlerhaft)') +
                        ' – bedingte Sichtbarkeit übersprungen (' +
                        short +
                        '). Native Liste zeigt das Feld ggf. immer; App-Wizard blendet korrekt aus.'
                );
            }
        } else if (fieldOk && Object.keys(body).length) {
            write('  · Feld ' + s.internalName);
        }
        await G().sleep(100);
    }
    try {
        const fmt = buildClientFormCustomFormatterString(profileId);
        await G().spoSetListClientFormCustomFormatter(siteWebUrl, spoToken, digest, listTitle, fmt);
        write('  · Formular-Header/Sektionen gesetzt');
    } catch (e) {
        write('  ! Formular-Layout: ' + (e && e.message ? e.message : e));
    }
}

/**
 * Title aus Nachname/Vorname nachziehen (bestehende Elemente).
 * @param {string} webUrl
 * @param {string|number} terminJahr
 * @param {(msg: string) => void} [logFn]
 * @param {string} [profileId]
 */
async function backfillItemTitles(webUrl, terminJahr, logFn, profileId) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    const profile = getProfile(profileId);
    const title = listTitleForYear(terminJahr, profile.id);
    const token = await ensureToken();
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const list = await findListByTitleCandidates(token, siteId, terminJahr, profile.id);
    if (!list || !list.id) throw new Error('Liste „' + title + '" nicht gefunden.');

    let patched = 0;
    let skipped = 0;
    let next =
        G().graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(list.id) +
        '/items?$expand=fields&$top=50';
    while (next) {
        const data = await G().graphJson('GET', next, token, undefined, 'v1.0');
        const items = (data && data.value) || [];
        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            const fields = (item && item.fields) || {};
            const want = buildItemTitle(fields.Nachname, fields.Vorname);
            if (!want) {
                skipped++;
                continue;
            }
            const cur = String(fields.Title || '').trim();
            const plan = buildPruefplanKurz(fields.Variante, profile.id);
            if (cur === want && String(fields.PruefplanKurz || '') === plan) {
                skipped++;
                continue;
            }
            await G().graphJson(
                'PATCH',
                G().graphPathSite(siteId) +
                    '/lists/' +
                    encodeURIComponent(list.id) +
                    '/items/' +
                    encodeURIComponent(item.id) +
                    '/fields',
                token,
                { Title: want, PruefplanKurz: plan },
                'v1.0'
            );
            patched++;
            await G().sleep(80);
        }
        next = (data && data['@odata.nextLink']) || '';
    }
    write('Titel nachgezogen: ' + patched + ' geändert, ' + skipped + ' übersprungen.');
    return { patched, skipped };
}

async function applyListPermissions(siteWebUrl, spoToken, digest, listTitle, graphToken, opts, write) {
    await G().spoBreakListInheritance(siteWebUrl, spoToken, digest, listTitle, true);
    write('  Vererbung gebrochen (bestehende Rollen kopiert).');

    const assignments = await G().spoListRoleAssignments(siteWebUrl, spoToken, digest, listTitle);
    let removed = 0;
    for (let i = 0; i < assignments.length; i++) {
        const a = assignments[i];
        const member = a.Member || a.member || {};
        if (!isBroadSiteAudience(member)) continue;
        const pid = member.Id != null ? member.Id : a.PrincipalId;
        try {
            await G().spoRemoveRoleAssignment(siteWebUrl, spoToken, digest, listTitle, pid);
            removed++;
            write('  − ' + (member.Title || pid));
        } catch (e) {
            write('  Hinweis Entfernen: ' + (e && e.message ? e.message : e));
        }
    }
    if (removed) write('  Breite Rollen entfernt: ' + removed);

    let grantListTitle = listTitle;

    async function grantGroup(rawOrId, roleId, label) {
        const raw = String(rawOrId || '').trim();
        if (!raw) return;
        let id = GUID_RE.test(raw) ? raw : '';
        if (!id) id = await resolveGroupId(graphToken, raw);
        if (!id) {
            write('  ! Gruppe nicht gefunden (' + label + '): ' + raw);
            return;
        }
        const principal = await G().spoEnsureUser(siteWebUrl, spoToken, digest, entraGroupLogonName(id));
        await G().spoAddRoleAssignment(siteWebUrl, spoToken, digest, grantListTitle, principal.id, roleId);
        write('  + ' + label + ' (Gruppe): ' + (principal.title || raw));
    }

    async function grantUser(mailOrUpn, roleId, label) {
        const mail = String(mailOrUpn || '').trim();
        if (!mail) return;
        const logon =
            typeof G().membershipLogonName === 'function' ? G().membershipLogonName(mail) : mail;
        if (!logon) {
            write('  ! Person ohne LoginName (' + label + '): ' + mail);
            return;
        }
        const principal = await G().spoEnsureUser(siteWebUrl, spoToken, digest, logon);
        await G().spoAddRoleAssignment(siteWebUrl, spoToken, digest, grantListTitle, principal.id, roleId);
        write('  + ' + label + ' (Person): ' + (principal.title || mail));
    }

    async function grantAudience(groups, legacyId, legacyLabel, users, roleId, label, targetList) {
        const prevList = grantListTitle;
        if (targetList) grantListTitle = targetList;
        const seen = new Set();
        const entries = Array.isArray(groups) ? groups.slice() : [];
        const lid = String(legacyId || '').trim();
        const ll = String(legacyLabel || '').trim();
        if (lid || ll) entries.unshift({ groupId: lid, groupLabel: ll });
        for (let gi = 0; gi < entries.length; gi++) {
            const g = entries[gi] || {};
            const gid = String(g.groupId || g.id || '').trim();
            const glabel = String(g.groupLabel || g.label || '').trim();
            const key = (gid || glabel).toLowerCase();
            if (!key || seen.has(key)) continue;
            seen.add(key);
            try {
                if (gid) await grantGroup(gid, roleId, label);
                else if (glabel) await grantGroup(glabel, roleId, label);
            } catch (e) {
                write(
                    '  ! Gruppe ' +
                        (glabel || gid || '?') +
                        ' (' +
                        label +
                        '): ' +
                        (e && e.message ? e.message : e)
                );
            }
        }
        const list = Array.isArray(users) ? users : [];
        for (let i = 0; i < list.length; i++) {
            const u = list[i] || {};
            const mail = String(u.mail || u.userPrincipalName || '').trim();
            try {
                await grantUser(mail, roleId, label);
            } catch (e) {
                write(
                    '  ! Person ' +
                        (mail || u.displayName || '?') +
                        ' (' +
                        label +
                        '): ' +
                        (e && e.message ? e.message : e)
                );
            }
        }
        grantListTitle = prevList;
    }

    await grantAudience(
        opts.adminGroups,
        opts.groupAdminId,
        opts.groupAdmin,
        opts.adminUsers,
        SPO_ROLE.fullControl,
        'Vollzugriff'
    );
    // Gestaltung (nicht nur Lesen): bei „nur eigene Elemente“ sehen Read/Contribute
    // sonst ausschließlich selbst erstellte Zeilen. Lesende Rechte brauchen Sicht auf alle Anmeldungen.
    await grantAudience(
        opts.lehrerGroups,
        opts.groupLehrerId,
        opts.groupLehrer,
        opts.lehrerUsers,
        SPO_ROLE.design,
        'Lesende'
    );
    await grantAudience(
        opts.kandidatenGroups,
        opts.groupKandidatenId,
        opts.groupKandidaten,
        opts.kandidatenUsers,
        SPO_ROLE.contribute,
        'Kandidaten'
    );

    const lehrerListTitle = defaultStammdatenSpoListTitles().lehrer;
    write('Lehrerliste „' + lehrerListTitle + '“: Lesen für Kandidat:innen …');
    await grantAudience(
        opts.kandidatenGroups,
        opts.groupKandidatenId,
        opts.groupKandidaten,
        opts.kandidatenUsers,
        SPO_ROLE.read,
        'Lehrerliste',
        lehrerListTitle
    );

    try {
        // ReadSecurity 2 + WriteSecurity 2: Kandidat:innen (Mitwirken) sehen/bearbeiten nur eigene Zeilen.
        // Vollzugriff und Lesende Rechte (Gestaltung) sehen weiterhin alle Anmeldungen.
        await G().spoSetListItemLevelPermissions(siteWebUrl, spoToken, digest, listTitle, 2, 2);
        write(
            '  Elementberechtigungen: Kandidat:innen nur eigene Anmeldung. Vollzugriff und Lesende Rechte sehen alle.'
        );
    } catch (e) {
        write('  Hinweis Elementrechte: ' + (e && e.message ? e.message : e));
    }
}

/**
 * Nur die Listenrechte neu schreiben (Liste, Spalten und Ansichten bleiben).
 * @param {string} webUrl
 * @param {object} opts
 * @param {(msg: string) => void} [logFn]
 */
async function reapplySrdpPermissions(webUrl, opts, logFn) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    const o = opts || {};
    if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');
    if (!/^\d{4}$/.test(String(o.terminJahr || '').trim())) {
        throw new Error('Terminjahr (YYYY) fehlt.');
    }
    const okAdmin = audienceHasGrant(o.adminGroups, o.groupAdminId, o.groupAdmin, o.adminUsers);
    const okLehrer = audienceHasGrant(o.lehrerGroups, o.groupLehrerId, o.groupLehrer, o.lehrerUsers);
    const okKand = audienceHasGrant(
        o.kandidatenGroups,
        o.groupKandidatenId,
        o.groupKandidaten,
        o.kandidatenUsers
    );
    if (!okAdmin || !okLehrer || !okKand) {
        throw new Error(
            'Für Rechte: je Stufe Gruppe oder Person(en) wählen (Vollzugriff, Lesende Rechte, Kandidat:innen).'
        );
    }
    const profile = getProfile(o.profileId);
    const title = listTitleForYear(o.terminJahr, profile.id);
    const token = await ensureToken();
    write('Löse Website auf …');
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const list = await findListByTitleCandidates(token, siteId, o.terminJahr, profile.id);
    if (!list || !list.id) {
        throw new Error('Liste „' + title + '“ fehlt. Bitte zuerst anlegen.');
    }
    let host = '';
    try {
        host = new URL(url).hostname;
    } catch {
        host = '';
    }
    if (!host) throw new Error('SharePoint-Host aus URL nicht lesbar.');
    write('SharePoint-Token (Rechte) …');
    const spoToken = await G().getGraphToken(['https://' + host + '/AllSites.FullControl']);
    const digest = await G().getSpoRequestDigest(url, spoToken);
    write('Schreibe Berechtigungen auf „' + (list.displayName || title) + '“ neu …');
    await applyListPermissions(url, spoToken, digest, list.displayName || title, token, o, write);
    write('Fertig. Rechte nachgezogen.');
    return { list, title };
}

/**
 * @param {string} webUrl
 * @param {string|number} terminJahr
 * @param {(msg: string) => void} [logFn]
 * @param {string} [profileId]
 */
async function probeSrdpHealth(webUrl, terminJahr, logFn, profileId) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    const profile = getProfile(profileId);
    const title = listTitleForYear(terminJahr, profile.id);
    const required = requiredColumnNamesForProfile(profile.id);
    const token = await ensureToken();
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    const list = await findListByTitleCandidates(token, siteId, terminJahr, profile.id);
    /** @type {string[]} */
    const notes = [];
    let missing = required.slice();
    let listOk = false;
    let listWebUrl = '';
    let listTitle = title;
    if (!list || !list.id) {
        write('Fehlt: Anmeldeliste „' + title + '"');
        notes.push('Anmeldeliste fehlt');
    } else {
        listOk = true;
        listWebUrl = list.webUrl || '';
        listTitle = list.displayName || title;
        const colsPath =
            G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(list.id) + '/columns?$top=200';
        const colsData = await G().graphJson('GET', colsPath, token, undefined, 'v1.0');
        const colNames = ((colsData && colsData.value) || [])
            .map((c) => String((c && c.name) || '').trim())
            .filter(Boolean);
        missing = required.filter((n) => colNames.indexOf(n) === -1);
        if (missing.length) {
            write('Anmeldeliste – fehlende Spalten: ' + missing.join(', '));
            notes.push('Anmeldeliste: Spalten fehlen');
        } else {
            write('Anmeldeliste OK · ' + (list.webUrl || list.displayName || title));
        }
    }

    let abaOk = false;
    let abaTitle = '';
    let abaCount = 0;
    let abaWebUrl = '';
    try {
        abaTitle = abschlussarbeitenListTitleForYear(terminJahr, profile.id);
        const abaCtx = await ensureAbschlussarbeitenList(url, {
            jahr: terminJahr,
            profileId: profile.id
        });
        const abaRows = await loadAbschlussarbeiten(abaCtx);
        abaOk = true;
        abaCount = abaRows.length;
        abaWebUrl = (abaCtx.list && abaCtx.list.webUrl) || '';
        abaTitle = (abaCtx.list && abaCtx.list.title) || abaTitle;
        write(
            'Abschlussarbeiten OK · „' +
                abaTitle +
                '“ · ' +
                abaCount +
                ' Einträge' +
                (abaWebUrl ? ' · ' + abaWebUrl : '')
        );
        const abaLink = $('srsAbaListLink');
        if (abaLink && abaWebUrl) {
            abaLink.hidden = false;
            abaLink.href = abaWebUrl;
            abaLink.textContent = 'Abschlussarbeiten öffnen: ' + abaTitle;
        }
        setAbaStatus(
            'Liste „' + abaTitle + '“ bereit · ' + abaCount + ' Einträge (ABA-Excel optional hochladen).'
        );
    } catch (e) {
        write('Abschlussarbeiten: ' + ((e && e.message) || e));
        notes.push('Abschlussarbeiten prüfen/anlegen');
        setAbaStatus('Abschlussarbeiten-Liste fehlt oder nicht lesbar – bitte anlegen.');
    }

    const ok = listOk && missing.length === 0 && abaOk;
    const panel = $('srsHealth');
    if (panel) {
        panel.hidden = false;
        panel.textContent = ok
            ? 'Status: OK – Anmeldung „' +
              listTitle +
              '“ und Abschlussarbeiten „' +
              abaTitle +
              '“ (' +
              abaCount +
              ') bereit.'
            : 'Status: ' + (notes.length ? notes.join(' · ') : 'bitte Protokoll prüfen') + '.';
        panel.classList.toggle('ok', ok);
        panel.classList.toggle('warn', !ok);
    }

    return {
        ok,
        title: listTitle,
        webUrl: listWebUrl,
        missingColumns: missing,
        missingList: !listOk,
        abaOk,
        abaTitle,
        abaCount,
        abaWebUrl,
        profileId: profile.id
    };
}

function setAbaStatus(msg) {
    const el = $('srsAbaStatus');
    if (el) el.textContent = String(msg || '');
}

function currentProfileId() {
    return String(($('srsProfile') && $('srsProfile').value) || 'hak').trim().toLowerCase() || 'hak';
}

function readPermissionsOpts() {
    const fromPickers = normalizeSrdpPermissionsConfig(readSrdpPermissionsFromPickers());
    const skip = !!($('srsSkipPerms') && $('srsSkipPerms').checked);
    return { ...fromPickers, skipPerms: skip };
}

function audienceHasGrant(groups, groupId, groupLabel, users) {
    const hasGroup =
        (Array.isArray(groups) && groups.some((g) => String((g && (g.groupId || g.id)) || '').trim())) ||
        String(groupId || '').trim() ||
        String(groupLabel || '').trim();
    return !!(hasGroup || (Array.isArray(users) && users.length));
}

function readFormOpts() {
    const terminJahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
    const profileId = currentProfileId();
    const klassen = Array.from(document.querySelectorAll('input[name="srsKlasse"]:checked')).map(
        (el) => el.value
    );
    const lfs = parseChoiceLines(($('srsLfs') && $('srsLfs').value) || '');
    const wahlfaecher = parseChoiceLines(($('srsWahlfach') && $('srsWahlfach').value) || '');
    const seminare = parseChoiceLines(($('srsSeminar') && $('srsSeminar').value) || '');
    const stammdaten = loadStammdaten();
    const lehrer = teacherChoiceLabels(stammdaten.teachers || []);
    const perms = readPermissionsOpts();
    return {
        profileId,
        terminJahr,
        klassen,
        lehrer,
        lfs,
        wahlfaecher,
        seminare,
        bfkDatum: normalizeIsoDate(($('srsBfk') && $('srsBfk').value) || ''),
        groupAdmin: perms.groupAdmin,
        groupAdminId: perms.groupAdminId,
        groupLehrer: perms.groupLehrer,
        groupLehrerId: perms.groupLehrerId,
        groupKandidaten: perms.groupKandidaten,
        groupKandidatenId: perms.groupKandidatenId,
        adminGroups: perms.adminGroups,
        lehrerGroups: perms.lehrerGroups,
        kandidatenGroups: perms.kandidatenGroups,
        adminUsers: perms.adminUsers,
        lehrerUsers: perms.lehrerUsers,
        kandidatenUsers: perms.kandidatenUsers,
        applyPermissions: !perms.skipPerms
    };
}

function persistForm() {
    const perms = readPermissionsOpts();
    persistSrdpPickersToStorage(perms.skipPerms);
    syncExtraSelectionFromDom();
    saveWizardState({
        siteUrl: String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim(),
        profileId: currentProfileId(),
        terminJahr: String(($('srsJahr') && $('srsJahr').value) || '').trim(),
        lfs: String(($('srsLfs') && $('srsLfs').value) || ''),
        wahlfach: String(($('srsWahlfach') && $('srsWahlfach').value) || ''),
        seminar: String(($('srsSeminar') && $('srsSeminar').value) || ''),
        lfsText: String(($('srsLfs') && $('srsLfs').value) || ''),
        wahlfachText: String(($('srsWahlfach') && $('srsWahlfach').value) || ''),
        seminarText: String(($('srsSeminar') && $('srsSeminar').value) || ''),
        bfkDatum: normalizeIsoDate(($('srsBfk') && $('srsBfk').value) || ''),
        skipPerms: perms.skipPerms,
        klassenSelected: checkedStammdatenKlassen(),
        extraKlassen: extraKlassen.slice()
    });
}

/** @type {Array<{ code: string, displayName: string, groupId: string, mail: string, selected: boolean }>} */
let extraKlassen = [];

function klasseKey(value) {
    return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '');
}

function normalizeExtraKlasse(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const displayName = String(raw.displayName || raw.code || '').trim();
    const code = String(raw.code || displayName).trim();
    if (!code) return null;
    return {
        code,
        displayName: displayName || code,
        groupId: String(raw.groupId || '').trim(),
        mail: String(raw.mail || '').trim(),
        selected: raw.selected !== false
    };
}

function abschlussClassRows(classes, jahr) {
    const codes = filterAbschlussklassen(classes, jahr);
    const want = new Set(codes.map((c) => klasseKey(c)));
    return (Array.isArray(classes) ? classes : []).filter((c) => {
        const code = klasseKey(c && (c.code || c.name));
        return code && want.has(code);
    });
}

function studentsInClass(students, code) {
    const want = klasseKey(code);
    if (!want) return 0;
    return (Array.isArray(students) ? students : []).filter(
        (s) => klasseKey(s && (s.klasse || s.classCode || s.code)) === want
    ).length;
}

function kvLabel(cls, teachers) {
    const name = String((cls && (cls.headName || cls.klassenvorstandName)) || '').trim();
    if (name) return name;
    const em = String((cls && (cls.headEmail || cls.kvEmail || cls.classTeacherEmail)) || '')
        .trim()
        .toLowerCase();
    if (!em) return '';
    const hit = (Array.isArray(teachers) ? teachers : []).find(
        (t) => String((t && t.email) || '').trim().toLowerCase() === em
    );
    return hit && hit.name ? String(hit.name).trim() : em;
}

function checkedStammdatenKlassen() {
    return Array.from(document.querySelectorAll('input[name="srsKlasse"][data-srs-source="stamm"]'))
        .filter((el) => el.checked)
        .map((el) => el.value);
}

function syncExtraSelectionFromDom() {
    document.querySelectorAll('input[name="srsKlasse"][data-srs-source="m365"]').forEach((el) => {
        const id = el.getAttribute('data-srs-extra') || '';
        const row = extraKlassen.find((x) => (x.groupId && x.groupId === id) || x.code === el.value);
        if (row) row.selected = !!el.checked;
    });
}

function rememberedStammChecks() {
    const live = Array.from(document.querySelectorAll('input[name="srsKlasse"][data-srs-source="stamm"]'));
    if (live.length) return new Set(live.filter((el) => el.checked).map((el) => el.value));
    const state = loadWizardState();
    if (Array.isArray(state.klassenSelected)) return new Set(state.klassenSelected.map((c) => String(c)));
    return null;
}

function metaLine(parent, text) {
    const span = document.createElement('span');
    span.className = 'srs-class-card__meta';
    span.textContent = text;
    parent.appendChild(span);
}

function renderStammCard(cls, checked) {
    const code = String((cls && (cls.code || cls.name)) || '').trim();
    const label = document.createElement('label');
    label.className = 'srs-class-card';
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.name = 'srsKlasse';
    input.value = code;
    input.checked = checked;
    input.setAttribute('data-srs-source', 'stamm');
    const body = document.createElement('span');
    body.className = 'srs-class-card__body';
    const title = document.createElement('span');
    title.className = 'srs-class-card__code';
    const name = String((cls && cls.name) || '').trim();
    title.textContent = name && name !== code ? code + ' · ' + name : code;
    body.appendChild(title);
    const kv = kvLabel(cls, (loadStammdaten().teachers || []));
    metaLine(body, kv ? 'Klassenvorstand: ' + kv : 'Klassenvorstand: nicht hinterlegt');
    const count = studentsInClass(loadStammdaten().students || [], code);
    metaLine(
        body,
        count
            ? count + (count === 1 ? ' Schüler:in in den Stammdaten' : ' Schüler:innen in den Stammdaten')
            : 'Keine Schülerliste in den Stammdaten'
    );
    const group = lookupClassM365GroupFromAppData(code);
    if (group && (group.displayName || group.mail)) {
        metaLine(
            body,
            'Microsoft 365: ' + (group.displayName || group.mail) + (group.mail ? ' · ' + group.mail : '')
        );
    } else {
        metaLine(body, 'Microsoft 365: noch nicht verknüpft');
    }
    if (cls && cls.year) metaLine(body, 'Abschlussjahr ' + String(cls.year));
    label.appendChild(input);
    label.appendChild(body);
    return label;
}

function renderExtraCard(row) {
    const wrap = document.createElement('div');
    wrap.className = 'srs-class-card-row';
    const label = document.createElement('label');
    label.className = 'srs-class-card';
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.name = 'srsKlasse';
    input.value = row.code;
    input.checked = row.selected !== false;
    input.setAttribute('data-srs-source', 'm365');
    input.setAttribute('data-srs-extra', row.groupId || row.code);
    const body = document.createElement('span');
    body.className = 'srs-class-card__body';
    const title = document.createElement('span');
    title.className = 'srs-class-card__code';
    title.textContent = row.displayName || row.code;
    body.appendChild(title);
    metaLine(body, 'Zusätzliche Maturaklasse aus Microsoft 365');
    if (row.mail) metaLine(body, row.mail);
    if (row.code && row.code !== row.displayName) metaLine(body, 'Klassenkürzel: ' + row.code);
    label.appendChild(input);
    label.appendChild(body);
    const remove = document.createElement('button');
    remove.type = 'button';
    remove.className = 'btn';
    remove.textContent = 'Entfernen';
    remove.setAttribute('data-srs-remove-extra', row.groupId || row.code);
    wrap.appendChild(label);
    wrap.appendChild(remove);
    return wrap;
}

function renderExtraKlassen() {
    const host = $('srsExtraKlassenHost');
    if (!host) return;
    syncExtraSelectionFromDom();
    host.replaceChildren();
    if (!extraKlassen.length) {
        const p = document.createElement('p');
        p.className = 'muted';
        p.style.margin = '0';
        p.textContent = 'Noch keine zusätzliche Gruppe. Die Abschlussklassen oben reichen, wenn alle Maturaklassen dort stehen.';
        host.appendChild(p);
        return;
    }
    extraKlassen.forEach((row) => host.appendChild(renderExtraCard(row)));
}

function fillKlassenCheckboxes() {
    const host = $('srsKlassenHost');
    if (!host) return;
    const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
    const stammdaten = loadStammdaten();
    const remembered = rememberedStammChecks();
    host.replaceChildren();
    if (!/^\d{4}$/.test(jahr)) {
        const p = document.createElement('p');
        p.className = 'muted';
        p.style.margin = '0';
        p.textContent = 'Zuerst ein gültiges Terminjahr (YYYY) wählen.';
        host.appendChild(p);
        renderExtraKlassen();
        return;
    }
    const rows = abschlussClassRows(stammdaten.classes || [], jahr);
    if (!rows.length) {
        const p = document.createElement('p');
        p.className = 'muted';
        p.style.margin = '0';
        p.textContent =
            'Keine Klassen mit Abschlussjahr ' +
            jahr +
            ' in den Stammdaten. Bitte Klassen pflegen, das Jahr prüfen oder unten eine Microsoft-365-Gruppe ergänzen.';
        host.appendChild(p);
        renderExtraKlassen();
        return;
    }
    rows.forEach((cls) => {
        const code = String(cls.code || cls.name || '').trim();
        const checked = remembered ? remembered.has(code) : true;
        host.appendChild(renderStammCard(cls, checked));
    });
    renderExtraKlassen();
}

async function pickExtraKlassen() {
    const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
    const rows = abschlussClassRows(loadStammdaten().classes || [], jahr);
    const queries = rows.map((c) => String(c.code || c.name || '').trim()).filter((c) => c.length >= 2);
    const stammKeys = new Set(rows.map((c) => klasseKey(c.code || c.name)));
    let picked = [];
    try {
        picked = await pickEntraGroupsMulti({
            title: 'Weitere Maturaklassen',
            hint: 'Microsoft-365-Gruppen suchen und als zusätzliche Klassen übernehmen. Gruppen, die schon als Abschlussklasse hinterlegt sind, bleiben oben.',
            searchPlaceholder: 'Klassenname, Mail oder Alias …',
            initialQueries: queries.length ? queries.slice(0, 8) : ['5'],
            prefixFilter: false
        });
    } catch (e) {
        toast('Gruppen konnten nicht geladen werden: ' + (e && e.message ? e.message : e));
        return;
    }
    let added = 0;
    let skipped = 0;
    (Array.isArray(picked) ? picked : []).forEach((g) => {
        const displayName = String((g && g.displayName) || '').trim();
        const mail = String((g && (g.mail || g.mailNickname)) || '').trim();
        const code = displayName || mail;
        const groupId = String((g && g.id) || '').trim();
        if (!code) return;
        if (stammKeys.has(klasseKey(code)) || stammKeys.has(klasseKey(g && g.mailNickname))) {
            skipped += 1;
            return;
        }
        const exists = extraKlassen.some(
            (x) => (groupId && x.groupId === groupId) || klasseKey(x.code) === klasseKey(code)
        );
        if (exists) return;
        extraKlassen.push({
            code,
            displayName: displayName || code,
            groupId,
            mail,
            selected: true
        });
        added += 1;
    });
    renderExtraKlassen();
    persistForm();
    if (added) toast(added + (added === 1 ? ' Klasse ergänzt.' : ' Klassen ergänzt.'));
    else if (skipped) toast('Diese Gruppen sind schon als Abschlussklassen dabei.');
}

function applyProfileDefaults(forceChoices) {
    const profile = getProfile(currentProfileId());
    const stammdaten = loadStammdaten();
    const note = $('srsProfileNote');
    if (note) note.textContent = profile.notes || '';
    const lfsLabel = $('srsLfsLabel');
    if (lfsLabel) lfsLabel.textContent = (profile.lfsFieldLabel || 'LFS') + '-Optionen';
    if (forceChoices || ($('srsLfs') && !$('srsLfs').value.trim())) {
        if ($('srsLfs')) {
            $('srsLfs').value = suggestLfsChoices(stammdaten.subjects || [], profile.lfsHints).join('\n');
        }
    }
    if (forceChoices || ($('srsWahlfach') && !$('srsWahlfach').value.trim())) {
        if ($('srsWahlfach')) {
            $('srsWahlfach').value = (profile.defaultWahlfaecher || []).join('\n');
        }
    }
    updateListPreview();
}

function prefillFromStammdatenAndState() {
    const state = loadWizardState();
    const stammdaten = loadStammdaten();

    try {
        const setup =
            window.ms365AppDataV2 && window.ms365AppDataV2.getSetup
                ? window.ms365AppDataV2.getSetup()
                : null;
        const saved = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
        if ($('srsSiteUrl') && !$('srsSiteUrl').value) {
            $('srsSiteUrl').value = state.siteUrl || saved || '';
        }
    } catch {
        if ($('srsSiteUrl') && state.siteUrl) $('srsSiteUrl').value = state.siteUrl;
    }

    if ($('srsProfile')) {
        const pid = state.profileId && PROFILE_IDS.indexOf(state.profileId) !== -1 ? state.profileId : 'hak';
        $('srsProfile').value = pid;
    }

    const yearNow = String(new Date().getFullYear());
    if ($('srsJahr')) $('srsJahr').value = state.terminJahr || yearNow;
    if ($('srsBfk')) $('srsBfk').value = normalizeIsoDate(state.bfkDatum);

    const savedLfs = state.lfs || state.lfsText || '';
    const savedWahl = state.wahlfach || state.wahlfachText || '';
    const savedSeminar = state.seminar || state.seminarText || '';
    if ($('srsLfs') && savedLfs) $('srsLfs').value = savedLfs;
    if ($('srsWahlfach') && savedWahl) $('srsWahlfach').value = savedWahl;
    if ($('srsSeminar')) $('srsSeminar').value = savedSeminar;
    // Legacy: Textfelder → einmalig in Picker-Config übernehmen
    if (state.groupAdmin || state.groupLehrer || state.groupKandidaten) {
        const cur = loadEffectiveSrdpPermissionsConfig();
        saveSrdpPermissionsConfig({
            ...cur,
            groupAdmin: cur.groupAdmin || state.groupAdmin || '',
            groupLehrer: cur.groupLehrer || state.groupLehrer || '',
            groupKandidaten: cur.groupKandidaten || state.groupKandidaten || '',
            skipPerms: state.skipPerms != null ? !!state.skipPerms : cur.skipPerms
        });
    }
    const permsCfg = loadEffectiveSrdpPermissionsConfig();
    if ($('srsSkipPerms')) {
        $('srsSkipPerms').checked =
            state.skipPerms != null ? !!state.skipPerms : !!permsCfg.skipPerms;
    }

    const teacherCount = teacherChoiceLabels(stammdaten.teachers || []).length;
    const el = $('srsTeacherHint');
    if (el) {
        el.textContent =
            teacherCount +
            ' Lehrkraft/Lehrkräfte aus Stammdaten → Choice-Felder Betreuung / Fach / Wahlfach.';
    }

    applyProfileDefaults(!savedLfs && !savedWahl);
    extraKlassen = (Array.isArray(state.extraKlassen) ? state.extraKlassen : [])
        .map(normalizeExtraKlasse)
        .filter(Boolean);
    fillKlassenCheckboxes();
    updateListPreview();
}

function renderTerminPreview() {
    const host = $('srsTerminPreview');
    const list = $('srsTerminPreviewList');
    const lead = $('srsTerminPreviewLead');
    const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
    if (!host || !list) return;
    if (!/^\d{4}$/.test(jahr)) {
        host.hidden = true;
        return;
    }
    host.hidden = false;
    const bundle = periodenForTerminJahr(jahr);
    const haupt = bundle.perioden.find((row) => row.id === 'haupt');
    const bfkDatum =
        normalizeIsoDate(($('srsBfk') && $('srsBfk').value) || '') ||
        normalizeIsoDate(loadWizardState().bfkDatum);
    /** @type {Array<{ sort: string, label: string, when: string, bfk?: boolean }>} */
    const rows = [];
    (haupt && haupt.block ? haupt.block.faecher || [] : []).forEach((fach) => {
        const daten = (fach.daten || []).slice().filter(Boolean).sort();
        rows.push({
            sort: daten[0] || '9999-99-99',
            label: fach.fach,
            when: daten.map((iso) => formatGermanDate(iso)).filter(Boolean).join(' und ')
        });
    });
    rows.push({
        sort: bfkDatum || '9999-99-99',
        label: 'BFK schriftlich',
        when: '',
        bfk: true
    });
    rows.sort((a, b) => a.sort.localeCompare(b.sort) || (a.bfk ? 1 : 0) - (b.bfk ? 1 : 0));
    list.replaceChildren();
    if (!haupt) {
        if (lead) {
            lead.textContent =
                'Für ' + jahr + ' stehen auf matura.gv.at noch keine Termine. Sobald sie dort sind, erscheinen sie hier.';
        }
    } else if (lead) {
        lead.textContent = 'Schuljahr ' + (bundle.schuljahr || '–') + ', gelesen von matura.gv.at. Nach Datum sortiert.';
    }
    rows.forEach((row) => {
        const li = document.createElement('li');
        if (row.bfk) {
            li.className = 'srs-termin-bfk';
            const label = document.createElement('label');
            label.htmlFor = 'srsBfk';
            label.textContent = row.label;
            const input = document.createElement('input');
            input.type = 'date';
            input.id = 'srsBfk';
            input.value = bfkDatum;
            li.append(label, input);
        } else {
            const name = document.createElement('span');
            name.textContent = row.label;
            const when = document.createElement('span');
            when.textContent = row.when;
            li.append(name, when);
        }
        list.append(li);
    });
}

let terminPreviewTicket = 0;

function refreshTerminPreview() {
    renderTerminPreview();
    const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
    if (!/^\d{4}$/.test(jahr)) return;
    const ticket = ++terminPreviewTicket;
    refreshPruefungstermineFromWeb({ terminJahr: jahr })
        .then(() => {
            if (ticket === terminPreviewTicket) renderTerminPreview();
        })
        .catch(() => {});
}

function setBfkStatus(text) {
    const el = $('srsBfkStatus');
    if (el && text) el.textContent = text;
}

async function resolveSrdpList(webUrl, terminJahr, profileId) {
    const token = await ensureToken();
    const site = await G().resolveSiteFromWebUrl(token, webUrl);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const list = await findListByTitleCandidates(token, siteId, terminJahr, profileId);
    return { token, siteId, list };
}

async function upsertBfkColumn(siteId, listId, token, iso, write) {
    const date = normalizeIsoDate(iso);
    const colsPath =
        G().graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(listId) +
        '/columns?$select=id,name&$top=200';
    const colsData = await G().graphJson('GET', colsPath, token, undefined, 'v1.0');
    const col = ((colsData && colsData.value) || []).find((row) => String(row.name || '') === 'BfkSchriftlich');
    const description = date ? 'BFK schriftlich ' + date : 'BFK schriftlich';
    const patch = { description: description, readOnly: true };
    if (date) patch.defaultValue = { value: date };
    const base = G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
    const createBody = Object.assign(
        {
            name: 'BfkSchriftlich',
            displayName: 'BFK schriftlich',
            text: { allowMultipleLines: false, maxLength: 10 }
        },
        patch
    );
    const writeColumn = async (method, url, body) => {
        try {
            await G().graphJson(method, url, token, body, 'v1.0');
        } catch (err) {
            const slim = { description: body.description };
            if (body.name) {
                slim.name = body.name;
                slim.displayName = body.displayName;
                slim.text = body.text;
            }
            if (date) slim.defaultValue = { value: date };
            try {
                await G().graphJson(method, url, token, slim, 'v1.0');
            } catch (second) {
                if (!body.name) throw second;
                await G().graphJson(
                    method,
                    url,
                    token,
                    {
                        name: body.name,
                        displayName: body.displayName,
                        description: body.description,
                        text: body.text
                    },
                    'v1.0'
                );
            }
        }
    };
    if (!col) await writeColumn('POST', base, createBody);
    else await writeColumn('PATCH', base + '/' + encodeURIComponent(col.id), patch);
    if (typeof write === 'function') {
        write(date ? 'BFK schriftlich gespeichert: ' + date : 'BFK schriftlich geleert.');
    }
    return date;
}

async function readBfkFromList(webUrl, terminJahr, profileId) {
    const resolved = await resolveSrdpList(webUrl, terminJahr, profileId);
    if (!resolved.list || !resolved.list.id) return { foundList: false, iso: '', known: false };
    const colsPath =
        G().graphPathSite(resolved.siteId) +
        '/lists/' +
        encodeURIComponent(resolved.list.id) +
        '/columns?$select=name,description,defaultValue&$top=200';
    const colsData = await G().graphJson('GET', colsPath, resolved.token, undefined, 'v1.0');
    const col = ((colsData && colsData.value) || []).find((row) => String(row.name || '') === 'BfkSchriftlich');
    return { foundList: true, iso: bfkIsoFromColumn(col), known: !!col };
}

let bfkSaveTimer = 0;

function scheduleBfkSave() {
    persistForm();
    window.clearTimeout(bfkSaveTimer);
    bfkSaveTimer = window.setTimeout(() => {
        saveBfkFromForm().catch((err) => {
            setBfkStatus(
                'Datum liegt in diesem Browser. Die Liste hat es noch nicht übernommen: ' +
                    (err && err.message ? err.message : err)
            );
        });
    }, 400);
}

async function saveBfkFromForm() {
    const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
    const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
    const iso = normalizeIsoDate(($('srsBfk') && $('srsBfk').value) || '');
    if (!webUrl || !/^\d{4}$/.test(jahr)) {
        setBfkStatus('Datum gemerkt. Es wird in die SharePoint-Liste geschrieben, sobald Website und Terminjahr feststehen.');
        return;
    }
    const resolved = await resolveSrdpList(webUrl, jahr, currentProfileId());
    if (!resolved.list || !resolved.list.id) {
        setBfkStatus('Datum gemerkt. Die Liste für ' + jahr + ' gibt es noch nicht. Beim Anlegen wird es mitgeschrieben.');
        return;
    }
    await upsertBfkColumn(resolved.siteId, resolved.list.id, resolved.token, iso);
    setBfkStatus(
        iso
            ? 'In der SharePoint-Liste gespeichert. Schülerinnen und Schüler sehen den ' + formatGermanDate(iso) + '.'
            : 'In der SharePoint-Liste geleert.'
    );
}

async function pullBfkFromList() {
    const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
    const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
    if (!webUrl || !/^\d{4}$/.test(jahr) || !$('srsBfk')) return;
    try {
        const read = await readBfkFromList(webUrl, jahr, currentProfileId());
        if (!read.foundList || !read.known) return;
        $('srsBfk').value = read.iso || '';
        persistForm();
        renderTerminPreview();
        if (read.iso) setBfkStatus('Aus der Liste gelesen: ' + formatGermanDate(read.iso) + '.');
    } catch {
        /* Liste noch nicht lesbar */
    }
}

function updateListPreview() {
    const el = $('srsListPreview');
    const y = String(($('srsJahr') && $('srsJahr').value) || '').trim();
    const pid = currentProfileId();
    const yearLabel = /^\d{4}$/.test(y) ? y : 'YYYY';
    try {
        if (el) {
            el.textContent = /^\d{4}$/.test(y)
                ? listTitleForYear(y, pid)
                : listTitleForYear(new Date().getFullYear(), pid).replace(/\d{4}$/, 'YYYY');
        }
    } catch {
        if (el) el.textContent = 'sRDP-Anmeldungen … YYYY';
    }
    try {
        const abaName = /^\d{4}$/.test(y)
            ? abschlussarbeitenListTitleForYear(y, pid)
            : abschlussarbeitenListTitleForYear(new Date().getFullYear(), pid).replace(/\d{4}$/, yearLabel);
        ['srsAbaListPreview', 'srsAbaListPreviewTop'].forEach((id) => {
            const node = $(id);
            if (node) node.textContent = abaName;
        });
    } catch {
        ['srsAbaListPreview', 'srsAbaListPreviewTop'].forEach((id) => {
            const node = $(id);
            if (node) node.textContent = 'sRDP-Abschlussarbeiten … YYYY';
        });
    }
    refreshTerminPreview();
}

async function runCreate() {
    const logEl = $('srsLog');
    if (logEl) logEl.textContent = '';
    persistForm();
    const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
    const opts = readFormOpts();
    if (!/^\d{4}$/.test(opts.terminJahr)) throw new Error('Terminjahr (YYYY) fehlt.');
    if (!opts.klassen.length) throw new Error('Mindestens eine Abschlussklasse auswählen.');
    if (!opts.lehrer.length) throw new Error('Keine Lehrer in den Stammdaten – bitte zuerst pflegen.');
    if (!opts.lfs.length) throw new Error('Mindestens eine LFS-/Sprach-Option eintragen.');
    if (!opts.wahlfaecher.length) throw new Error('Mindestens ein Wahlfach eintragen.');
    if (opts.applyPermissions) {
        const okAdmin = audienceHasGrant(
            opts.adminGroups,
            opts.groupAdminId,
            opts.groupAdmin,
            opts.adminUsers
        );
        const okLehrer = audienceHasGrant(
            opts.lehrerGroups,
            opts.groupLehrerId,
            opts.groupLehrer,
            opts.lehrerUsers
        );
        const okKand = audienceHasGrant(
            opts.kandidatenGroups,
            opts.groupKandidatenId,
            opts.groupKandidaten,
            opts.kandidatenUsers
        );
        if (!okAdmin || !okLehrer || !okKand) {
            throw new Error(
                'Für Rechte: je Stufe Gruppe oder Person(en) wählen (Vollzugriff, Lesende Rechte, Kandidat:innen) – oder „Rechte überspringen“.'
            );
        }
    }
    const result = await createSrdpList(webUrl, opts);
    log('Lege Abschlussarbeiten-Liste an (ABA-Portal) …');
    const abaCtx = await ensureAbschlussarbeitenList(webUrl, {
        jahr: opts.terminJahr,
        profileId: opts.profileId
    });
    log(
        'Abschlussarbeiten bereit: „' +
            (abaCtx.list && abaCtx.list.title) +
            '“' +
            (abaCtx.list && abaCtx.list.webUrl ? ' · ' + abaCtx.list.webUrl : '')
    );
    saveWizardState({
        siteUrl: webUrl,
        listId: result && result.list && result.list.id ? String(result.list.id) : '',
        profileId: opts.profileId,
        jahr: opts.terminJahr,
        terminJahr: opts.terminJahr,
        klassen: Array.isArray(opts.klassen) ? opts.klassen.slice() : [],
        lfs: Array.isArray(opts.lfs) ? opts.lfs.join('\n') : '',
        wahlfach: Array.isArray(opts.wahlfaecher) ? opts.wahlfaecher.join('\n') : '',
        seminar: Array.isArray(opts.seminare) ? opts.seminare.join('\n') : '',
        lfsText: Array.isArray(opts.lfs) ? opts.lfs.join('\n') : '',
        wahlfachText: Array.isArray(opts.wahlfaecher) ? opts.wahlfaecher.join('\n') : '',
        seminarText: Array.isArray(opts.seminare) ? opts.seminare.join('\n') : ''
    });
    try {
        patchSchoolToolSitesSetup({ planerSiteUrls: { srdp: webUrl } });
    } catch {
        /* ignore */
    }
    toast(
        'Listen angelegt bzw. ergänzt (' +
            getProfile(opts.profileId).label +
            '): Anmeldung + Abschlussarbeiten.'
    );
    const link = $('srsListLink');
    if (link && result.webUrl) {
        link.hidden = false;
        link.href = result.webUrl;
        link.textContent = 'Anmeldeliste öffnen: ' + result.title;
    }
    const abaLink = $('srsAbaListLink');
    if (abaLink && abaCtx.list && abaCtx.list.webUrl) {
        abaLink.hidden = false;
        abaLink.href = abaCtx.list.webUrl;
        abaLink.textContent = 'Abschlussarbeiten öffnen: ' + (abaCtx.list.title || '');
    }
    setAbaStatus(
        'Liste „' +
            (abaCtx.list && abaCtx.list.title) +
            '“ bereit – jetzt optional ABA-Excel hochladen.'
    );
    return { ...result, aba: abaCtx };
}

window.ms365SpoSrdp = {
    createList: createSrdpList,
    probeHealth: probeSrdpHealth,
    backfillItemTitles: backfillItemTitles,
    listTitleForYear,
    getProfile,
    PROFILE_IDS,
    STORAGE_KEY
};

function clampWizardStep(n) {
    const step = parseInt(String(n), 10);
    if (!Number.isFinite(step)) return 1;
    return Math.max(1, Math.min(SRS_WIZARD_STEP_COUNT, step));
}

function loadWizardStep() {
    try {
        return clampWizardStep(parseInt(localStorage.getItem(SRS_WIZARD_STEP_STORAGE_KEY) || '1', 10));
    } catch {
        return 1;
    }
}

function saveWizardStep(n) {
    try {
        localStorage.setItem(SRS_WIZARD_STEP_STORAGE_KEY, String(clampWizardStep(n)));
    } catch {
        /* ignore */
    }
}

function validateWizardStep(step) {
    const s = clampWizardStep(step);
    if (s === 1) {
        const site = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
        const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
        if (!site) return 'Bitte die SharePoint-Website eintragen.';
        if (!/^\d{4}$/.test(jahr)) return 'Bitte ein gültiges Terminjahr (YYYY) eintragen.';
    }
    if (s === 2) {
        const klassen = Array.from(document.querySelectorAll('input[name="srsKlasse"]:checked'));
        if (!klassen.length) return 'Mindestens eine Abschlussklasse auswählen.';
    }
    if (s === 3) {
        const lfs = parseChoiceLines(($('srsLfs') && $('srsLfs').value) || '');
        const wahl = parseChoiceLines(($('srsWahlfach') && $('srsWahlfach').value) || '');
        if (!lfs.length) return 'Mindestens eine LFS-/Sprach-Option eintragen.';
        if (!wahl.length) return 'Mindestens ein Wahlfach eintragen.';
    }
    return null;
}

function refreshWizardSummary() {
    const box = $('srsWizardSummary');
    if (!box) return;
    const opts = readFormOpts();
    const site = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim() || '—';
    const profile = getProfile(opts.profileId);
    let listName = '—';
    try {
        listName = /^\d{4}$/.test(opts.terminJahr)
            ? listTitleForYear(opts.terminJahr, opts.profileId)
            : '—';
    } catch {
        listName = '—';
    }
    box.innerHTML =
        '<li><strong>Website:</strong> ' +
        site +
        '</li>' +
        '<li><strong>Schulform / Anmeldeliste:</strong> ' +
        profile.label +
        ' · ' +
        listName +
        '</li>' +
        '<li><strong>Abschlussarbeiten:</strong> ' +
        (function () {
            try {
                return /^\d{4}$/.test(opts.terminJahr)
                    ? abschlussarbeitenListTitleForYear(opts.terminJahr, opts.profileId)
                    : '—';
            } catch {
                return '—';
            }
        })() +
        '</li>' +
        '<li><strong>Klassen:</strong> ' +
        (opts.klassen.length ? opts.klassen.join(', ') : '—') +
        '</li>' +
        '<li><strong>Vollzugriff:</strong> ' +
        formatSrdpAudienceSummary({
            label: opts.groupAdmin,
            id: opts.groupAdminId,
            groups: opts.adminGroups,
            users: opts.adminUsers
        }) +
        '</li>' +
        '<li><strong>Lesende Rechte:</strong> ' +
        formatSrdpAudienceSummary({
            label: opts.groupLehrer,
            id: opts.groupLehrerId,
            groups: opts.lehrerGroups,
            users: opts.lehrerUsers
        }) +
        '</li>' +
        '<li><strong>Kandidat:innen:</strong> ' +
        formatSrdpAudienceSummary({
            label: opts.groupKandidaten,
            id: opts.groupKandidatenId,
            groups: opts.kandidatenGroups,
            users: opts.kandidatenUsers
        }) +
        '</li>' +
        '<li><strong>Rechte:</strong> ' +
        (opts.applyPermissions ? 'mit anwenden' : 'diesmal überspringen') +
        '</li>';
}

function showWizardStep(n) {
    const step = clampWizardStep(n);
    saveWizardStep(step);
    for (let i = 1; i <= SRS_WIZARD_STEP_COUNT; i++) {
        const panel = $('srsWizardStep' + i);
        if (panel) panel.hidden = i !== step;
    }
    document.querySelectorAll('#srsWizardGlance [data-srs-wizard-step]').forEach((btn) => {
        const s = parseInt(btn.getAttribute('data-srs-wizard-step') || '0', 10);
        const active = s === step;
        btn.classList.toggle('is-active', active);
        btn.setAttribute('aria-selected', active ? 'true' : 'false');
    });
    const hint = $('srsPhaseHint');
    if (hint) {
        hint.textContent =
            'Schritt ' + step + ' von ' + SRS_WIZARD_STEP_COUNT + ': ' + SRS_WIZARD_STEP_LABELS[step - 1];
    }
    const back = $('srsWizardBack');
    const next = $('srsWizardNext');
    if (back) back.disabled = step <= 1;
    if (next) {
        next.hidden = step >= SRS_WIZARD_STEP_COUNT;
        if (step < SRS_WIZARD_STEP_COUNT) {
            const label = step === SRS_WIZARD_STEP_COUNT - 1 ? 'Zur Übersicht' : 'Weiter';
            next.innerHTML = label + ' <i class="bi bi-arrow-right"></i>';
        }
    }
    if (step === 2) fillKlassenCheckboxes();
    if (step === 5) {
        mountSrdpEntraPermHost($('srsEntraPermHost'), () => {
            persistSrdpPickersToStorage($('srsSkipPerms') && $('srsSkipPerms').checked);
        });
    }
    if (step === 6) refreshWizardSummary();
}

function wireUi() {
    if ($('srsProfile')) {
        PROFILE_IDS.forEach((id) => {
            const p = getProfile(id);
            const opt = document.createElement('option');
            opt.value = p.id;
            opt.textContent = p.label + ' (' + p.examShort + ')';
            $('srsProfile').appendChild(opt);
        });
    }

    prefillFromStammdatenAndState();
    mountSrdpEntraPermHost($('srsEntraPermHost'), () => {
        persistSrdpPickersToStorage($('srsSkipPerms') && $('srsSkipPerms').checked);
    });
    showWizardStep(loadWizardStep());
    void pullBfkFromList();

    if ($('srsProfile')) {
        $('srsProfile').addEventListener('change', () => {
            applyProfileDefaults(true);
            persistForm();
            updateListPreview();
        });
    }

    if ($('srsJahr')) {
        $('srsJahr').addEventListener('input', () => {
            fillKlassenCheckboxes();
            updateListPreview();
        });
        $('srsJahr').addEventListener('change', () => {
            fillKlassenCheckboxes();
            updateListPreview();
            void pullBfkFromList();
        });
    }

    const terminList = $('srsTerminPreviewList');
    if (terminList) {
        terminList.addEventListener('change', (ev) => {
            const target = ev.target;
            if (!(target instanceof HTMLInputElement) || target.id !== 'srsBfk') return;
            scheduleBfkSave();
            renderTerminPreview();
        });
    }

    ['srsLfs', 'srsWahlfach', 'srsSeminar'].forEach((id) => {
        const el = $(id);
        if (!el) return;
        el.addEventListener('change', () => persistForm());
    });

    if ($('srsSkipPerms')) {
        $('srsSkipPerms').addEventListener('change', () => {
            persistSrdpPickersToStorage($('srsSkipPerms').checked);
        });
    }

    const back = $('srsWizardBack');
    if (back) {
        back.addEventListener('click', () => {
            persistForm();
            showWizardStep(loadWizardStep() - 1);
        });
    }
    const next = $('srsWizardNext');
    if (next) {
        next.addEventListener('click', () => {
            const cur = loadWizardStep();
            const err = validateWizardStep(cur);
            if (err) {
                toast(err);
                return;
            }
            persistForm();
            showWizardStep(cur + 1);
        });
    }
    document.querySelectorAll('#srsWizardGlance [data-srs-wizard-step]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const target = clampWizardStep(btn.getAttribute('data-srs-wizard-step'));
            const cur = loadWizardStep();
            if (target > cur) {
                for (let s = cur; s < target; s++) {
                    const err = validateWizardStep(s);
                    if (err) {
                        toast(err);
                        showWizardStep(s);
                        return;
                    }
                }
            }
            persistForm();
            showWizardStep(target);
        });
    });

    const runBtn = $('srsBtnRun');
    if (runBtn) {
        runBtn.addEventListener('click', () => {
            const y = String(($('srsJahr') && $('srsJahr').value) || '').trim();
            const pid = currentProfileId();
            let title = 'Anmeldung';
            try {
                title = listTitleForYear(y, pid);
            } catch {
                /* ignore */
            }
            let abaTitle = 'Abschlussarbeiten';
            try {
                abaTitle = abschlussarbeitenListTitleForYear(y, pid);
            } catch {
                /* ignore */
            }
            if (
                !window.confirm(
                    'Listen anlegen bzw. ergänzen?\n\n• ' +
                        title +
                        '\n• ' +
                        abaTitle +
                        '\n\nRechte werden gesetzt (außer übersprungen).'
                )
            ) {
                return;
            }
            const panel = $('srsLogPanel');
            if (panel) panel.open = true;
            runCreate().catch((e) => {
                log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                toast('Fehler: ' + (e && e.message ? e.message : e));
            });
        });
    }

    const probeBtn = $('srsBtnProbe');
    if (probeBtn) {
        probeBtn.addEventListener('click', () => {
            if ($('srsLog')) $('srsLog').textContent = '';
            const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
            if (!webUrl) {
                toast('Website-URL fehlt.');
                return;
            }
            ensureToken()
                .then((token) => G().resolveSiteFromWebUrl(token, webUrl))
                .then((site) => {
                    log('Site gefunden: ' + (site.displayName || '') + '\nid: ' + (site.id || ''));
                    if (site.webUrl) log('webUrl: ' + site.webUrl);
                    toast('Website erkannt.');
                    void pullBfkFromList();
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    const healthBtn = $('srsBtnHealth');
    if (healthBtn) {
        healthBtn.addEventListener('click', () => {
            if ($('srsLog')) $('srsLog').textContent = '';
            const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
            const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
            probeSrdpHealth(webUrl, jahr, log, currentProfileId())
                .then((summary) => {
                    toast(summary && summary.ok ? 'Listen-Check OK' : 'Listen-Check: bitte Protokoll prüfen');
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    async function runAbaEnsureOnly() {
        const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
        const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
        if (!webUrl) throw new Error('Website-URL fehlt.');
        if (!/^\d{4}$/.test(jahr)) throw new Error('Terminjahr (YYYY) fehlt.');
        const ctx = await ensureAbschlussarbeitenList(webUrl, {
            jahr,
            profileId: currentProfileId()
        });
        log('Abschlussarbeiten-Liste: „' + (ctx.list && ctx.list.title) + '“');
        const abaLink = $('srsAbaListLink');
        if (abaLink && ctx.list && ctx.list.webUrl) {
            abaLink.hidden = false;
            abaLink.href = ctx.list.webUrl;
            abaLink.textContent = 'Abschlussarbeiten öffnen: ' + (ctx.list.title || '');
        }
        setAbaStatus('Liste „' + (ctx.list && ctx.list.title) + '“ bereit.');
        toast('Abschlussarbeiten-Liste bereit.');
        return ctx;
    }

    async function runAbaUpload(file) {
        if (!file) return;
        const XLSX = window.XLSX;
        if (!XLSX) throw new Error('Excel-Bibliothek (SheetJS) nicht geladen – Seite neu laden.');
        const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
        const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
        if (!webUrl) throw new Error('Website-URL fehlt.');
        if (!/^\d{4}$/.test(jahr)) throw new Error('Terminjahr (YYYY) fehlt.');
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const parsed = parseAbaPortalRows(jsonRows, { terminJahr: jahr });
        if (!parsed.rows.length) {
            throw new Error(
                'Keine Zeilen für Terminjahr ' +
                    jahr +
                    ' (übersprungen Jahr: ' +
                    parsed.summary.skippedYear +
                    ', leer: ' +
                    parsed.summary.skippedEmpty +
                    ').'
            );
        }
        const ctx = await ensureAbschlussarbeitenList(webUrl, {
            jahr,
            profileId: currentProfileId()
        });
        const result = await replaceAbschlussarbeiten(ctx, parsed.rows, {
            jahr,
            profileId: currentProfileId()
        });
        log(
            'ABA-Import: ' +
                result.count +
                ' Personen · genehmigt in Datei ' +
                parsed.summary.genehmigt +
                ' · Liste „' +
                (ctx.list && ctx.list.title) +
                '“'
        );
        setAbaStatus(
            'Import OK: ' + result.count + ' Personen in „' + (ctx.list && ctx.list.title) + '“.'
        );
        const abaLink = $('srsAbaListLink');
        if (abaLink && ctx.list && ctx.list.webUrl) {
            abaLink.hidden = false;
            abaLink.href = ctx.list.webUrl;
            abaLink.textContent = 'Abschlussarbeiten öffnen: ' + (ctx.list.title || '');
        }
        toast('ABA-Excel importiert: ' + result.count + ' Personen.');
        return result;
    }

    const abaEnsureBtn = $('srsBtnAbaEnsure');
    if (abaEnsureBtn) {
        abaEnsureBtn.addEventListener('click', () => {
            if ($('srsLog')) $('srsLog').textContent = '';
            const panel = $('srsLogPanel');
            if (panel) panel.open = true;
            runAbaEnsureOnly().catch((e) => {
                log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                toast('Fehler: ' + (e && e.message ? e.message : e));
            });
        });
    }

    const abaUploadBtn = $('srsBtnAbaUpload');
    const abaFile = $('srsAbaFile');
    if (abaUploadBtn && abaFile) {
        abaUploadBtn.addEventListener('click', () => abaFile.click());
        abaFile.addEventListener('change', () => {
            const file = abaFile.files && abaFile.files[0];
            if (!file) return;
            if ($('srsLog')) $('srsLog').textContent = '';
            const panel = $('srsLogPanel');
            if (panel) panel.open = true;
            runAbaUpload(file)
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                })
                .finally(() => {
                    abaFile.value = '';
                });
        });
    }

    const permsBtn = $('srsBtnPerms');
    if (permsBtn) {
        permsBtn.addEventListener('click', () => {
            if ($('srsLog')) $('srsLog').textContent = '';
            const panel = $('srsLogPanel');
            if (panel) panel.open = true;
            persistForm();
            const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
            const opts = readFormOpts();
            opts.applyPermissions = true;
            reapplySrdpPermissions(webUrl, opts, log)
                .then(() => toast('Rechte nachgezogen.'))
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    const titleBtn = $('srsBtnTitles');
    if (titleBtn) {
        titleBtn.addEventListener('click', () => {
            if ($('srsLog')) $('srsLog').textContent = '';
            const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
            const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
            backfillItemTitles(webUrl, jahr, log, currentProfileId())
                .then((r) => {
                    toast('Titel: ' + (r.patched || 0) + ' aktualisiert');
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    const saveBtn = $('srsBtnSave');
    if (saveBtn) {
        saveBtn.addEventListener('click', () => {
            persistForm();
            toast('Einstellungen gespeichert.');
        });
    }

    const refreshBtn = $('srsBtnRefreshKlassen');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', () => fillKlassenCheckboxes());
    }

    const pickKlassenBtn = $('srsBtnPickKlassenGroups');
    if (pickKlassenBtn) {
        pickKlassenBtn.addEventListener('click', () => {
            pickExtraKlassen();
        });
    }

    document.addEventListener('click', (ev) => {
        const t = ev.target;
        if (!(t instanceof Element)) return;
        const remove = t.closest('[data-srs-remove-extra]');
        if (!remove) return;
        const id = remove.getAttribute('data-srs-remove-extra') || '';
        extraKlassen = extraKlassen.filter((x) => (x.groupId || x.code) !== id);
        renderExtraKlassen();
        persistForm();
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wireUi);
} else {
    wireUi();
}
