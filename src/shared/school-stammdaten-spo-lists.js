/**
 * Schul-Stammdaten als SharePoint-Listen (Lehrer, Klassen inkl. KV).
 * Wahrheit liegt im Schul-Tenant – App und manuelle SP-Pflege nutzen dieselben Listen.
 *
 * JSON-IT-Backup bleibt Absicherung; Live-Listen sind der Betriebs-Pfad.
 */

import {
    DEFAULT_INTRANET_LIST_TITLES,
    resolveIntranetListTitle
} from './intranet-list-title-logic.js';

/** @typedef {{ name: string, displayName?: string, text?: object, personOrGroup?: object, number?: object, boolean?: object, choice?: object, dateTime?: object }} SpoColumnDef */
/** @typedef {{ code: string, name: string, email?: string, upn?: string }} StammdatenTeacher */
/**
 * @typedef {{
 *   code: string,
 *   name: string,
 *   year?: string,
 *   headName?: string,
 *   headEmail?: string,
 *   kvEmail?: string,
 *   m365GroupId?: string,
 *   m365GroupMail?: string,
 *   m365GroupName?: string,
 *   studentEmails?: string[]
 * }} StammdatenClass
 */

export const STAMMDATEN_SPO_LIST_KEYS = {
    lehrer: 'lehrer',
    klassen: 'klassen'
};

/**
 * Standard-Titel (anpassbar über Intranet-Listen-Titel in den Stammdaten).
 */
export function defaultStammdatenSpoListTitles() {
    return {
        lehrer: DEFAULT_INTRANET_LIST_TITLES.lehrer || 'Lehrerinnen',
        klassen: DEFAULT_INTRANET_LIST_TITLES.klassen || 'Klassen'
    };
}

/**
 * Titel aus Setup / gespeicherten Intranet-Namen + Defaults.
 * @param {Record<string, string>|null|undefined} stored
 */
export function resolveStammdatenSpoListTitles(stored) {
    const s = stored && typeof stored === 'object' ? stored : {};
    return {
        lehrer: resolveIntranetListTitle('lehrer', s.lehrer),
        klassen: resolveIntranetListTitle('klassen', s.klassen)
    };
}

/**
 * Browser: Titel aus App-Data / Tenant-Settings.
 * @param {Window|null|undefined} [win]
 */
export function resolveStammdatenSpoListTitlesFromWindow(win) {
    const w = win || (typeof window !== 'undefined' ? window : null);
    if (!w) return defaultStammdatenSpoListTitles();
    try {
        const setup =
            w.ms365AppDataV2 && typeof w.ms365AppDataV2.getSetup === 'function'
                ? w.ms365AppDataV2.getSetup() || {}
                : {};
        const fromSetup =
            setup.intranetListTitles && typeof setup.intranetListTitles === 'object'
                ? setup.intranetListTitles
                : null;
        const fromTenant =
            typeof w.ms365TenantSettingsLoad === 'function'
                ? (w.ms365TenantSettingsLoad() || {}).intranetListTitles
                : null;
        return resolveStammdatenSpoListTitles(fromSetup || fromTenant || {});
    } catch {
        return defaultStammdatenSpoListTitles();
    }
}

/**
 * @param {unknown} v
 */
function personFieldEmail(v) {
    if (v == null) return '';
    if (typeof v === 'string') return v.trim();
    if (Array.isArray(v)) return personFieldEmail(v[0]);
    if (typeof v === 'object') {
        return String(v.Email || v.email || v.UserPrincipalName || v.upn || '').trim();
    }
    return '';
}

/**
 * Mehrfach-Personenfeld → E-Mail-Liste.
 * @param {unknown} v
 * @returns {string[]}
 */
function personFieldEmails(v) {
    if (v == null || v === '') return [];
    const list = Array.isArray(v) ? v : [v];
    const out = [];
    const seen = new Set();
    list.forEach((row) => {
        const em = String(personFieldEmail(row) || '')
            .trim()
            .toLowerCase();
        if (!em || em.indexOf('@') < 0 || seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

/**
 * @param {unknown} v
 */
function personFieldName(v) {
    if (v == null || typeof v === 'string') return '';
    if (Array.isArray(v)) return personFieldName(v[0]);
    if (typeof v === 'object') {
        return String(v.LookupValue || v.DisplayName || v.Title || v.name || '').trim();
    }
    return '';
}

function normClassCode(s) {
    return String(s || '')
        .trim()
        .toUpperCase();
}

/**
 * LoginName für Entra-Gruppe in SharePoint (EnsureUser).
 * @param {string} groupObjectId
 */
export function entraGroupLogonNameForSite(groupObjectId) {
    const id = String(groupObjectId || '').trim();
    if (!id) return '';
    return 'c:0o.c|federateddirectoryclaimprovider|' + id;
}

/**
 * Klasse → verknüpfte M365-Gruppe aus lokalen Stammdaten (classTeams / classGroupMatchByKey).
 * @param {string} klasseCode
 * @param {Window|null|undefined} [win]
 * @returns {{ groupId: string, displayName: string, mail: string, mailNickname: string, code: string }|null}
 */
export function lookupClassM365GroupFromAppData(klasseCode, win) {
    const code = String(klasseCode || '').trim();
    if (!code) return null;
    if (/[,;~]/.test(code) || /(?:\d+[A-Za-z]+){2,}/.test(code)) return null;
    const w = win || (typeof window !== 'undefined' ? window : null);
    const api = w && w.ms365AppDataV2 ? w.ms365AppDataV2 : null;
    if (!api || typeof api.getContainer !== 'function') return null;

    const container = api.getContainer();
    const teams =
        typeof api.normalizeCoreClassTeams === 'function'
            ? api.normalizeCoreClassTeams((container.core && container.core.classTeams) || [])
            : (container.core && container.core.classTeams) || [];
    const want = normClassCode(code);

    for (let i = 0; i < teams.length; i++) {
        const t = teams[i];
        if (!t) continue;
        const cc = normClassCode(t.classCode || '');
        const dn = String(t.displayName || '').trim();
        const match = (dn && code === dn) || (cc && want === cc);
        if (!match) continue;
        return {
            groupId: t.graphGroupId ? String(t.graphGroupId).trim() : '',
            displayName: dn || code,
            mail: String(t.mail || t.gruppenmail || '')
                .trim()
                .toLowerCase(),
            mailNickname: String(t.mailNickname || t.stableMailNickname || '').trim(),
            code: cc || want
        };
    }

    const setup = typeof api.getSetup === 'function' ? api.getSetup() : null;
    const map = (setup && setup.classGroupMatchByKey) || {};
    const entry = map[want] || map[code] || null;
    if (!entry || typeof entry !== 'object') return null;
    return {
        groupId: String(entry.groupId || entry.graphGroupId || '').trim(),
        displayName: String(entry.displayName || entry.name || code).trim(),
        mail: String(entry.mail || entry.email || '')
            .trim()
            .toLowerCase(),
        mailNickname: String(entry.mailNickname || entry.alias || '').trim(),
        code: want
    };
}

/**
 * Ergänzt Klasse um M365-Gruppenfelder (App-Data / classTeams).
 * @param {StammdatenClass|Record<string, unknown>|null|undefined} c
 * @param {Window|null|undefined} [win]
 * @returns {StammdatenClass}
 */
export function enrichClassWithM365Group(c, win) {
    const row = c && typeof c === 'object' ? { ...c } : { code: '', name: '' };
    const code = String(row.code || '').trim();
    const existingId = String(row.m365GroupId || '').trim();
    const existingMail = String(row.m365GroupMail || '')
        .trim()
        .toLowerCase();
    if (existingId && existingMail) {
        return /** @type {StammdatenClass} */ (row);
    }
    const link = lookupClassM365GroupFromAppData(code || String(row.name || ''), win);
    if (!link) return /** @type {StammdatenClass} */ (row);
    if (!existingId && link.groupId) row.m365GroupId = link.groupId;
    if (!existingMail && link.mail) row.m365GroupMail = link.mail;
    if (!String(row.m365GroupName || '').trim() && link.displayName) {
        row.m365GroupName = link.displayName;
    }
    return /** @type {StammdatenClass} */ (row);
}

/**
 * SP-Felder Lehrerinnen → App-Teacher.
 * @param {Record<string, unknown>|null|undefined} fields
 * @returns {StammdatenTeacher|null}
 */
export function mapLehrerFieldsToTeacher(fields) {
    const f = fields || {};
    const code = String(f.LehrerCode || f.lehrerCode || '').trim();
    const email = String(f.EMail || f.Email || f.UPN || personFieldEmail(f.Lehrkraft) || '')
        .trim()
        .toLowerCase();
    const upn = String(f.UPN || email).trim().toLowerCase();
    const fromParts = [f.Vorname, f.Nachname]
        .map((x) => String(x || '').trim())
        .filter(Boolean)
        .join(' ');
    const name =
        String(f.Title || '').trim() ||
        fromParts ||
        personFieldName(f.Lehrkraft) ||
        code ||
        email;
    if (!code && !email && !name) return null;
    return {
        code: code || email || name,
        name: name || code || email,
        email: email || upn || '',
        upn: upn || email || ''
    };
}

/**
 * SP-Felder Klassen → App-Class (inkl. KV).
 * @param {Record<string, unknown>|null|undefined} fields
 * @returns {StammdatenClass|null}
 */
export function mapKlassenFieldsToClass(fields) {
    const f = fields || {};
    const code = String(f.KlassenCode || f.klasseCode || f.Kuerzel || f.Kürzel || '').trim();
    const name = String(f.Title || code).trim() || code;
    const year = String(f.Abschlussjahr || f.year || '').trim();
    const headEmail = String(f.KVEmail || personFieldEmail(f.Klassenvorstand) || '')
        .trim()
        .toLowerCase();
    const headName =
        String(f.KVName || '').trim() || personFieldName(f.Klassenvorstand) || '';
    const m365GroupId = String(f.M365GroupId || f.m365GroupId || '').trim();
    const m365FromPerson = personFieldEmail(f.M365Gruppe);
    const m365GroupMail = String(f.M365GroupMail || f.m365GroupMail || m365FromPerson || '')
        .trim()
        .toLowerCase();
    const m365GroupName =
        String(f.M365GroupName || f.m365GroupName || '').trim() ||
        personFieldName(f.M365Gruppe) ||
        '';
    const studentEmails = personFieldEmails(f.Schuelerinnen || f.Schueler);
    if (!code && !name) return null;
    /** @type {StammdatenClass} */
    const out = {
        code: code || name,
        name: name || code,
        year,
        headName,
        headEmail,
        kvEmail: headEmail
    };
    if (m365GroupId) out.m365GroupId = m365GroupId;
    if (m365GroupMail) out.m365GroupMail = m365GroupMail;
    if (m365GroupName) out.m365GroupName = m365GroupName;
    if (studentEmails.length) out.studentEmails = studentEmails;
    return out;
}

/**
 * Live bevorzugt, sonst lokal (z. B. Demo / leere Listen).
 * @param {{ teachers?: unknown[], classes?: unknown[], students?: unknown[] }|null|undefined} local
 * @param {{ teachers?: StammdatenTeacher[], classes?: StammdatenClass[] }|null|undefined} live
 */
export function mergeStammdatenPreferLive(local, live) {
    const loc = local || {};
    const liv = live || {};
    const liveTeachers = Array.isArray(liv.teachers) ? liv.teachers : [];
    const liveClasses = Array.isArray(liv.classes) ? liv.classes : [];
    return {
        teachers: liveTeachers.length
            ? liveTeachers
            : Array.isArray(loc.teachers)
              ? loc.teachers
              : [],
        classes: liveClasses.length
            ? liveClasses
            : Array.isArray(loc.classes)
              ? loc.classes
              : [],
        students: Array.isArray(loc.students) ? loc.students : []
    };
}

/** @type {SpoColumnDef[]} */
export const LEHRER_LIST_COLUMNS = [
    { name: 'Vorname', displayName: 'Vorname', text: { allowMultipleLines: false, maxLength: 120 } },
    { name: 'Nachname', displayName: 'Nachname', text: { allowMultipleLines: false, maxLength: 120 } },
    { name: 'LehrerCode', displayName: 'Kürzel', text: { allowMultipleLines: false, maxLength: 40 } },
    { name: 'EMail', displayName: 'E-Mail', text: { allowMultipleLines: false, maxLength: 255 } },
    { name: 'UPN', displayName: 'UPN', text: { allowMultipleLines: false, maxLength: 255 } },
    {
        name: 'Lehrkraft',
        displayName: 'Lehrkraft (Person)',
        personOrGroup: { allowMultipleSelection: false, chooseFromType: 'peopleOnly' }
    }
];

/** @type {SpoColumnDef[]} */
export const KLASSEN_LIST_COLUMNS = [
    { name: 'KlassenCode', displayName: 'Kürzel', text: { allowMultipleLines: false, maxLength: 40 } },
    { name: 'Abschlussjahr', displayName: 'Abschlussjahr', text: { allowMultipleLines: false, maxLength: 4 } },
    { name: 'KVName', displayName: 'Klassenvorstand', text: { allowMultipleLines: false, maxLength: 120 } },
    { name: 'KVEmail', displayName: 'KV E-Mail', text: { allowMultipleLines: false, maxLength: 255 } },
    {
        name: 'Klassenvorstand',
        displayName: 'KV (Person)',
        personOrGroup: { allowMultipleSelection: false, chooseFromType: 'peopleOnly' }
    },
    {
        name: 'M365GroupId',
        displayName: 'M365-Gruppen-ID',
        text: { allowMultipleLines: false, maxLength: 80 }
    },
    {
        name: 'M365GroupMail',
        displayName: 'M365-Gruppen-E-Mail',
        text: { allowMultipleLines: false, maxLength: 255 }
    },
    {
        name: 'M365GroupName',
        displayName: 'M365-Gruppenname',
        text: { allowMultipleLines: false, maxLength: 200 }
    },
    {
        name: 'M365Gruppe',
        displayName: 'M365-Gruppe (Team)',
        personOrGroup: { allowMultipleSelection: false, chooseFromType: 'peopleAndGroups' }
    }
];

export const STAMMDATEN_SPO_REQUIRED_COLUMNS = {
    lehrer: LEHRER_LIST_COLUMNS.map((c) => c.name),
    klassen: KLASSEN_LIST_COLUMNS.map((c) => c.name)
};

/**
 * @param {SpoColumnDef} def
 */
export function stammdatenColumnToGraphBody(def) {
    const body = {
        name: def.name,
        displayName: def.displayName || def.name
    };
    if (def.text) body.text = def.text;
    if (def.number) body.number = def.number;
    if (def.boolean) body.boolean = def.boolean;
    if (def.dateTime) body.dateTime = def.dateTime;
    if (def.choice) body.choice = def.choice;
    if (def.personOrGroup) body.personOrGroup = def.personOrGroup;
    return body;
}

/**
 * Paket Lehrer + Klassen auf einer Site sicherstellen (nur Schema, keine Zeilen).
 * Idempotent: vorhandene Listen behalten, fehlende Spalten ergänzen.
 *
 * @param {{
 *   token: string,
 *   siteId: string,
 *   graph: { graphPathSite: Function, graphJson: Function, sleep?: Function },
 *   write?: (msg: string) => void,
 *   titles?: { lehrer?: string, klassen?: string }
 * }} opts
 */
export async function ensureSchoolStammdatenShellLists(opts) {
    const o = opts || {};
    const write = typeof o.write === 'function' ? o.write : function () {};
    const G = o.graph;
    const token = o.token;
    const siteId = o.siteId;
    if (!G || !token || !siteId) throw new Error('ensureSchoolStammdatenShellLists: token/siteId/graph fehlen.');

    const defaults = defaultStammdatenSpoListTitles();
    const titles = {
        lehrer: String((o.titles && o.titles.lehrer) || defaults.lehrer).trim() || defaults.lehrer,
        klassen: String((o.titles && o.titles.klassen) || defaults.klassen).trim() || defaults.klassen
    };

    const packs = [
        { key: 'lehrer', title: titles.lehrer, columns: LEHRER_LIST_COLUMNS },
        { key: 'klassen', title: titles.klassen, columns: KLASSEN_LIST_COLUMNS }
    ];

    /** @type {Record<string, { id: string, webUrl: string, title: string, created: boolean }>} */
    const lists = {};

    for (let p = 0; p < packs.length; p++) {
        const pack = packs[p];
        write('Stammdaten-Paket: „' + pack.title + '" …');
        let list = await findListByTitle(G, token, siteId, pack.title);
        let created = false;
        if (!list || !list.id) {
            write('  Erstelle Liste „' + pack.title + '" …');
            const createdList = await G.graphJson(
                'POST',
                G.graphPathSite(siteId) + '/lists',
                token,
                { displayName: pack.title, list: { template: 'genericList' } },
                'v1.0'
            );
            const listId = createdList && createdList.id ? String(createdList.id) : '';
            if (!listId) throw new Error('Listen-ID fehlt für „' + pack.title + '“.');
            list = { id: listId, displayName: pack.title, webUrl: createdList.webUrl || '' };
            created = true;
            write('  Liste angelegt.');
        } else {
            write('  Liste vorhanden.');
        }
        const colResult = await addMissingColumns(G, token, siteId, list.id, pack.columns, write);
        if (colResult.added) write('  ' + colResult.added + ' Spalte(n) ergänzt.');
        else write('  Spalten vollständig.');
        lists[pack.key] = {
            id: String(list.id),
            webUrl: list.webUrl || '',
            title: pack.title,
            created: created
        };
    }

    return { titles, lists };
}

/**
 * Health für Lehrer-/Klassen-Listen.
 * @param {{
 *   token: string,
 *   siteId: string,
 *   graph: { graphPathSite: Function, graphJson: Function },
 *   titles?: { lehrer?: string, klassen?: string }
 * }} opts
 */
/**
 * @param {{ graphPathSite: Function, graphJson: Function }} G
 * @param {string} token
 * @param {string} siteId
 * @param {string} listId
 * @returns {Promise<{ count: number, readable: boolean, error: string }>}
 */
async function countListItemsReadable(G, token, siteId, listId) {
    try {
        const path =
            G.graphPathSite(siteId) +
            '/lists/' +
            encodeURIComponent(listId) +
            '/items?$select=id&$top=200';
        const data = await G.graphJson('GET', path, token, undefined, 'v1.0');
        let n = ((data && data.value) || []).length;
        let next = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
        while (next && n < 2000) {
            const more = await G.graphJson('GET', next, token, undefined, 'v1.0');
            n += ((more && more.value) || []).length;
            next = more && more['@odata.nextLink'] ? more['@odata.nextLink'] : '';
        }
        return { count: n, readable: true, error: '' };
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e);
        return { count: 0, readable: false, error: msg };
    }
}

export async function healthSchoolStammdatenShellLists(opts) {
    const o = opts || {};
    const G = o.graph;
    const token = o.token;
    const siteId = o.siteId;
    const defaults = defaultStammdatenSpoListTitles();
    const titles = {
        lehrer: String((o.titles && o.titles.lehrer) || defaults.lehrer).trim() || defaults.lehrer,
        klassen: String((o.titles && o.titles.klassen) || defaults.klassen).trim() || defaults.klassen
    };
    const required = {
        [titles.lehrer]: STAMMDATEN_SPO_REQUIRED_COLUMNS.lehrer,
        [titles.klassen]: STAMMDATEN_SPO_REQUIRED_COLUMNS.klassen
    };
    /** @type {{ title: string, ok: boolean, detail: string, key: string, itemCount?: number }[]} */
    const report = [];
    for (const key of ['lehrer', 'klassen']) {
        const title = titles[key];
        const list = await findListByTitle(G, token, siteId, title);
        if (!list) {
            report.push({ title, ok: false, detail: 'fehlt', key, itemCount: 0 });
            continue;
        }
        const colsPath =
            G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(list.id) + '/columns?$top=200';
        const colsData = await G.graphJson('GET', colsPath, token, undefined, 'v1.0');
        const names = new Set(((colsData && colsData.value) || []).map((c) => c.name));
        const need = required[title] || [];
        const missing = need.filter((n) => !names.has(n));
        if (missing.length) {
            report.push({
                title,
                ok: false,
                detail: 'fehlt: ' + missing.join(', '),
                key,
                itemCount: 0
            });
            continue;
        }
        const items = await countListItemsReadable(G, token, siteId, String(list.id));
        if (!items.readable) {
            report.push({
                title,
                ok: false,
                detail: 'Spalten ok, aber Lesen fehlgeschlagen (' + (items.error || 'Rechte?') + ')',
                key,
                itemCount: 0
            });
            continue;
        }
        const empty = items.count === 0;
        report.push({
            title,
            ok: !empty,
            detail: empty
                ? '0 Einträge – bitte „Nur Stammdaten befüllen“ oder Paket anlegen'
                : items.count + ' Einträge',
            key,
            itemCount: items.count
        });
    }
    return report;
}

async function findListByTitle(G, token, siteId, listTitle) {
    const title = String(listTitle || '').trim();
    const path =
        G.graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await G.graphJson('GET', path, token, undefined, 'v1.0');
    return ((data && data.value) || [])[0] || null;
}

/**
 * @param {{ graphPathSite: Function, graphJson: Function }} G
 * @param {string} token
 * @param {string} siteId
 * @param {string} listId
 */
async function fetchAllListItems(G, token, siteId, listId) {
    let path =
        G.graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(listId) +
        '/items?$expand=fields&$top=100';
    const out = [];
    while (path) {
        const data = await G.graphJson(
            'GET',
            path.indexOf('http') === 0 ? path : path,
            token,
            undefined,
            'v1.0'
        );
        const rows = (data && data.value) || [];
        for (let i = 0; i < rows.length; i++) out.push(rows[i]);
        path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
    }
    return out;
}

const SPO_SELECT_LEHRER = 'Id,Title,LehrerCode,Vorname,Nachname,EMail,UPN';
const SPO_SELECT_KLASSEN =
    'Id,Title,KlassenCode,Abschlussjahr,KVName,KVEmail,M365GroupId,M365GroupMail,M365GroupName';

/**
 * SharePoint-REST (wie Browser-UI) – Fallback wenn Graph leer/403 liefert.
 * @param {{ getGraphToken: Function, getSpoRequestDigest: Function, spoRestFetch: Function }} G
 * @param {string} siteWebUrl
 * @param {string} listTitle
 * @param {string} [selectFields]
 * @returns {Promise<{ fields: Record<string, unknown> }[]>}
 */
async function fetchAllListItemsViaSpoRest(G, siteWebUrl, listTitle, selectFields) {
    const url = String(siteWebUrl || '').trim().replace(/\/+$/, '');
    const title = String(listTitle || '').trim();
    if (!url || !title) return [];
    if (typeof G.getSpoRequestDigest !== 'function' || typeof G.spoRestFetch !== 'function') {
        throw new Error('SPO-REST-Hilfen fehlen.');
    }
    let host = '';
    try {
        host = new URL(url).hostname;
    } catch {
        throw new Error('Site-URL ungültig für SPO-REST.');
    }
    const spoScopes = [
        ['https://' + host + '/AllSites.Read', 'https://graph.microsoft.com/User.Read'],
        ['https://' + host + '/AllSites.Write', 'https://graph.microsoft.com/User.Read']
    ];
    let spoToken = '';
    let lastErr = '';
    for (let i = 0; i < spoScopes.length; i++) {
        try {
            if (typeof window !== 'undefined' && typeof window.ms365AuthAcquireTokenSilent === 'function') {
                try {
                    spoToken = await window.ms365AuthAcquireTokenSilent(spoScopes[i]);
                } catch {
                    spoToken = await G.getGraphToken(spoScopes[i]);
                }
            } else {
                spoToken = await G.getGraphToken(spoScopes[i]);
            }
            if (spoToken) break;
        } catch (e) {
            lastErr = e && e.message ? String(e.message) : String(e);
        }
    }
    if (!spoToken) throw new Error('SharePoint-Token fehlt (' + (lastErr || 'AllSites.Read') + ')');

    const digest = await G.getSpoRequestDigest(url, spoToken);
    const esc = title.replace(/'/g, "''");
    const select = String(selectFields || SPO_SELECT_KLASSEN).trim();
    /** @type {{ fields: Record<string, unknown> }[]} */
    const out = [];
    let api =
        "/_api/web/lists/getbytitle('" +
        esc +
        "')/items?$top=100&$select=" +
        select;
    while (api) {
        const res = await G.spoRestFetch(url, spoToken, digest, 'GET', api);
        if (!res.ok) {
            throw new Error(
                'SPO „' + title + '": HTTP ' + res.status + ' ' + String(res.text || '').slice(0, 180)
            );
        }
        const rows = (res.data && res.data.value) || [];
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i] || {};
            // odata=nometadata: Felder liegen flach auf dem Item
            out.push({ fields: row, id: row.Id != null ? row.Id : row.id });
        }
        const next = String(
            (res.data && (res.data['@odata.nextLink'] || res.data['odata.nextLink'])) || ''
        ).trim();
        api = next || '';
    }
    return out;
}

function mapRowsFromItems(items, mapFn) {
    const seen = new Set();
    const out = [];
    for (let i = 0; i < items.length; i++) {
        const mapped = mapFn(items[i] && items[i].fields);
        if (!mapped) continue;
        const key = String(mapped.code || mapped.email || '')
            .trim()
            .toUpperCase();
        if (!key || seen.has(key)) continue;
        seen.add(key);
        out.push(mapped);
    }
    return out;
}

/**
 * @param {{ getGraphToken: Function, getSpoRequestDigest?: Function, spoRestFetch?: Function }} G
 * @param {string} siteWebUrl
 */
function canUseSpoRestFallback(G, siteWebUrl) {
    return !!(
        String(siteWebUrl || '').trim() &&
        G &&
        typeof G.getGraphToken === 'function' &&
        typeof G.getSpoRequestDigest === 'function' &&
        typeof G.spoRestFetch === 'function'
    );
}

/**
 * Lehrer + Klassen live von der Schul-Site lesen (Betriebs-Wahrheit).
 *
 * @param {{
 *   token: string,
 *   siteId: string,
 *   siteWebUrl?: string,
 *   graph: { graphPathSite: Function, graphJson: Function, getGraphToken?: Function, getSpoRequestDigest?: Function, spoRestFetch?: Function },
 *   titles?: { lehrer?: string, klassen?: string }
 * }} opts
 * @returns {Promise<{
 *   teachers: StammdatenTeacher[],
 *   classes: StammdatenClass[],
 *   lists: Record<string, { id: string, title: string, webUrl: string }>,
 *   missing: string[],
 *   readErrors: string[],
 *   source?: string
 * }>}
 */
export async function fetchSchoolStammdatenFromSite(opts) {
    const o = opts || {};
    const G = o.graph;
    const token = o.token;
    const siteId = o.siteId;
    const siteWebUrl = String(o.siteWebUrl || '').trim();
    if (!G || !token || !siteId) {
        throw new Error('fetchSchoolStammdatenFromSite: token/siteId/graph fehlen.');
    }
    const titles = resolveStammdatenSpoListTitles(o.titles || defaultStammdatenSpoListTitles());
    /** @type {Record<string, { id: string, title: string, webUrl: string }>} */
    const lists = {};
    /** @type {string[]} */
    const missing = [];
    /** @type {string[]} */
    const readErrors = [];
    /** @type {StammdatenTeacher[]} */
    let teachers = [];
    /** @type {StammdatenClass[]} */
    let classes = [];
    let usedSpoFallback = false;

    try {
        const lehrerList = await findListByTitle(G, token, siteId, titles.lehrer);
        if (!lehrerList || !lehrerList.id) {
            missing.push(titles.lehrer);
        } else {
            lists.lehrer = {
                id: String(lehrerList.id),
                title: titles.lehrer,
                webUrl: lehrerList.webUrl || ''
            };
            try {
                const rows = await fetchAllListItems(G, token, siteId, String(lehrerList.id));
                teachers = mapRowsFromItems(rows, mapLehrerFieldsToTeacher);
                teachers.sort((a, b) => String(a.name || '').localeCompare(String(b.name || ''), 'de'));
            } catch (e) {
                const msg = e && e.message ? String(e.message) : String(e);
                const st = e && e.status != null ? e.status : '';
                readErrors.push(
                    titles.lehrer +
                        ': ' +
                        (st === 403 || /403|accessdenied|forbidden/i.test(msg)
                            ? 'kein Leserecht (403) – Einreicher-Gruppe/Person braucht Lesen auf dieser Liste'
                            : msg)
                );
            }
        }
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e);
        if (/403|accessdenied|forbidden/i.test(msg) || (e && e.status === 403)) {
            readErrors.push(titles.lehrer + ': Liste nicht auffindbar/lesbar (403)');
        } else {
            throw e;
        }
    }

    try {
        const klassenList = await findListByTitle(G, token, siteId, titles.klassen);
        if (!klassenList || !klassenList.id) {
            missing.push(titles.klassen);
        } else {
            lists.klassen = {
                id: String(klassenList.id),
                title: titles.klassen,
                webUrl: klassenList.webUrl || ''
            };
            try {
                const rows = await fetchAllListItems(G, token, siteId, String(klassenList.id));
                classes = mapRowsFromItems(rows, mapKlassenFieldsToClass);
                classes.sort((a, b) => String(a.code || '').localeCompare(String(b.code || ''), 'de'));
            } catch (e) {
                const msg = e && e.message ? String(e.message) : String(e);
                const st = e && e.status != null ? e.status : '';
                readErrors.push(
                    titles.klassen +
                        ': ' +
                        (st === 403 || /403|accessdenied|forbidden/i.test(msg)
                            ? 'kein Leserecht (403) – Einreicher braucht Lesen auf „Klassen“ (IT: Nur Berechtigungen erneut)'
                            : msg)
                );
            }
        }
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e);
        if (/403|accessdenied|forbidden/i.test(msg) || (e && e.status === 403)) {
            readErrors.push(titles.klassen + ': Liste nicht auffindbar/lesbar (403)');
        } else {
            throw e;
        }
    }

    // Graph leer/403, aber Lehrer sieht die Liste im Browser → SPO-REST (AllSites.Read)
    const needLehrerSpo =
        canUseSpoRestFallback(G, siteWebUrl) &&
        (!teachers.length || readErrors.some((x) => String(x).indexOf(titles.lehrer) === 0));
    const needKlassenSpo =
        canUseSpoRestFallback(G, siteWebUrl) &&
        (!classes.length || readErrors.some((x) => String(x).indexOf(titles.klassen) === 0));

    if (needLehrerSpo || needKlassenSpo) {
        const cleared = [];
        if (needLehrerSpo) {
            try {
                const rows = await fetchAllListItemsViaSpoRest(
                    G,
                    siteWebUrl,
                    titles.lehrer,
                    SPO_SELECT_LEHRER
                );
                const mapped = mapRowsFromItems(rows, mapLehrerFieldsToTeacher);
                if (mapped.length) {
                    teachers = mapped.sort((a, b) =>
                        String(a.name || '').localeCompare(String(b.name || ''), 'de')
                    );
                    usedSpoFallback = true;
                    if (!lists.lehrer) {
                        lists.lehrer = { id: '', title: titles.lehrer, webUrl: siteWebUrl };
                    }
                    for (let i = missing.length - 1; i >= 0; i--) {
                        if (missing[i] === titles.lehrer) missing.splice(i, 1);
                    }
                    cleared.push(titles.lehrer);
                }
            } catch (e) {
                const msg = e && e.message ? String(e.message) : String(e);
                if (!readErrors.some((x) => String(x).indexOf(titles.lehrer) === 0)) {
                    readErrors.push(titles.lehrer + ' (SPO): ' + msg);
                }
            }
        }
        if (needKlassenSpo) {
            try {
                const rows = await fetchAllListItemsViaSpoRest(
                    G,
                    siteWebUrl,
                    titles.klassen,
                    SPO_SELECT_KLASSEN
                );
                const mapped = mapRowsFromItems(rows, mapKlassenFieldsToClass);
                if (mapped.length) {
                    classes = mapped.sort((a, b) =>
                        String(a.code || '').localeCompare(String(b.code || ''), 'de')
                    );
                    usedSpoFallback = true;
                    if (!lists.klassen) {
                        lists.klassen = { id: '', title: titles.klassen, webUrl: siteWebUrl };
                    }
                    for (let i = missing.length - 1; i >= 0; i--) {
                        if (missing[i] === titles.klassen) missing.splice(i, 1);
                    }
                    cleared.push(titles.klassen);
                }
            } catch (e) {
                const msg = e && e.message ? String(e.message) : String(e);
                if (!readErrors.some((x) => String(x).indexOf(titles.klassen) === 0)) {
                    readErrors.push(titles.klassen + ' (SPO): ' + msg);
                }
            }
        }
        if (cleared.length) {
            for (let i = readErrors.length - 1; i >= 0; i--) {
                const line = String(readErrors[i] || '');
                if (cleared.some((t) => line.indexOf(t) === 0)) readErrors.splice(i, 1);
            }
        }
    }

    // Fehlende M365-Gruppenfelder aus lokalen Klassen-Teams nachziehen
    const enrichedClasses = (classes || []).map((c) => enrichClassWithM365Group(c));

    return {
        teachers,
        classes: enrichedClasses,
        lists,
        missing,
        readErrors,
        source: usedSpoFallback ? 'spo-rest' : 'graph'
    };
}

/**
 * Für Planer/Matura: Live-Listen lesen und mit lokalem Fallback mergen.
 * Fehler → lokaler Stand unverändert (kein Hard-Fail).
 *
 * @param {{
 *   token: string,
 *   siteId: string,
 *   siteWebUrl?: string,
 *   graph: { graphPathSite: Function, graphJson: Function, getGraphToken?: Function, getSpoRequestDigest?: Function, spoRestFetch?: Function },
 *   local?: { teachers?: unknown[], classes?: unknown[], students?: unknown[] },
 *   titles?: { lehrer?: string, klassen?: string }
 * }} opts
 */
export async function hydrateStammdatenFromLiveSite(opts) {
    const o = opts || {};
    const local = o.local || { teachers: [], classes: [], students: [] };
    try {
        const live = await fetchSchoolStammdatenFromSite({
            token: o.token,
            siteId: o.siteId,
            siteWebUrl: o.siteWebUrl || '',
            graph: o.graph,
            titles: o.titles || resolveStammdatenSpoListTitlesFromWindow()
        });
        const readErrors = Array.isArray(live.readErrors) ? live.readErrors : [];
        const ok = readErrors.length === 0 || !!(live.classes && live.classes.length);
        return {
            stammdaten: mergeStammdatenPreferLive(local, live),
            live,
            ok,
            error: readErrors.length ? readErrors.join(' · ') : ''
        };
    } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        const denied = /403|accessdenied|forbidden/i.test(msg) || (e && e.status === 403);
        return {
            stammdaten: {
                teachers: Array.isArray(local.teachers) ? local.teachers : [],
                classes: Array.isArray(local.classes) ? local.classes : [],
                students: Array.isArray(local.students) ? local.students : []
            },
            live: null,
            ok: false,
            error: denied
                ? 'Kein Zugriff auf Stammdaten-Listen (403). IT: Einreicher-Rechte auf Klassen/Lehrerinnen erneut setzen.'
                : msg
        };
    }
}

async function addMissingColumns(G, token, siteId, listId, defs, write) {
    const colsPath = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns?$top=200';
    const colsData = await G.graphJson('GET', colsPath, token, undefined, 'v1.0');
    const existing = new Set(
        ((colsData && colsData.value) || [])
            .map((c) => String((c && c.name) || '').trim())
            .filter(Boolean)
    );
    const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
    let added = 0;
    for (let i = 0; i < defs.length; i++) {
        const def = defs[i];
        if (String(def.name || '') === 'Title' && existing.has('Title')) continue;
        if (existing.has(def.name)) continue;
        await G.graphJson('POST', base, token, stammdatenColumnToGraphBody(def), 'v1.0');
        added++;
        write('  + Spalte ' + def.name);
        if (typeof G.sleep === 'function') await G.sleep(120);
        else await new Promise((r) => setTimeout(r, 120));
        existing.add(def.name);
    }
    return { added };
}

function sleepGraph(G, ms) {
    if (typeof G.sleep === 'function') return G.sleep(ms);
    return new Promise((r) => setTimeout(r, ms));
}

function splitTeacherName(name) {
    const n = String(name || '').trim();
    if (!n) return { vorname: '', nachname: '' };
    const i = n.indexOf(' ');
    if (i < 0) return { vorname: '', nachname: n };
    return { vorname: n.slice(0, i).trim(), nachname: n.slice(i + 1).trim() };
}

/**
 * Lokale Stammdaten (Lehrer/Klassen) aus Browser-API.
 * @param {Window|null|undefined} [win]
 * @returns {{ teachers: unknown[], classes: unknown[] }}
 */
export function loadLocalSchoolStammdatenFromWindow(win) {
    const w = win || (typeof window !== 'undefined' ? window : null);
    if (!w) return { teachers: [], classes: [] };
    try {
        if (typeof w.ms365TenantSettingsLoad === 'function') {
            const s = w.ms365TenantSettingsLoad() || {};
            return {
                teachers: Array.isArray(s.teachers) ? s.teachers : [],
                classes: Array.isArray(s.classes) ? s.classes : []
            };
        }
        if (w.ms365AppDataV2 && typeof w.ms365AppDataV2.getCore === 'function') {
            const core = w.ms365AppDataV2.getCore() || {};
            const year =
                typeof w.ms365AppDataV2.getCurrentYear === 'function'
                    ? w.ms365AppDataV2.getCurrentYear() || {}
                    : {};
            return {
                teachers: Array.isArray(core.teachers) ? core.teachers : [],
                classes: Array.isArray(year.classes) ? year.classes : []
            };
        }
    } catch {
        /* ignore */
    }
    return { teachers: [], classes: [] };
}

/**
 * App-Lehrer → SharePoint-Felder Lehrerinnen.
 * @param {Record<string, unknown>|null|undefined} t
 */
export function teacherToLehrerListFields(t) {
    const row = t && typeof t === 'object' ? t : {};
    const email = String(row.email || row.upn || '')
        .trim()
        .toLowerCase();
    const name = String(row.name || '').trim() || String(row.code || '').trim() || email;
    const code = String(row.code || '').trim() || email || name;
    const parts = splitTeacherName(name);
    return {
        Title: name,
        Vorname: parts.vorname,
        Nachname: parts.nachname,
        LehrerCode: code,
        EMail: email,
        UPN: String(row.upn || email)
            .trim()
            .toLowerCase()
    };
}

/**
 * App-Klasse → SharePoint-Felder Klassen (inkl. KV-Text).
 * @param {Record<string, unknown>|null|undefined} c
 */
export function classToKlassenListFields(c) {
    const enriched = enrichClassWithM365Group(c && typeof c === 'object' ? c : {});
    const code = String(enriched.code || '').trim();
    const name = String(enriched.name || '').trim() || code;
    const headEmail = String(enriched.headEmail || enriched.kvEmail || '')
        .trim()
        .toLowerCase();
    const fields = {
        Title: name || code,
        KlassenCode: code || name,
        Abschlussjahr: String(enriched.year || '').trim(),
        KVName: String(enriched.headName || '').trim(),
        KVEmail: headEmail
    };
    const gid = String(enriched.m365GroupId || '').trim();
    const gmail = String(enriched.m365GroupMail || '')
        .trim()
        .toLowerCase();
    const gname = String(enriched.m365GroupName || '').trim();
    if (gid) fields.M365GroupId = gid;
    if (gmail) fields.M365GroupMail = gmail;
    if (gname) fields.M365GroupName = gname;
    return fields;
}

function textFieldsEqual(prev, next, keys) {
    for (let i = 0; i < keys.length; i++) {
        const k = keys[i];
        if (String((prev && prev[k]) || '') !== String((next && next[k]) || '')) return false;
    }
    return true;
}

/**
 * Zeilen per Schlüssel abgleichen (anlegen / patchen / orphan entfernen).
 * @param {{
 *   graph: { graphPathSite: Function, graphJson: Function, sleep?: Function },
 *   token: string,
 *   siteId: string,
 *   listId: string,
 *   rows: unknown[],
 *   keyFromRow: (row: unknown) => string,
 *   keyFromItem: (fields: Record<string, unknown>) => string,
 *   fieldsFromRow: (row: unknown) => Promise<Record<string, unknown>>|Record<string, unknown>,
 *   compareKeys: string[],
 *   removeOrphans?: boolean,
 *   write?: (msg: string) => void,
 *   label?: string
 * }} params
 */
async function syncKeyedRowsToList(params) {
    const G = params.graph;
    const write = typeof params.write === 'function' ? params.write : () => {};
    const label = params.label || 'Zeilen';
    const items = await fetchAllListItems(G, params.token, params.siteId, params.listId);
    const byKey = new Map();
    for (let i = 0; i < items.length; i++) {
        const it = items[i];
        const fields = (it && it.fields) || {};
        const key = params.keyFromItem(fields);
        if (!key || byKey.has(key)) continue;
        byKey.set(key, it);
    }

    const desired = new Set();
    let added = 0;
    let updated = 0;
    let unchanged = 0;
    const itemsPath =
        G.graphPathSite(params.siteId) + '/lists/' + encodeURIComponent(params.listId) + '/items';

    for (let i = 0; i < params.rows.length; i++) {
        const row = params.rows[i];
        const key = params.keyFromRow(row);
        if (!key) continue;
        desired.add(key);
        const nextFields = await Promise.resolve(params.fieldsFromRow(row));
        const existing = byKey.get(key);
        if (!existing) {
            await G.graphJson('POST', itemsPath, params.token, { fields: nextFields }, 'v1.0');
            added++;
            if (added % 25 === 0) write('  … ' + added + ' neu');
            await sleepGraph(G, 80);
            continue;
        }
        const prev = existing.fields || {};
        if (textFieldsEqual(prev, nextFields, params.compareKeys)) {
            unchanged++;
            continue;
        }
        const itemId = existing.id != null ? String(existing.id) : '';
        if (!itemId) continue;
        await G.graphJson(
            'PATCH',
            itemsPath + '/' + encodeURIComponent(itemId) + '/fields',
            params.token,
            nextFields,
            'v1.0'
        );
        updated++;
        await sleepGraph(G, 80);
    }

    let removed = 0;
    if (params.removeOrphans) {
        for (const [key, it] of byKey.entries()) {
            if (desired.has(key)) continue;
            const itemId = it && it.id != null ? String(it.id) : '';
            if (!itemId) continue;
            try {
                await G.graphJson(
                    'DELETE',
                    itemsPath + '/' + encodeURIComponent(itemId),
                    params.token,
                    undefined,
                    'v1.0'
                );
                removed++;
                await sleepGraph(G, 60);
            } catch {
                /* ignore */
            }
        }
    }

    write(
        '  ' +
            label +
            ': ' +
            added +
            ' neu, ' +
            updated +
            ' aktualisiert, ' +
            unchanged +
            ' unverändert' +
            (removed ? ', ' + removed + ' entfernt' : '') +
            '.'
    );
    return { added, updated, unchanged, removed };
}

/**
 * Lehrer- und Klassenliste aus lokalen Stammdaten befüllen (idempotent).
 * Schema wird zuvor über ensureSchoolStammdatenShellLists sichergestellt.
 *
 * @param {{
 *   token: string,
 *   siteId: string,
 *   siteWebUrl?: string,
 *   graph: {
 *     graphPathSite: Function,
 *     graphJson: Function,
 *     sleep?: Function,
 *     createSitePersonResolver?: Function,
 *     applySinglePersonLookupField?: Function
 *   },
 *   write?: (msg: string) => void,
 *   titles?: { lehrer?: string, klassen?: string },
 *   teachers?: unknown[],
 *   classes?: unknown[],
 *   personFields?: boolean,
 *   removeOrphans?: boolean
 * }} opts
 */
export async function seedSchoolStammdatenListsFromLocal(opts) {
    const o = opts || {};
    const write = typeof o.write === 'function' ? o.write : () => {};
    const G = o.graph;
    const token = o.token;
    const siteId = o.siteId;
    if (!G || !token || !siteId) {
        throw new Error('seedSchoolStammdatenListsFromLocal: token/siteId/graph fehlen.');
    }

    const shell = await ensureSchoolStammdatenShellLists({
        token,
        siteId,
        graph: G,
        write,
        titles: o.titles
    });

    const local =
        Array.isArray(o.teachers) || Array.isArray(o.classes)
            ? {
                  teachers: Array.isArray(o.teachers) ? o.teachers : [],
                  classes: Array.isArray(o.classes) ? o.classes : []
              }
            : loadLocalSchoolStammdatenFromWindow();

    const teachers = (local.teachers || []).filter(
        (t) => t && (String(t.code || '').trim() || String(t.email || '').trim() || String(t.name || '').trim())
    );
    const classes = (local.classes || []).filter(
        (c) => c && (String(c.code || '').trim() || String(c.name || '').trim())
    );

    write(
        'Stammdaten lokal: ' + teachers.length + ' Lehrkräfte, ' + classes.length + ' Klassen – schreibe nach SharePoint …'
    );
    if (!teachers.length && !classes.length) {
        write('  ! Keine lokalen Stammdaten – Listen bleiben leer. Bitte zuerst in den Stammdaten pflegen.');
        return { shell, teachers: 0, classes: 0, lehrer: null, klassen: null, skipped: true };
    }

    const withPersons = o.personFields !== false;
    let resolver = null;
    if (withPersons && o.siteWebUrl && typeof G.createSitePersonResolver === 'function') {
        write('Personenfelder: Benutzer auf der Site auflösen …');
        try {
            resolver = await G.createSitePersonResolver(o.siteWebUrl, token, siteId, {
                write,
                ensureMissing: true
            });
        } catch (e) {
            write('  Hinweis Personen: ' + (e && e.message ? e.message : e));
        }
    }

    /** @type {{ added: number, updated: number, unchanged: number, removed: number }|null} */
    let lehrerSync = null;
    if (teachers.length && shell.lists.lehrer) {
        write('Befülle „' + shell.titles.lehrer + '" …');
        const compareKeys = ['Title', 'Vorname', 'Nachname', 'LehrerCode', 'EMail', 'UPN'];
        lehrerSync = await syncKeyedRowsToList({
            graph: G,
            token,
            siteId,
            listId: shell.lists.lehrer.id,
            rows: teachers,
            keyFromRow: (t) => {
                const code = String((t && t.code) || '')
                    .trim()
                    .toLowerCase();
                if (code) return code;
                return String((t && t.email) || '')
                    .trim()
                    .toLowerCase();
            },
            keyFromItem: (f) => {
                const code = String((f && f.LehrerCode) || '')
                    .trim()
                    .toLowerCase();
                if (code) return code;
                return String((f && (f.EMail || f.Email)) || '')
                    .trim()
                    .toLowerCase();
            },
            fieldsFromRow: async (t) => {
                const fields = teacherToLehrerListFields(t);
                if (!resolver || !withPersons) return fields;
                const email = String((t && t.email) || '').trim();
                if (!email) return fields;
                const id = await resolver.lookupIdForEmail(email);
                if (id && typeof G.applySinglePersonLookupField === 'function') {
                    G.applySinglePersonLookupField(fields, 'Lehrkraft', id);
                }
                return fields;
            },
            compareKeys,
            removeOrphans: !!o.removeOrphans,
            write,
            label: 'Lehrerinnen'
        });
    } else if (!teachers.length) {
        write('  Lehrer: keine lokalen Zeilen – übersprungen.');
    }

    /** @type {{ added: number, updated: number, unchanged: number, removed: number }|null} */
    let klassenSync = null;
    if (classes.length && shell.lists.klassen) {
        write('Befülle „' + shell.titles.klassen + '" …');
        const compareKeys = [
            'Title',
            'KlassenCode',
            'Abschlussjahr',
            'KVName',
            'KVEmail',
            'M365GroupId',
            'M365GroupMail',
            'M365GroupName'
        ];
        /** @type {Map<string, { mail: string, displayName: string }>} */
        const groupMailCache = new Map();
        async function resolveGroupMail(groupId) {
            const gid = String(groupId || '').trim();
            if (!gid) return { mail: '', displayName: '' };
            if (groupMailCache.has(gid)) return groupMailCache.get(gid) || { mail: '', displayName: '' };
            try {
                const g = await G.graphJson(
                    'GET',
                    '/groups/' + encodeURIComponent(gid) + '?$select=id,displayName,mail,mailNickname',
                    token,
                    undefined,
                    'v1.0'
                );
                const info = {
                    mail: String((g && g.mail) || '')
                        .trim()
                        .toLowerCase(),
                    displayName: String((g && g.displayName) || '').trim()
                };
                groupMailCache.set(gid, info);
                return info;
            } catch {
                const empty = { mail: '', displayName: '' };
                groupMailCache.set(gid, empty);
                return empty;
            }
        }
        async function lookupIdForGroup(groupId) {
            const gid = String(groupId || '').trim();
            if (!gid || !o.siteWebUrl || typeof G.spoEnsureUser !== 'function') return 0;
            const logon = entraGroupLogonNameForSite(gid);
            if (!logon) return 0;
            try {
                let spoToken = '';
                let digest = '';
                let host = '';
                try {
                    host = new URL(String(o.siteWebUrl)).hostname;
                } catch {
                    host = '';
                }
                if (!host) return 0;
                spoToken = await G.getGraphToken(['https://' + host + '/AllSites.Write']);
                digest = await G.getSpoRequestDigest(o.siteWebUrl, spoToken);
                const principal = await G.spoEnsureUser(o.siteWebUrl, spoToken, digest, logon);
                return principal && principal.id ? Number(principal.id) : 0;
            } catch (e) {
                write(
                    '  M365-Gruppe ' +
                        gid.slice(0, 8) +
                        '… nicht auf Site verknüpfbar: ' +
                        (e && e.message ? e.message : e)
                );
                return 0;
            }
        }
        klassenSync = await syncKeyedRowsToList({
            graph: G,
            token,
            siteId,
            listId: shell.lists.klassen.id,
            rows: classes,
            keyFromRow: (c) => {
                const code = String((c && c.code) || '')
                    .trim()
                    .toLowerCase();
                if (code) return code;
                return String((c && c.name) || '')
                    .trim()
                    .toLowerCase();
            },
            keyFromItem: (f) => {
                const code = String((f && f.KlassenCode) || '')
                    .trim()
                    .toLowerCase();
                if (code) return code;
                return String((f && f.Title) || '')
                    .trim()
                    .toLowerCase();
            },
            fieldsFromRow: async (c) => {
                const enriched = enrichClassWithM365Group(c);
                if (enriched.m365GroupId && !enriched.m365GroupMail) {
                    const info = await resolveGroupMail(enriched.m365GroupId);
                    if (info.mail) enriched.m365GroupMail = info.mail;
                    if (info.displayName && !enriched.m365GroupName) {
                        enriched.m365GroupName = info.displayName;
                    }
                }
                const fields = classToKlassenListFields(enriched);
                if (resolver && withPersons) {
                    const kvEm = String(enriched.headEmail || enriched.kvEmail || '').trim();
                    if (kvEm) {
                        const id = await resolver.lookupIdForEmail(kvEm);
                        if (id && typeof G.applySinglePersonLookupField === 'function') {
                            G.applySinglePersonLookupField(fields, 'Klassenvorstand', id);
                        }
                    }
                }
                if (enriched.m365GroupId && typeof G.applySinglePersonLookupField === 'function') {
                    const gidLookup = await lookupIdForGroup(enriched.m365GroupId);
                    if (gidLookup) G.applySinglePersonLookupField(fields, 'M365Gruppe', gidLookup);
                }
                return fields;
            },
            compareKeys,
            removeOrphans: !!o.removeOrphans,
            write,
            label: 'Klassen'
        });
    } else if (!classes.length) {
        write('  Klassen: keine lokalen Zeilen – übersprungen.');
    }

    write('Stammdaten-Befüllung fertig.');
    return {
        shell,
        teachers: teachers.length,
        classes: classes.length,
        lehrer: lehrerSync,
        klassen: klassenSync,
        skipped: false
    };
}
