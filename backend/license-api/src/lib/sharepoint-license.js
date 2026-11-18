'use strict';

const { getConfig } = require('./config');
const { getOperatorGraphToken } = require('./msal-app-only');
const { findFieldsForTenant, parseDomainList, parseContactEmails } = require('./evaluate-license');
const {
    ENTITLEMENTS_FIELD,
    parseEntitlementsRaw,
    serializeEntitlements,
    resolveFeatureEntitlements,
    defaultNewSchoolEntitlements,
    allPackageIds
} = require('./feature-catalog');
const {
    SCHOOL_STAMMDATEN_FIELDS,
    assertTenantIdAllowed,
    isPendingTenantId,
    makePendingTenantId,
    normalizeTenantIdInput
} = require('./school-stammdaten-catalog');

const GRAPH = 'https://graph.microsoft.com/v1.0';

/** SharePoint-Site der Schule (nur URL-Zeiger, keine Stammdaten-Inhalte). */
const TOOL_SITE_URL_FIELD = 'ToolSiteUrl';

/**
 * URL-Wert aus Listenzeile lesen (Text oder Hyperlink-Objekt).
 * @param {Record<string, unknown>|null|undefined} fields
 * @returns {string}
 */
function looksLikeToolSiteUrlFieldName(name) {
    const n = String(name || '').trim();
    if (!n) return false;
    const decoded = n.replace(/_x002d_/gi, '-').replace(/_x0020_/gi, ' ');
    return (
        /tool.?site.?url/i.test(n) ||
        /tool.?site.?url/i.test(decoded) ||
        /schul.?sharepoint/i.test(n) ||
        /schul.?sharepoint/i.test(decoded)
    );
}

function pickToolSiteUrlFromFields(fields) {
    const f = fields && typeof fields === 'object' ? fields : {};
    const keys = Object.keys(f);
    const preferred = [TOOL_SITE_URL_FIELD, 'toolSiteUrl'];
    for (let i = 0; i < preferred.length; i++) {
        const v = normalizeToolSiteUrlValue(f[preferred[i]]);
        if (v) return v;
    }
    for (let i = 0; i < keys.length; i++) {
        const k = keys[i];
        if (!looksLikeToolSiteUrlFieldName(k)) continue;
        const v = normalizeToolSiteUrlValue(f[k]);
        if (v) return v;
    }
    return '';
}

/**
 * @param {unknown} raw
 * @returns {string}
 */
function normalizeToolSiteUrlValue(raw) {
    if (raw == null || raw === '') return '';
    if (typeof raw === 'object') {
        const o = /** @type {Record<string, unknown>} */ (raw);
        const url = o.Url != null ? o.Url : o.url;
        return String(url == null ? '' : url).trim();
    }
    return String(raw).trim();
}

/**
 * Interne Spaltennamen für Tool-Site-URL in der Liste finden.
 * @param {Array<Record<string, unknown>>} rawColumns
 * @returns {{ name: string, kind: 'text'|'hyperlink' }|null}
 */
function findToolSiteUrlColumnMeta(rawColumns) {
    const rows = Array.isArray(rawColumns) ? rawColumns : [];
    let fallback = null;
    for (let i = 0; i < rows.length; i++) {
        const c = rows[i];
        if (!c || c.hidden) continue;
        const name = String(c.name || '').trim();
        const display = String(c.displayName || '').trim();
        if (!name) continue;
        const nameHit = name.toLowerCase() === TOOL_SITE_URL_FIELD.toLowerCase();
        const fuzzy =
            looksLikeToolSiteUrlFieldName(name) ||
            looksLikeToolSiteUrlFieldName(display) ||
            (/sharepoint/i.test(display) && /tool|site|schul/i.test(display));
        if (!nameHit && !fuzzy) continue;
        const kind = c.hyperlinkOrPicture ? 'hyperlink' : 'text';
        const meta = { name: name, kind: kind };
        if (nameHit) return meta;
        if (!fallback) fallback = meta;
    }
    return fallback;
}

/**
 * Graph-Feldwert für Tool-Site-URL je Spaltentyp.
 * @param {string} url
 * @param {'text'|'hyperlink'} [kind]
 */
function toolSiteUrlFieldValue(url, kind) {
    const u = String(url || '').trim();
    if (kind === 'hyperlink') {
        return u ? { Url: u, Description: u } : null;
    }
    return u;
}

/** Feste Kernspalten – nicht als „Extra“ in der Admin-UI. */
const CORE_FIELD_NAMES = new Set([
    'id',
    'title',
    'tenantid',
    'primarydomain',
    'additionaldomains',
    'status',
    'validuntil',
    'contactemail',
    'notes',
    'itcontacts',
    'licenseentitlements',
    'toolsiteurl'
]);

/** SharePoint-/Graph-Systemspalten, die nie als Extra gelten. */
const SYSTEM_COLUMN_NAMES = new Set([
    'id',
    'contenttype',
    'contenttypeid',
    'modified',
    'created',
    'author',
    'editor',
    '_uiversionstring',
    'attachments',
    'edit',
    'linktitlenomenu',
    'linktitle',
    'docicon',
    'itemchildcount',
    'folderchildcount',
    'complianceassetid',
    'appauthor',
    'appeditor',
    '_colorhex',
    '_colortag',
    'colorhex',
    'colortag'
]);

/**
 * @param {string} method
 * @param {string} url
 * @param {string} token
 * @param {unknown} [body]
 */
async function graphJson(method, url, token, body) {
    const headers = {
        Authorization: 'Bearer ' + token,
        Accept: 'application/json'
    };
    if (body !== undefined) {
        headers['Content-Type'] = 'application/json';
    }
    const res = await fetch(url, {
        method,
        headers,
        body: body !== undefined ? JSON.stringify(body) : undefined
    });
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = { raw: text };
        }
    }
    if (!res.ok) {
        const msg =
            data && data.error && data.error.message
                ? data.error.message
                : text || 'HTTP ' + res.status;
        const err = new Error('Graph ' + method + ' ' + url + ': ' + msg);
        err.status = res.status;
        err.payload = data;
        throw err;
    }
    return data || {};
}

/**
 * @param {string} webUrl
 * @param {string} token
 */
async function resolveSiteId(webUrl, token) {
    const u = new URL(String(webUrl || '').trim());
    const host = u.hostname;
    let path = u.pathname.replace(/\/+$/, '');
    if (!path) path = '/';
    const siteUrl =
        GRAPH + '/sites/' + host + ':' + path + '?$select=id,displayName,webUrl';
    const site = await graphJson('GET', siteUrl, token);
    if (!site.id) throw new Error('Site-ID konnte nicht aufgelöst werden.');
    return site;
}

/**
 * @param {string} siteId
 * @param {string} listDisplayName
 * @param {string} token
 */
async function resolveListId(siteId, listDisplayName, token) {
    const want = String(listDisplayName || '').trim().toLowerCase();
    let url =
        GRAPH +
        '/sites/' +
        encodeURIComponent(siteId) +
        '/lists?$select=id,displayName&$top=200';
    while (url) {
        const page = await graphJson('GET', url, token);
        const rows = page.value || [];
        for (let i = 0; i < rows.length; i++) {
            if (String(rows[i].displayName || '').trim().toLowerCase() === want) {
                return rows[i];
            }
        }
        url = page['@odata.nextLink'] || '';
    }
    throw new Error('Liste „' + listDisplayName + '“ nicht gefunden.');
}

/**
 * @param {string} siteId
 * @param {string} listId
 * @param {string} token
 */
async function loadAllListItems(siteId, listId, token) {
    const items = [];
    let url =
        GRAPH +
        '/sites/' +
        encodeURIComponent(siteId) +
        '/lists/' +
        encodeURIComponent(listId) +
        '/items?$expand=fields&$top=100';
    while (url) {
        const page = await graphJson('GET', url, token);
        const rows = page.value || [];
        for (let i = 0; i < rows.length; i++) items.push(rows[i]);
        url = page['@odata.nextLink'] || '';
    }
    return items;
}

/** @type {{ siteId: string, listId: string, at: number } | null} */
let idCache = null;
const ID_CACHE_MS = 10 * 60 * 1000;

/** @type {{ siteId: string, listId: string, at: number, columns?: object[] } | null} */
let columnCache = null;
const COLUMN_CACHE_MS = 60 * 1000;

async function resolveSiteAndList(token) {
    const cfg = getConfig();
    const now = Date.now();
    if (idCache && now - idCache.at < ID_CACHE_MS) {
        return idCache;
    }
    const site = await resolveSiteId(cfg.siteWebUrl, token);
    const list = await resolveListId(site.id, cfg.listDisplayName, token);
    idCache = { siteId: site.id, listId: list.id, at: now };
    return idCache;
}

function formatValidUntil(raw) {
    if (raw == null || String(raw).trim() === '') return null;
    const d = new Date(raw);
    if (Number.isNaN(d.getTime())) return null;
    return d.toISOString().slice(0, 10);
}

/**
 * @param {Record<string, unknown>} col
 */
function classifyColumnType(col) {
    if (!col || typeof col !== 'object') return 'text';
    if (col.text) {
        return col.text.allowMultipleLines ? 'multiline' : 'text';
    }
    if (col.number) return 'number';
    if (col.boolean) return 'boolean';
    if (col.dateTime) return 'date';
    if (col.choice) return 'choice';
    return 'text';
}

/**
 * @param {Record<string, unknown>} col
 */
function mapColumnMeta(col) {
    const name = String(col.name || '').trim();
    const displayName = String(col.displayName || name).trim() || name;
    const type = classifyColumnType(col);
    /** @type {string[]|null} */
    let choices = null;
    if (type === 'choice' && col.choice && Array.isArray(col.choice.choices)) {
        choices = col.choice.choices.map((c) => String(c));
    }
    return {
        name,
        displayName,
        type,
        choices,
        readOnly: !!col.readOnly,
        required: !!col.required
    };
}

/**
 * @param {string} name
 */
function isExtraColumnName(name) {
    const key = String(name || '')
        .trim()
        .toLowerCase();
    if (!key) return false;
    if (key.startsWith('_')) return false;
    if (SYSTEM_COLUMN_NAMES.has(key)) return false;
    if (CORE_FIELD_NAMES.has(key)) return false;
    // ToolSiteUrl0 / manuell umbenannt – kein Extra-Feld in der Admin-UI
    if (looksLikeToolSiteUrlFieldName(name)) return false;
    return true;
}

/**
 * @param {unknown} raw
 * @param {string} type
 */
function normalizeExtraValue(raw, type) {
    if (raw === undefined) return undefined;
    if (raw === null || raw === '') {
        if (type === 'boolean') return false;
        if (type === 'number') return null;
        if (type === 'date') return null;
        return '';
    }
    if (type === 'boolean') {
        if (typeof raw === 'boolean') return raw;
        const s = String(raw).trim().toLowerCase();
        return s === '1' || s === 'true' || s === 'yes' || s === 'ja';
    }
    if (type === 'number') {
        const n = typeof raw === 'number' ? raw : Number(String(raw).replace(',', '.'));
        return Number.isFinite(n) ? n : null;
    }
    if (type === 'date') {
        return formatValidUntil(raw);
    }
    return String(raw);
}

/**
 * @param {Record<string, unknown>} item
 * @param {Array<{ name: string, type: string }>|null} [extraCols]
 */
function mapListItem(item, extraCols) {
    const fields = (item && item.fields) || {};
    const primaryDomain = String(fields.PrimaryDomain || '').trim() || null;
    const additionalDomains = String(fields.AdditionalDomains || '').trim() || null;
    /** @type {Record<string, unknown>} */
    const extra = {};
    const cols = Array.isArray(extraCols) ? extraCols : null;
    if (cols && cols.length) {
        for (let i = 0; i < cols.length; i++) {
            const c = cols[i];
            const n = c && c.name ? String(c.name) : '';
            if (!n || !(n in fields)) {
                if (n) extra[n] = null;
                continue;
            }
            const v = fields[n];
            if (c.type === 'date') extra[n] = formatValidUntil(v);
            else if (c.type === 'boolean') extra[n] = !!v;
            else if (c.type === 'number') {
                const num = typeof v === 'number' ? v : Number(v);
                extra[n] = Number.isFinite(num) ? num : null;
            } else {
                extra[n] = v == null || v === '' ? null : String(v);
            }
        }
    } else {
        Object.keys(fields).forEach((key) => {
            if (!isExtraColumnName(key)) return;
            const v = fields[key];
            if (v == null || v === '') extra[key] = null;
            else if (typeof v === 'boolean' || typeof v === 'number') extra[key] = v;
            else extra[key] = String(v);
        });
    }
    const entitlementDoc = parseEntitlementsRaw(fields[ENTITLEMENTS_FIELD]);
    const entitlements = resolveFeatureEntitlements(entitlementDoc, { legacyFullAccess: true });

    return {
        id: String(item.id || ''),
        schoolName: String(fields.Title || '').trim() || null,
        tenantId: String(fields.TenantId || '').trim() || null,
        primaryDomain,
        additionalDomains,
        domains: parseDomainList(fields.PrimaryDomain, fields.AdditionalDomains),
        status: String(fields.Status || '').trim().toLowerCase() || null,
        validUntil: formatValidUntil(fields.ValidUntil),
        contactEmail: String(fields.ContactEmail || '').trim() || null,
        contactEmails: parseContactEmails(fields.ContactEmail, fields.ItContacts),
        notes: String(fields.Notes || '').trim() || null,
        /** Nur Zeiger auf Schul-SharePoint – keine Klassen/Lehrer-Inhalte. */
        toolSiteUrl: pickToolSiteUrlFromFields(fields) || null,
        entitlementsDoc: entitlementDoc,
        entitlements,
        extra
    };
}

/**
 * @param {Record<string, unknown>} body
 * @param {{ partial?: boolean, extraColumns?: Array<{ name: string, type: string }> }} [opts]
 */
function fieldsFromBody(body, opts) {
    const partial = !!(opts && opts.partial);
    const src = body && typeof body === 'object' ? body : {};
    /** @type {Record<string, unknown>} */
    const fields = {};

    function set(name, value) {
        if (value === undefined) return;
        fields[name] = value;
    }

    if (!partial || src.schoolName !== undefined || src.Title !== undefined) {
        set('Title', String(src.schoolName != null ? src.schoolName : src.Title || '').trim());
    }
    if (!partial || src.tenantId !== undefined || src.TenantId !== undefined) {
        const schoolHint =
            src.schoolName != null
                ? src.schoolName
                : src.Title != null
                  ? src.Title
                  : undefined;
        const tid = assertTenantIdAllowed(
            src.tenantId != null ? src.tenantId : src.TenantId,
            schoolHint
        );
        set('TenantId', isPendingTenantId(tid) ? tid : tid.toLowerCase());
    }
    if (!partial || src.primaryDomain !== undefined || src.PrimaryDomain !== undefined) {
        set(
            'PrimaryDomain',
            String(src.primaryDomain != null ? src.primaryDomain : src.PrimaryDomain || '').trim()
        );
    }
    if (
        !partial ||
        src.additionalDomains !== undefined ||
        src.AdditionalDomains !== undefined
    ) {
        let add =
            src.additionalDomains != null ? src.additionalDomains : src.AdditionalDomains;
        if (Array.isArray(add)) add = add.join('\n');
        const addStr = String(add == null ? '' : add).trim();
        if (addStr) set('AdditionalDomains', addStr);
    }
    if (!partial || src.status !== undefined || src.Status !== undefined) {
        set(
            'Status',
            String(src.status != null ? src.status : src.Status || '')
                .trim()
                .toLowerCase()
        );
    }
    if (!partial || src.validUntil !== undefined || src.ValidUntil !== undefined) {
        const vu = src.validUntil != null ? src.validUntil : src.ValidUntil;
        if (vu === null || vu === '') set('ValidUntil', null);
        else if (vu !== undefined) set('ValidUntil', formatValidUntil(vu));
    }
    if (!partial || src.contactEmail !== undefined || src.ContactEmail !== undefined) {
        set(
            'ContactEmail',
            String(src.contactEmail != null ? src.contactEmail : src.ContactEmail || '').trim()
        );
    }
    if (!partial || src.itContacts !== undefined || src.ItContacts !== undefined) {
        let itc = src.itContacts != null ? src.itContacts : src.ItContacts;
        if (Array.isArray(itc)) itc = itc.join('\n');
        set('ItContacts', String(itc == null ? '' : itc).trim());
    }
    if (!partial || src.notes !== undefined || src.Notes !== undefined) {
        set('Notes', String(src.notes != null ? src.notes : src.Notes || '').trim());
    }
    if (
        !partial ||
        src.toolSiteUrl !== undefined ||
        src.ToolSiteUrl !== undefined
    ) {
        const raw =
            src.toolSiteUrl != null ? src.toolSiteUrl : src.ToolSiteUrl;
        const url = String(raw == null ? '' : raw).trim();
        const colName =
            opts && opts.toolSiteUrlFieldName
                ? String(opts.toolSiteUrlFieldName).trim()
                : TOOL_SITE_URL_FIELD;
        const kind =
            opts && opts.toolSiteUrlFieldKind === 'hyperlink' ? 'hyperlink' : 'text';
        set(colName || TOOL_SITE_URL_FIELD, toolSiteUrlFieldValue(url, kind));
    }

    if (
        !partial ||
        src.entitlements !== undefined ||
        src.entitlementsDoc !== undefined ||
        src.LicenseEntitlements !== undefined
    ) {
        let ent =
            src.entitlementsDoc != null
                ? src.entitlementsDoc
                : src.entitlements != null
                  ? src.entitlements
                  : src.LicenseEntitlements;
        if (typeof ent === 'string') {
            const parsed = parseEntitlementsRaw(ent);
            ent = parsed || { packages: [], toolsInclude: [], toolsExclude: [] };
        } else if (ent && typeof ent === 'object') {
            /* packages direkt oder resolved-Objekt mit packages */
            ent = {
                packages: Array.isArray(ent.packages) ? ent.packages : [],
                toolsInclude: Array.isArray(ent.toolsInclude) ? ent.toolsInclude : [],
                toolsExclude: Array.isArray(ent.toolsExclude) ? ent.toolsExclude : [],
                notes: ent.notes
            };
        } else {
            ent = { packages: [], toolsInclude: [], toolsExclude: [] };
        }
        set(ENTITLEMENTS_FIELD, serializeEntitlements(ent));
    }

    const extraSrc =
        src.extra && typeof src.extra === 'object' && !Array.isArray(src.extra) ? src.extra : null;
    const knownExtra = Array.isArray(opts && opts.extraColumns) ? opts.extraColumns : [];
    const typeByName = {};
    for (let i = 0; i < knownExtra.length; i++) {
        const c = knownExtra[i];
        if (c && c.name) typeByName[String(c.name)] = c.type || 'text';
    }

    if (extraSrc) {
        Object.keys(extraSrc).forEach((name) => {
            if (!isExtraColumnName(name)) return;
            const type = typeByName[name] || 'text';
            const normalized = normalizeExtraValue(extraSrc[name], type);
            if (normalized === undefined) return;
            set(name, normalized);
        });
    }

    return fields;
}

/**
 * @param {string} displayName
 * @param {string} [explicit]
 */
function sanitizeInternalName(displayName, explicit) {
    let raw = String(explicit || displayName || '').trim();
    raw = raw.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    raw = raw.replace(/[^a-zA-Z0-9_]/g, '');
    if (!raw) raw = 'ExtraField';
    if (!/^[A-Za-z]/.test(raw)) raw = 'X' + raw;
    return raw.slice(0, 32);
}

/**
 * @param {Record<string, unknown>} body
 */
function buildColumnDefinition(body) {
    const src = body && typeof body === 'object' ? body : {};
    const displayName = String(src.displayName || src.name || '').trim();
    if (!displayName) {
        const err = new Error('Anzeigename (displayName) ist erforderlich.');
        err.status = 400;
        throw err;
    }
    const name = sanitizeInternalName(displayName, src.name);
    if (!isExtraColumnName(name)) {
        const err = new Error(
            '„' + name + '“ ist eine System-/Kernspalte und kann nicht als Extra angelegt werden.'
        );
        err.status = 400;
        throw err;
    }
    const type = String(src.type || 'text')
        .trim()
        .toLowerCase();
    /** @type {Record<string, unknown>} */
    const def = { name, displayName };
    if (type === 'multiline') {
        def.text = { allowMultipleLines: true };
    } else if (type === 'number') {
        def.number = {};
    } else if (type === 'boolean' || type === 'yesno') {
        def.boolean = {};
    } else if (type === 'date') {
        def.dateTime = { format: 'dateOnly' };
    } else if (type === 'choice') {
        let choices = src.choices;
        if (typeof choices === 'string') {
            choices = choices
                .split(/[\n,;]+/)
                .map((s) => String(s).trim())
                .filter(Boolean);
        }
        if (!Array.isArray(choices) || !choices.length) {
            const err = new Error('Choice-Spalte braucht mindestens eine Option (choices).');
            err.status = 400;
            throw err;
        }
        def.choice = {
            allowTextEntry: false,
            choices: choices.map((c) => String(c).trim()).filter(Boolean)
        };
    } else {
        def.text = { allowMultipleLines: false, maxLength: 255 };
    }
    return def;
}

function assertCreateFields(fields) {
    if (!fields.Title) {
        const err = new Error('Schulname (schoolName) ist erforderlich.');
        err.status = 400;
        throw err;
    }
    fields.TenantId = assertTenantIdAllowed(fields.TenantId, fields.Title);
    if (!fields.Status) fields.Status = 'trial';
}

/**
 * @param {string} siteId
 * @param {string} listId
 * @param {string} token
 */
async function loadListColumnsRaw(siteId, listId, token) {
    const cols = [];
    let url =
        GRAPH +
        '/sites/' +
        encodeURIComponent(siteId) +
        '/lists/' +
        encodeURIComponent(listId) +
        '/columns?$top=200';
    while (url) {
        const page = await graphJson('GET', url, token);
        const rows = page.value || [];
        for (let i = 0; i < rows.length; i++) cols.push(rows[i]);
        url = page['@odata.nextLink'] || '';
    }
    return cols;
}

/**
 * @param {{ force?: boolean }} [opts]
 */
async function listExtraColumns(opts) {
    const force = !!(opts && opts.force);
    const now = Date.now();
    if (!force && columnCache && now - columnCache.at < COLUMN_CACHE_MS && columnCache.columns) {
        return columnCache.columns.slice();
    }
    const token = await getOperatorGraphToken();
    const ids = await resolveSiteAndList(token);
    const raw = await loadListColumnsRaw(ids.siteId, ids.listId, token);
    const columns = raw
        .filter((c) => c && isExtraColumnName(c.name) && !c.readOnly && !c.hidden)
        .map(mapColumnMeta)
        .filter((c) => c.name)
        .sort((a, b) =>
            String(a.displayName).localeCompare(String(b.displayName), 'de', { sensitivity: 'base' })
        );
    columnCache = {
        siteId: ids.siteId,
        listId: ids.listId,
        at: now,
        columns
    };
    return columns.slice();
}

/**
 * @param {Record<string, unknown>} body
 */
async function createExtraColumn(body) {
    const def = buildColumnDefinition(body);
    const token = await getOperatorGraphToken();
    const ids = await resolveSiteAndList(token);
    const existing = await loadListColumnsRaw(ids.siteId, ids.listId, token);
    const clash = existing.find(
        (c) => String(c.name || '').toLowerCase() === String(def.name).toLowerCase()
    );
    if (clash) {
        const err = new Error('Spalte „' + def.name + '“ existiert bereits.');
        err.status = 409;
        throw err;
    }
    const created = await graphJson(
        'POST',
        GRAPH +
            '/sites/' +
            encodeURIComponent(ids.siteId) +
            '/lists/' +
            encodeURIComponent(ids.listId) +
            '/columns',
        token,
        def
    );
    columnCache = null;
    return mapColumnMeta(created);
}

/**
 * @param {string} columnName
 */
async function deleteExtraColumn(columnName) {
    const name = String(columnName || '').trim();
    if (!name) {
        const err = new Error('Spaltenname fehlt.');
        err.status = 400;
        throw err;
    }
    if (!isExtraColumnName(name)) {
        const err = new Error('Kern- oder Systemspalten können nicht gelöscht werden.');
        err.status = 400;
        throw err;
    }
    const token = await getOperatorGraphToken();
    const ids = await resolveSiteAndList(token);
    await graphJson(
        'DELETE',
        GRAPH +
            '/sites/' +
            encodeURIComponent(ids.siteId) +
            '/lists/' +
            encodeURIComponent(ids.listId) +
            '/columns/' +
            encodeURIComponent(name),
        token
    );
    columnCache = null;
    return { deleted: true, name };
}

/**
 * @param {string} tenantId
 */
async function lookupLicenseFields(tenantId) {
    const token = await getOperatorGraphToken();
    const ids = await resolveSiteAndList(token);
    const items = await loadAllListItems(ids.siteId, ids.listId, token);
    return findFieldsForTenant(items, tenantId);
}

/**
 * Text-Spalte ToolSiteUrl anlegen, falls fehlend (Schul-SharePoint-Zeiger).
 * Akzeptiert auch manuell angelegte Spalten mit ähnlichem Anzeigenamen.
 */
async function ensureToolSiteUrlColumn(siteId, listId, token) {
    const raw = await loadListColumnsRaw(siteId, listId, token);
    const existing = findToolSiteUrlColumnMeta(raw);
    if (existing) {
        return {
            created: false,
            name: existing.name,
            kind: existing.kind,
            reused: existing.name !== TOOL_SITE_URL_FIELD
        };
    }
    await graphJson(
        'POST',
        GRAPH +
            '/sites/' +
            encodeURIComponent(siteId) +
            '/lists/' +
            encodeURIComponent(listId) +
            '/columns',
        token,
        {
            name: TOOL_SITE_URL_FIELD,
            displayName: 'Tool-Site-URL (Schul-SharePoint)',
            text: { allowMultipleLines: false, maxLength: 500 }
        }
    );
    columnCache = null;
    return { created: true, name: TOOL_SITE_URL_FIELD, kind: 'text', reused: false };
}

/**
 * Multiline-Spalte LicenseEntitlements anlegen, falls fehlend.
 */
async function ensureEntitlementsColumn(siteId, listId, token) {
    const raw = await loadListColumnsRaw(siteId, listId, token);
    const exists = raw.some(
        (c) =>
            String((c && c.name) || '')
                .trim()
                .toLowerCase() === ENTITLEMENTS_FIELD.toLowerCase()
    );
    if (exists) return { created: false, name: ENTITLEMENTS_FIELD };
    await graphJson(
        'POST',
        GRAPH +
            '/sites/' +
            encodeURIComponent(siteId) +
            '/lists/' +
            encodeURIComponent(listId) +
            '/columns',
        token,
        {
            name: ENTITLEMENTS_FIELD,
            displayName: 'Lizenz-Module (JSON)',
            text: { allowMultipleLines: true }
        }
    );
    columnCache = null;
    return { created: true, name: ENTITLEMENTS_FIELD };
}

/**
 * Stammdaten-Spalten idempotent anlegen (Adresse, Rollen, Telefon, …).
 * @param {string} siteId
 * @param {string} listId
 * @param {string} token
 */
async function ensureStammdatenColumns(siteId, listId, token) {
    const raw = await loadListColumnsRaw(siteId, listId, token);
    /** @type {Record<string, true>} */
    const have = {};
    for (let i = 0; i < raw.length; i++) {
        const n = String((raw[i] && raw[i].name) || '')
            .trim()
            .toLowerCase();
        if (n) have[n] = true;
    }

    const created = [];
    const existing = [];
    for (let i = 0; i < SCHOOL_STAMMDATEN_FIELDS.length; i++) {
        const field = SCHOOL_STAMMDATEN_FIELDS[i];
        const key = String(field.name).toLowerCase();
        if (have[key]) {
            existing.push(field.name);
            continue;
        }
        const def = buildColumnDefinition({
            name: field.name,
            displayName: field.displayName,
            type: field.type
        });
        await graphJson(
            'POST',
            GRAPH +
                '/sites/' +
                encodeURIComponent(siteId) +
                '/lists/' +
                encodeURIComponent(listId) +
                '/columns',
            token,
            def
        );
        have[key] = true;
        created.push(field.name);
    }
    if (created.length) columnCache = null;
    return { created, existing, total: SCHOOL_STAMMDATEN_FIELDS.length };
}

/**
 * Kernspalte ItContacts (Mehrzeilig) anlegen, falls fehlend.
 */
async function ensureItContactsColumn(siteId, listId, token) {
    const raw = await loadListColumnsRaw(siteId, listId, token);
    const exists = raw.some(
        (c) =>
            String((c && c.name) || '')
                .trim()
                .toLowerCase() === 'itcontacts'
    );
    if (exists) return { created: false, name: 'ItContacts' };
    await graphJson(
        'POST',
        GRAPH +
            '/sites/' +
            encodeURIComponent(siteId) +
            '/lists/' +
            encodeURIComponent(listId) +
            '/columns',
        token,
        {
            name: 'ItContacts',
            displayName: 'IT-Kontakte',
            text: { allowMultipleLines: true }
        }
    );
    columnCache = null;
    return { created: true, name: 'ItContacts' };
}

/**
 * Entitlements + Stammdaten-Spalten sicherstellen.
 */
async function ensureLicenseListSchema(siteId, listId, token) {
    const entitlementsColumn = await ensureEntitlementsColumn(siteId, listId, token);
    const itContactsColumn = await ensureItContactsColumn(siteId, listId, token);
    const toolSiteUrlColumn = await ensureToolSiteUrlColumn(siteId, listId, token);
    const stammdaten = await ensureStammdatenColumns(siteId, listId, token);
    return { entitlementsColumn, itContactsColumn, toolSiteUrlColumn, stammdaten };
}

async function listLicenses() {
    const token = await getOperatorGraphToken();
    const ids = await resolveSiteAndList(token);
    await ensureLicenseListSchema(ids.siteId, ids.listId, token);
    const [items, columns] = await Promise.all([
        loadAllListItems(ids.siteId, ids.listId, token),
        listExtraColumns({ force: true })
    ]);
    const schools = items.map((it) => mapListItem(it, columns)).filter((x) => x.id);
    return { schools, columns };
}

/**
 * @param {Record<string, unknown>} body
 */
async function createLicense(body) {
    const token = await getOperatorGraphToken();
    const ids = await resolveSiteAndList(token);
    const schema = await ensureLicenseListSchema(ids.siteId, ids.listId, token);
    const columns = await listExtraColumns({ force: true });
    const src = body && typeof body === 'object' ? Object.assign({}, body) : {};
    if (src.entitlements === undefined && src.entitlementsDoc === undefined) {
        src.entitlementsDoc = defaultNewSchoolEntitlements();
    }
    if (src.tenantId !== undefined || src.TenantId !== undefined) {
        src.tenantId = assertTenantIdAllowed(
            src.tenantId != null ? src.tenantId : src.TenantId,
            src.schoolName || src.Title
        );
    } else {
        src.tenantId = makePendingTenantId(src.schoolName || src.Title);
    }
    const toolCol = (schema && schema.toolSiteUrlColumn) || {};
    const fields = fieldsFromBody(src, {
        partial: false,
        extraColumns: columns,
        toolSiteUrlFieldName: toolCol.name || TOOL_SITE_URL_FIELD,
        toolSiteUrlFieldKind: toolCol.kind === 'hyperlink' ? 'hyperlink' : 'text'
    });
    assertCreateFields(fields);
    const created = await graphJson(
        'POST',
        GRAPH +
            '/sites/' +
            encodeURIComponent(ids.siteId) +
            '/lists/' +
            encodeURIComponent(ids.listId) +
            '/items',
        token,
        { fields }
    );
    const full = await graphJson(
        'GET',
        GRAPH +
            '/sites/' +
            encodeURIComponent(ids.siteId) +
            '/lists/' +
            encodeURIComponent(ids.listId) +
            '/items/' +
            encodeURIComponent(created.id) +
            '?$expand=fields',
        token
    );
    return mapListItem(full, columns);
}

/**
 * @param {string} itemId
 * @param {Record<string, unknown>} body
 */
async function updateLicense(itemId, body) {
    const id = String(itemId || '').trim();
    if (!id) {
        const err = new Error('Listen-ID fehlt.');
        err.status = 400;
        throw err;
    }
    const token = await getOperatorGraphToken();
    const ids = await resolveSiteAndList(token);
    // Immer Schema sichern (u. a. ToolSiteUrl) – sonst schlägt PATCH fehl / Feld bleibt leer
    const schema = await ensureLicenseListSchema(ids.siteId, ids.listId, token);
    const columns = await listExtraColumns({ force: true });
    const toolCol = (schema && schema.toolSiteUrlColumn) || {};
    const fields = fieldsFromBody(body, {
        partial: true,
        extraColumns: columns,
        toolSiteUrlFieldName: toolCol.name || TOOL_SITE_URL_FIELD,
        toolSiteUrlFieldKind: toolCol.kind === 'hyperlink' ? 'hyperlink' : 'text'
    });
    if (!Object.keys(fields).length) {
        const err = new Error('Keine Felder zum Aktualisieren.');
        err.status = 400;
        throw err;
    }
    try {
        await graphJson(
            'PATCH',
            GRAPH +
                '/sites/' +
                encodeURIComponent(ids.siteId) +
                '/lists/' +
                encodeURIComponent(ids.listId) +
                '/items/' +
                encodeURIComponent(id) +
                '/fields',
            token,
            fields
        );
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e);
        if (
            body &&
            (body.toolSiteUrl !== undefined || body.ToolSiteUrl !== undefined) &&
            /toolsiteurl|field|invalid|not recognized|nicht erkannt/i.test(msg)
        ) {
            const err = new Error(
                'Tool-Site-URL konnte nicht geschrieben werden. In der Liste „' +
                    (getConfig().listDisplayName || 'Übersicht-Tenants') +
                    '“ muss eine Text-/URL-Spalte existieren (interner Name idealerweise ToolSiteUrl). Details: ' +
                    msg
            );
            err.status = e.status || 400;
            throw err;
        }
        throw e;
    }
    const full = await graphJson(
        'GET',
        GRAPH +
            '/sites/' +
            encodeURIComponent(ids.siteId) +
            '/lists/' +
            encodeURIComponent(ids.listId) +
            '/items/' +
            encodeURIComponent(id) +
            '?$expand=fields',
        token
    );
    return mapListItem(full, columns);
}

/**
 * @param {string} itemId
 */
async function deleteLicense(itemId) {
    const id = String(itemId || '').trim();
    if (!id) {
        const err = new Error('Listen-ID fehlt.');
        err.status = 400;
        throw err;
    }
    const token = await getOperatorGraphToken();
    const ids = await resolveSiteAndList(token);
    await graphJson(
        'DELETE',
        GRAPH +
            '/sites/' +
            encodeURIComponent(ids.siteId) +
            '/lists/' +
            encodeURIComponent(ids.listId) +
            '/items/' +
            encodeURIComponent(id),
        token
    );
    return { deleted: true, id };
}

function clearSharePointCache() {
    idCache = null;
    columnCache = null;
}

/**
 * Schulen ohne LicenseEntitlements-JSON explizit befüllen.
 * mode:
 * - legacy-full: alle Pakete (bisheriges Verhalten materialisieren)
 * - starter: Core + Personen + Schuljahr
 * @param {{ mode?: string, dryRun?: boolean }} [opts]
 */
async function migrateSchoolEntitlements(opts) {
    const mode = String((opts && opts.mode) || 'legacy-full')
        .trim()
        .toLowerCase();
    const dryRun = !!(opts && opts.dryRun);
    const token = await getOperatorGraphToken();
    const ids = await resolveSiteAndList(token);
    await ensureEntitlementsColumn(ids.siteId, ids.listId, token);
    const items = await loadAllListItems(ids.siteId, ids.listId, token);
    const columns = await listExtraColumns();

    const doc =
        mode === 'starter'
            ? defaultNewSchoolEntitlements()
            : {
                  v: 1,
                  packages: allPackageIds(),
                  toolsInclude: [],
                  toolsExclude: [],
                  notes: 'migration:legacy-full'
              };

    /** @type {Array<{ id: string, schoolName: string|null }>} */
    const candidates = [];
    for (let i = 0; i < items.length; i++) {
        const mapped = mapListItem(items[i], columns);
        if (!mapped.id) continue;
        if (mapped.entitlementsDoc) continue;
        candidates.push({ id: mapped.id, schoolName: mapped.schoolName });
    }

    if (dryRun) {
        return {
            mode,
            dryRun: true,
            updated: 0,
            candidates: candidates.length,
            schools: candidates
        };
    }

    let updated = 0;
    for (let i = 0; i < candidates.length; i++) {
        await updateLicense(candidates[i].id, { entitlementsDoc: doc });
        updated += 1;
    }
    return {
        mode,
        dryRun: false,
        updated,
        candidates: candidates.length,
        schools: candidates
    };
}

module.exports = {
    graphJson,
    lookupLicenseFields,
    listLicenses,
    createLicense,
    updateLicense,
    deleteLicense,
    listExtraColumns,
    createExtraColumn,
    deleteExtraColumn,
    ensureEntitlementsColumn,
    ensureItContactsColumn,
    ensureToolSiteUrlColumn,
    ensureStammdatenColumns,
    ensureLicenseListSchema,
    TOOL_SITE_URL_FIELD,
    migrateSchoolEntitlements,
    resolveSiteId,
    resolveListId,
    loadAllListItems,
    clearSharePointCache,
    mapListItem,
    fieldsFromBody,
    pickToolSiteUrlFromFields,
    findToolSiteUrlColumnMeta,
    normalizeToolSiteUrlValue,
    toolSiteUrlFieldValue,
    isExtraColumnName,
    sanitizeInternalName,
    buildColumnDefinition,
    isPendingTenantId,
    makePendingTenantId,
    normalizeTenantIdInput,
    assertTenantIdAllowed
};
