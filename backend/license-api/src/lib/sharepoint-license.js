'use strict';

const { getConfig } = require('./config');
const { getOperatorGraphToken } = require('./msal-app-only');
const { findFieldsForTenant, parseDomainList } = require('./evaluate-license');

const GRAPH = 'https://graph.microsoft.com/v1.0';

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
    'notes'
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
        notes: String(fields.Notes || '').trim() || null,
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
        set(
            'TenantId',
            String(src.tenantId != null ? src.tenantId : src.TenantId || '')
                .trim()
                .toLowerCase()
        );
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
    if (!partial || src.notes !== undefined || src.Notes !== undefined) {
        set('Notes', String(src.notes != null ? src.notes : src.Notes || '').trim());
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
    if (!fields.TenantId) {
        const err = new Error('Tenant-ID ist erforderlich.');
        err.status = 400;
        throw err;
    }
    if (!fields.Status) fields.Status = 'active';
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

async function listLicenses() {
    const token = await getOperatorGraphToken();
    const ids = await resolveSiteAndList(token);
    const [items, columns] = await Promise.all([
        loadAllListItems(ids.siteId, ids.listId, token),
        listExtraColumns()
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
    const columns = await listExtraColumns();
    const fields = fieldsFromBody(body, { partial: false, extraColumns: columns });
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
    const columns = await listExtraColumns();
    const fields = fieldsFromBody(body, { partial: true, extraColumns: columns });
    if (!Object.keys(fields).length) {
        const err = new Error('Keine Felder zum Aktualisieren.');
        err.status = 400;
        throw err;
    }
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
    resolveSiteId,
    resolveListId,
    loadAllListItems,
    clearSharePointCache,
    mapListItem,
    fieldsFromBody,
    isExtraColumnName,
    sanitizeInternalName,
    buildColumnDefinition
};
