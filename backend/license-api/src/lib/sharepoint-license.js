'use strict';

const { getConfig } = require('./config');
const { getOperatorGraphToken } = require('./msal-app-only');
const { findFieldsForTenant, parseDomainList } = require('./evaluate-license');

const GRAPH = 'https://graph.microsoft.com/v1.0';

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
 * @param {Record<string, unknown>} item
 */
function mapListItem(item) {
    const fields = (item && item.fields) || {};
    const primaryDomain = String(fields.PrimaryDomain || '').trim() || null;
    const additionalDomains = String(fields.AdditionalDomains || '').trim() || null;
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
        notes: String(fields.Notes || '').trim() || null
    };
}

/**
 * @param {Record<string, unknown>} body
 * @param {{ partial?: boolean }} [opts]
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
        set('AdditionalDomains', String(add == null ? '' : add).trim());
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

    return fields;
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
    const items = await loadAllListItems(ids.siteId, ids.listId, token);
    return items.map(mapListItem).filter((x) => x.id);
}

/**
 * @param {Record<string, unknown>} body
 */
async function createLicense(body) {
    const token = await getOperatorGraphToken();
    const ids = await resolveSiteAndList(token);
    const fields = fieldsFromBody(body, { partial: false });
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
    return mapListItem(full);
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
    const fields = fieldsFromBody(body, { partial: true });
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
    return mapListItem(full);
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
}

module.exports = {
    lookupLicenseFields,
    listLicenses,
    createLicense,
    updateLicense,
    deleteLicense,
    resolveSiteId,
    resolveListId,
    loadAllListItems,
    clearSharePointCache,
    mapListItem,
    fieldsFromBody
};
