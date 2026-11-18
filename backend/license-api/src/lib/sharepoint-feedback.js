'use strict';

const { getConfig } = require('./config');
const { getOperatorGraphToken } = require('./msal-app-only');
const { graphJson, resolveSiteId, resolveListId } = require('./sharepoint-license');

const GRAPH = 'https://graph.microsoft.com/v1.0';

const FEEDBACK_LIST_DEFAULT = 'MS365schule-Feedback';
const FEEDBACK_FOLDER = 'Feedback';
const MAX_ATTACHMENTS = 3;
const MAX_ATTACHMENT_BYTES = 1.5 * 1024 * 1024;

const CATEGORIES = ['Feedback', 'Bug', 'Idee', 'Frage'];
const STATUSES = ['Neu', 'In Bearbeitung', 'Erledigt', 'Abgelehnt'];

/** @type {{ siteId: string, listId: string, at: number } | null} */
let feedbackListCache = null;
const LIST_CACHE_MS = 10 * 60 * 1000;

/**
 * @returns {string}
 */
function feedbackListDisplayName() {
    const cfg = getConfig();
    return String(cfg.feedbackListDisplayName || FEEDBACK_LIST_DEFAULT).trim() || FEEDBACK_LIST_DEFAULT;
}

/**
 * @param {string} rel
 */
function encodeDrivePath(rel) {
    return String(rel || '')
        .replace(/\\/g, '/')
        .split('/')
        .filter(Boolean)
        .map((seg) => encodeURIComponent(seg))
        .join('/');
}

/**
 * @param {string} method
 * @param {string} url
 * @param {string} token
 * @param {Buffer|string} [body]
 * @param {string} [contentType]
 */
async function graphBinary(method, url, token, body, contentType) {
    const headers = {
        Authorization: 'Bearer ' + token,
        Accept: 'application/json'
    };
    if (body !== undefined) {
        headers['Content-Type'] = contentType || 'application/octet-stream';
    }
    const res = await fetch(url, {
        method,
        headers,
        body
    });
    const text = Buffer.from(await res.arrayBuffer()).toString('utf8');
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
 * @returns {Array<{ name: string, displayName: string, def: Record<string, unknown> }>}
 */
function feedbackColumnDefs() {
    return [
        {
            name: 'Beschreibung',
            displayName: 'Beschreibung',
            def: { text: { allowMultipleLines: true } }
        },
        {
            name: 'Kategorie',
            displayName: 'Kategorie',
            def: { choice: { allowTextEntry: false, choices: CATEGORIES } }
        },
        {
            name: 'Status',
            displayName: 'Status',
            def: { choice: { allowTextEntry: false, choices: STATUSES } }
        },
        {
            name: 'MelderEmail',
            displayName: 'Melder-E-Mail',
            def: { text: { allowMultipleLines: false, maxLength: 255 } }
        },
        {
            name: 'MelderName',
            displayName: 'Melder-Name',
            def: { text: { allowMultipleLines: false, maxLength: 255 } }
        },
        {
            name: 'TenantId',
            displayName: 'Tenant-ID',
            def: { text: { allowMultipleLines: false, maxLength: 64 } }
        },
        {
            name: 'Schulname',
            displayName: 'Schulname',
            def: { text: { allowMultipleLines: false, maxLength: 255 } }
        },
        {
            name: 'Domain',
            displayName: 'Domain',
            def: { text: { allowMultipleLines: false, maxLength: 255 } }
        },
        {
            name: 'SeitenUrl',
            displayName: 'Seiten-URL',
            def: { text: { allowMultipleLines: true } }
        },
        {
            name: 'UserAgent',
            displayName: 'User-Agent',
            def: { text: { allowMultipleLines: true } }
        },
        {
            name: 'AnhangUrls',
            displayName: 'Anhang-URLs (JSON)',
            def: { text: { allowMultipleLines: true } }
        }
    ];
}

/**
 * @param {string} siteId
 * @param {string} token
 * @param {string} listName
 */
async function ensureFeedbackList(siteId, token, listName) {
    try {
        return await resolveListId(siteId, listName, token);
    } catch (e) {
        if (!/nicht gefunden/i.test(String(e && e.message))) throw e;
    }
    const created = await graphJson(
        'POST',
        GRAPH + '/sites/' + encodeURIComponent(siteId) + '/lists',
        token,
        {
            displayName: listName,
            description: 'Feedback und Bug-Reports aus MS365-Schultools (nur Betreiber).',
            list: { template: 'genericList' }
        }
    );
    if (!created || !created.id) {
        throw new Error('Feedback-Liste konnte nicht angelegt werden.');
    }
    return created;
}

/**
 * @param {string} siteId
 * @param {string} listId
 * @param {string} token
 */
async function ensureFeedbackColumns(siteId, listId, token) {
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
    const defs = feedbackColumnDefs();
    for (let i = 0; i < defs.length; i++) {
        const col = defs[i];
        if (have[col.name.toLowerCase()]) {
            existing.push(col.name);
            continue;
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
            Object.assign({ name: col.name, displayName: col.displayName }, col.def)
        );
        created.push(col.name);
        have[col.name.toLowerCase()] = true;
    }
    return { created, existing };
}

/**
 * Liste + Spalten idempotent sicherstellen.
 */
async function ensureFeedbackSchema() {
    const token = await getOperatorGraphToken();
    const cfg = getConfig();
    const listName = feedbackListDisplayName();
    const site = await resolveSiteId(cfg.siteWebUrl, token);
    const list = await ensureFeedbackList(site.id, token, listName);
    const columns = await ensureFeedbackColumns(site.id, list.id, token);
    feedbackListCache = { siteId: site.id, listId: list.id, at: Date.now() };
    return {
        siteWebUrl: cfg.siteWebUrl,
        listDisplayName: listName,
        listId: list.id,
        siteId: site.id,
        columns
    };
}

async function resolveFeedbackList(token) {
    const now = Date.now();
    if (feedbackListCache && now - feedbackListCache.at < LIST_CACHE_MS) {
        return feedbackListCache;
    }
    const cfg = getConfig();
    const site = await resolveSiteId(cfg.siteWebUrl, token);
    let list;
    try {
        list = await resolveListId(site.id, feedbackListDisplayName(), token);
    } catch (e) {
        if (!/nicht gefunden/i.test(String(e && e.message))) throw e;
        const schema = await ensureFeedbackSchema();
        return { siteId: schema.siteId, listId: schema.listId, at: Date.now() };
    }
    feedbackListCache = { siteId: site.id, listId: list.id, at: now };
    return feedbackListCache;
}

/**
 * @param {unknown} raw
 * @returns {Array<{ name: string, url: string, path?: string }>}
 */
function parseAnhangUrls(raw) {
    if (raw == null || raw === '') return [];
    if (Array.isArray(raw)) {
        return raw
            .map((a) => {
                if (!a || typeof a !== 'object') return null;
                const name = String(a.name || '').trim();
                const url = String(a.url || a.webUrl || '').trim();
                if (!url) return null;
                return { name: name || 'Datei', url, path: a.path ? String(a.path) : undefined };
            })
            .filter(Boolean);
    }
    const s = String(raw).trim();
    if (!s) return [];
    try {
        return parseAnhangUrls(JSON.parse(s));
    } catch {
        return s
            .split(/\r?\n/)
            .map((line) => String(line || '').trim())
            .filter(Boolean)
            .map((url) => ({ name: 'Anhang', url }));
    }
}

/**
 * @param {Record<string, unknown>} item
 */
function mapFeedbackItem(item) {
    const fields = (item && item.fields) || {};
    const created =
        fields.Created ||
        item.createdDateTime ||
        fields.createdDateTime ||
        null;
    return {
        id: String(item.id || ''),
        title: String(fields.Title || '').trim(),
        description: String(fields.Beschreibung || '').trim(),
        category: String(fields.Kategorie || 'Feedback').trim() || 'Feedback',
        status: String(fields.Status || 'Neu').trim() || 'Neu',
        reporterEmail: String(fields.MelderEmail || '').trim(),
        reporterName: String(fields.MelderName || '').trim(),
        tenantId: String(fields.TenantId || '').trim(),
        schoolName: String(fields.Schulname || '').trim(),
        domain: String(fields.Domain || '').trim(),
        pageUrl: String(fields.SeitenUrl || '').trim(),
        userAgent: String(fields.UserAgent || '').trim(),
        attachments: parseAnhangUrls(fields.AnhangUrls),
        createdAt: created ? String(created) : null,
        modifiedAt: fields.Modified ? String(fields.Modified) : null
    };
}

/**
 * @param {string} raw
 */
function normalizeCategory(raw) {
    const s = String(raw || '').trim();
    const hit = CATEGORIES.find((c) => c.toLowerCase() === s.toLowerCase());
    return hit || 'Feedback';
}

/**
 * @param {string} raw
 */
function normalizeStatus(raw) {
    const s = String(raw || '').trim();
    const hit = STATUSES.find((c) => c.toLowerCase() === s.toLowerCase());
    return hit || 'Neu';
}

/**
 * @param {unknown} attachments
 * @returns {Array<{ name: string, contentType: string, bytes: Buffer }>}
 */
function normalizeAttachments(attachments) {
    if (!Array.isArray(attachments)) return [];
    if (attachments.length > MAX_ATTACHMENTS) {
        const err = new Error('Maximal ' + MAX_ATTACHMENTS + ' Screenshots.');
        err.status = 400;
        throw err;
    }
    const out = [];
    for (let i = 0; i < attachments.length; i++) {
        const a = attachments[i];
        if (!a || typeof a !== 'object') continue;
        const name = sanitizeFileName(String(a.name || 'screenshot-' + (i + 1) + '.png'));
        const contentType = String(a.contentType || a.type || 'image/png')
            .split(';')[0]
            .trim()
            .toLowerCase();
        if (!/^image\//.test(contentType)) {
            const err = new Error('Nur Bild-Anhänge sind erlaubt.');
            err.status = 400;
            throw err;
        }
        let bytes;
        if (Buffer.isBuffer(a.bytes)) {
            bytes = a.bytes;
        } else if (typeof a.base64 === 'string' && a.base64.trim()) {
            const b64 = a.base64.replace(/^data:[^;]+;base64,/i, '').trim();
            bytes = Buffer.from(b64, 'base64');
        } else {
            continue;
        }
        if (!bytes.length) continue;
        if (bytes.length > MAX_ATTACHMENT_BYTES) {
            const err = new Error('Screenshot „' + name + '“ ist größer als 1,5 MB.');
            err.status = 413;
            throw err;
        }
        out.push({ name, contentType, bytes });
    }
    return out;
}

/**
 * @param {string} name
 */
function sanitizeFileName(name) {
    const base = String(name || 'screenshot.png')
        .replace(/[\\/:*?"<>|]+/g, '_')
        .replace(/\s+/g, '_')
        .slice(0, 80);
    return base || 'screenshot.png';
}

/**
 * @param {string} siteId
 * @param {string} token
 */
async function resolveSiteDrive(siteId, token) {
    const drive = await graphJson(
        'GET',
        GRAPH + '/sites/' + encodeURIComponent(siteId) + '/drive?$select=id,webUrl',
        token
    );
    if (!drive || !drive.id) {
        throw new Error('Site-Dokumentbibliothek nicht gefunden.');
    }
    return drive;
}

/**
 * @param {string} driveId
 * @param {string} folderPath
 * @param {string} token
 */
async function ensureDriveFolder(driveId, folderPath, token) {
    const parts = String(folderPath || '')
        .replace(/\\/g, '/')
        .split('/')
        .filter(Boolean);
    let current = '';
    for (let i = 0; i < parts.length; i++) {
        current = current ? current + '/' + parts[i] : parts[i];
        const metaUrl =
            GRAPH +
            '/drives/' +
            encodeURIComponent(driveId) +
            '/root:/' +
            encodeDrivePath(current) +
            '?$select=id,folder';
        try {
            await graphJson('GET', metaUrl, token);
        } catch (e) {
            if (e && e.status !== 404) throw e;
            const parent = current.includes('/') ? current.slice(0, current.lastIndexOf('/')) : '';
            const createUrl = parent
                ? GRAPH +
                  '/drives/' +
                  encodeURIComponent(driveId) +
                  '/root:/' +
                  encodeDrivePath(parent) +
                  ':/children'
                : GRAPH + '/drives/' + encodeURIComponent(driveId) + '/root/children';
            await graphJson('POST', createUrl, token, {
                name: parts[i],
                folder: {},
                '@microsoft.graph.conflictBehavior': 'fail'
            }).catch(async (err) => {
                if (err && err.status === 409) return;
                // Race: Ordner parallel angelegt
                try {
                    await graphJson('GET', metaUrl, token);
                } catch {
                    throw err;
                }
            });
        }
    }
}

/**
 * @param {{
 *   title: string,
 *   description: string,
 *   category?: string,
 *   pageUrl?: string,
 *   userAgent?: string,
 *   reporterEmail: string,
 *   reporterName?: string,
 *   tenantId?: string,
 *   schoolName?: string,
 *   domain?: string,
 *   attachments?: unknown
 * }} input
 */
async function createFeedback(input) {
    const title = String((input && input.title) || '').trim();
    const description = String((input && input.description) || '').trim();
    if (!title) {
        const err = new Error('Titel fehlt.');
        err.status = 400;
        throw err;
    }
    if (!description) {
        const err = new Error('Beschreibung fehlt.');
        err.status = 400;
        throw err;
    }
    if (title.length > 200) {
        const err = new Error('Titel ist zu lang (max. 200 Zeichen).');
        err.status = 400;
        throw err;
    }
    if (description.length > 8000) {
        const err = new Error('Beschreibung ist zu lang (max. 8000 Zeichen).');
        err.status = 400;
        throw err;
    }

    const attachments = normalizeAttachments(input && input.attachments);
    const token = await getOperatorGraphToken();
    // Schema beim ersten Eintrag sicherstellen
    await ensureFeedbackSchema();
    const ids = await resolveFeedbackList(token);

    const fields = {
        Title: title,
        Beschreibung: description,
        Kategorie: normalizeCategory(input.category),
        Status: 'Neu',
        MelderEmail: String(input.reporterEmail || '').trim().slice(0, 255),
        MelderName: String(input.reporterName || '').trim().slice(0, 255),
        TenantId: String(input.tenantId || '').trim().slice(0, 64),
        Schulname: String(input.schoolName || '').trim().slice(0, 255),
        Domain: String(input.domain || '').trim().slice(0, 255),
        SeitenUrl: String(input.pageUrl || '').trim().slice(0, 2000),
        UserAgent: String(input.userAgent || '').trim().slice(0, 1000),
        AnhangUrls: '[]'
    };

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

    const itemId = String(created.id || '');
    /** @type {Array<{ name: string, url: string, path: string }>} */
    const uploaded = [];
    if (attachments.length && itemId) {
        const drive = await resolveSiteDrive(ids.siteId, token);
        const folder = FEEDBACK_FOLDER + '/' + itemId;
        await ensureDriveFolder(drive.id, folder, token);
        for (let i = 0; i < attachments.length; i++) {
            const file = attachments[i];
            const path = folder + '/' + file.name;
            const putUrl =
                GRAPH +
                '/drives/' +
                encodeURIComponent(drive.id) +
                '/root:/' +
                encodeDrivePath(path) +
                ':/content';
            const meta = await graphBinary('PUT', putUrl, token, file.bytes, file.contentType);
            uploaded.push({
                name: file.name,
                url: String(meta.webUrl || ''),
                path
            });
        }
        if (uploaded.length) {
            await graphJson(
                'PATCH',
                GRAPH +
                    '/sites/' +
                    encodeURIComponent(ids.siteId) +
                    '/lists/' +
                    encodeURIComponent(ids.listId) +
                    '/items/' +
                    encodeURIComponent(itemId) +
                    '/fields',
                token,
                { AnhangUrls: JSON.stringify(uploaded) }
            );
        }
    }

    const full = await graphJson(
        'GET',
        GRAPH +
            '/sites/' +
            encodeURIComponent(ids.siteId) +
            '/lists/' +
            encodeURIComponent(ids.listId) +
            '/items/' +
            encodeURIComponent(itemId) +
            '?$expand=fields',
        token
    );
    return mapFeedbackItem(full);
}

/**
 * @param {{ status?: string, top?: number }} [opts]
 */
async function listFeedback(opts) {
    const token = await getOperatorGraphToken();
    const ids = await resolveFeedbackList(token);
    const top = Math.min(200, Math.max(1, Number((opts && opts.top) || 100) || 100));
    const wantStatus = opts && opts.status ? normalizeStatus(opts.status) : '';
    const items = [];
    let url =
        GRAPH +
        '/sites/' +
        encodeURIComponent(ids.siteId) +
        '/lists/' +
        encodeURIComponent(ids.listId) +
        '/items?$expand=fields&$top=100';
    while (url) {
        const page = await graphJson('GET', url, token);
        const rows = page.value || [];
        for (let i = 0; i < rows.length; i++) {
            const mapped = mapFeedbackItem(rows[i]);
            if (wantStatus && mapped.status !== wantStatus) continue;
            items.push(mapped);
        }
        url = page['@odata.nextLink'] || '';
    }
    items.sort((a, b) => {
        const ta = Date.parse(a.createdAt || '') || 0;
        const tb = Date.parse(b.createdAt || '') || 0;
        return tb - ta;
    });
    return { items: items.slice(0, top), listDisplayName: feedbackListDisplayName() };
}

/**
 * @param {string} itemId
 * @param {{ status?: string }} body
 */
async function updateFeedback(itemId, body) {
    const id = String(itemId || '').trim();
    if (!id) {
        const err = new Error('Feedback-ID fehlt.');
        err.status = 400;
        throw err;
    }
    const status = normalizeStatus(body && body.status);
    const token = await getOperatorGraphToken();
    const ids = await resolveFeedbackList(token);
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
        { Status: status }
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
    return mapFeedbackItem(full);
}

/**
 * Anhang aus Site-Drive lesen (nur Pfade unter Feedback/).
 * @param {string} relPath
 * @returns {Promise<{ bytes: Buffer, contentType: string, name: string }>}
 */
async function readFeedbackFile(relPath) {
    const path = String(relPath || '')
        .replace(/\\/g, '/')
        .replace(/^\/+/, '')
        .trim();
    if (!path || !/^Feedback\//i.test(path)) {
        const err = new Error('Ungültiger Anhang-Pfad.');
        err.status = 400;
        throw err;
    }
    if (path.includes('..')) {
        const err = new Error('Ungültiger Anhang-Pfad.');
        err.status = 400;
        throw err;
    }
    const token = await getOperatorGraphToken();
    const cfg = getConfig();
    const site = await resolveSiteId(cfg.siteWebUrl, token);
    const drive = await resolveSiteDrive(site.id, token);
    const url =
        GRAPH +
        '/drives/' +
        encodeURIComponent(drive.id) +
        '/root:/' +
        encodeDrivePath(path) +
        ':/content';
    const res = await fetch(url, {
        method: 'GET',
        headers: { Authorization: 'Bearer ' + token }
    });
    const bytes = Buffer.from(await res.arrayBuffer());
    if (!res.ok) {
        const err = new Error('Anhang nicht lesbar (HTTP ' + res.status + ').');
        err.status = res.status === 404 ? 404 : 502;
        throw err;
    }
    const contentType =
        (res.headers.get('content-type') || 'application/octet-stream').split(';')[0].trim() ||
        'application/octet-stream';
    return {
        bytes,
        contentType,
        name: path.split('/').pop() || 'anhang'
    };
}

function clearFeedbackCache() {
    feedbackListCache = null;
}

module.exports = {
    FEEDBACK_LIST_DEFAULT,
    CATEGORIES,
    STATUSES,
    MAX_ATTACHMENTS,
    MAX_ATTACHMENT_BYTES,
    feedbackListDisplayName,
    feedbackColumnDefs,
    ensureFeedbackSchema,
    createFeedback,
    listFeedback,
    updateFeedback,
    readFeedbackFile,
    mapFeedbackItem,
    parseAnhangUrls,
    normalizeCategory,
    normalizeStatus,
    normalizeAttachments,
    clearFeedbackCache
};
