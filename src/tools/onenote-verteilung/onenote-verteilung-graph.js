/**
 * Graph: Teams suchen + OneNote Kursnotizbuch-Verteilung (delegiert).
 * Quelle: eigenes OneDrive (/me), zentrales Site-Notizbuch (MS365-Schultools)
 * oder License-/Katalog-API (Schul-Tenants ohne Site-Zugriff).
 */

/** Betreiber-SharePoint mit MS365-Katalog / notebooks */
export const CENTRAL_CATALOG_SITE_WEB_URL =
    'https://kurtrocks.sharepoint.com/sites/MS365-Schultools';

/** Bevorzugtes Vorlagen-Notizbuch im Ordner notebooks */
export const CENTRAL_TEMPLATE_NOTEBOOK_NAME = 'MS365-Vorlagen-Notizbuch';

export const GRAPH_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Group.Read.All',
    'https://graph.microsoft.com/Sites.Read.All',
    'https://graph.microsoft.com/Notes.ReadWrite.All'
];

function G() {
    const api = window.ms365GraphUnifiedGroups;
    if (!api) throw new Error('Graph-Modul nicht geladen.');
    return api;
}

function catalogApiConfigured() {
    const cfg = window.MS365_LICENSE_API || {};
    return !!String(cfg.baseUrl || '').trim();
}

async function catalogToken() {
    if (window.ms365LicenseApi && typeof window.ms365LicenseApi.acquireLicenseToken === 'function') {
        // User-Aktion (Vorlagen laden / Vorschau / Verteilen) → Popup für Consent erlaubt
        return window.ms365LicenseApi.acquireLicenseToken({ popup: true });
    }
    throw new Error('License-API-Client fehlt (ms365LicenseApi).');
}

function catalogClient() {
    const api = window.ms365LicenseApi;
    if (!api || typeof api.fetchCatalogOnenoteNotebooks !== 'function') {
        throw new Error('OneNote-Katalog-Client fehlt.');
    }
    return api;
}

/**
 * @returns {Promise<string>}
 */
export async function getToken() {
    if (typeof window.ms365AuthAcquireTokenPopup === 'function') {
        return window.ms365AuthAcquireTokenPopup(GRAPH_SCOPES);
    }
    if (typeof window.ms365AuthAcquireToken === 'function') {
        return window.ms365AuthAcquireToken(GRAPH_SCOPES);
    }
    return G().getGraphToken();
}

/**
 * @param {string} method
 * @param {string} path
 * @param {string} token
 * @param {object} [body]
 */
export async function graphJson(method, path, token, body) {
    return G().graphJson(method, path, token, body);
}

function guidLike(s) {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
        String(s || '').trim()
    );
}

/**
 * @param {string|null|undefined|{ kind: string, id?: string }} scope
 * @returns {{ kind: 'me'|'group'|'site'|'catalog', id: string }}
 */
function normalizeScope(scope) {
    if (!scope) return { kind: 'me', id: '' };
    if (typeof scope === 'string') {
        const id = String(scope).trim();
        return id ? { kind: 'group', id } : { kind: 'me', id: '' };
    }
    if (scope.kind === 'catalog') return { kind: 'catalog', id: '' };
    const kind = scope.kind === 'site' ? 'site' : scope.kind === 'group' ? 'group' : 'me';
    return { kind, id: String(scope.id || '').trim() };
}

/**
 * @param {{ kind: string, id: string }} scope
 * @param {string} rest path after /onenote
 */
function onenoteBase(scope, rest) {
    const s = normalizeScope(scope);
    const tail = String(rest || '').replace(/^\//, '');
    if (s.kind === 'catalog') {
        throw new Error('Katalog-Quelle nutzt die License-API, nicht Graph direkt.');
    }
    if (s.kind === 'site' && s.id) {
        return '/sites/' + encodeURIComponent(s.id) + '/onenote/' + tail;
    }
    if (s.kind === 'group' && s.id) {
        return '/groups/' + encodeURIComponent(s.id) + '/onenote/' + tail;
    }
    return '/me/onenote/' + tail;
}

/**
 * @param {string} query
 * @returns {Promise<Array<{ id: string, displayName: string, mail: string, mailNickname: string }>>}
 */
export async function searchTeams(query) {
    const q = String(query || '').trim();
    if (!q) throw new Error('Suchbegriff fehlt.');
    const token = await getToken();
    const api = G();

    if (guidLike(q)) {
        const g = await api.fetchGroup(token, q);
        if (!g || !g.id) throw new Error('Gruppe nicht gefunden.');
        return [
            {
                id: g.id,
                displayName: g.displayName || '',
                mail: g.mail || '',
                mailNickname: g.mailNickname || ''
            }
        ];
    }

    const list = await api.searchUnifiedGroups(token, q);
    return (list || []).map(function (g) {
        return {
            id: g.id,
            displayName: g.displayName || '',
            mail: g.mail || '',
            mailNickname: g.mailNickname || ''
        };
    });
}

function mapNotebook(row) {
    const lastBy =
        (row.lastModifiedBy && row.lastModifiedBy.user && row.lastModifiedBy.user.displayName) ||
        (row.lastModifiedBy && row.lastModifiedBy.application && row.lastModifiedBy.application.displayName) ||
        row.lastModifiedByName ||
        '';
    return {
        id: String(row.id || ''),
        displayName: String(row.displayName || ''),
        isDefault: !!row.isDefault,
        userRole: String(row.userRole || ''),
        links: row.links || null,
        lastModifiedDateTime: String(row.lastModifiedDateTime || row.updatedAt || '').trim(),
        lastModifiedByName: String(lastBy || '').trim()
    };
}

function mapSection(row) {
    return {
        id: String(row.id || ''),
        displayName: String(row.displayName || ''),
        pagesUrl: (row.pagesUrl && String(row.pagesUrl)) || ''
    };
}

function mapSectionGroup(row) {
    const name = String(row.displayName || '');
    const lower = name.toLowerCase();
    let kind = 'other';
    if (lower.includes('content library') || lower.includes('inhaltsbibliothek')) kind = 'contentLibrary';
    else if (lower.includes('collaboration') || lower.includes('zusammenarbeit')) kind = 'collaboration';
    else if (lower.includes('teacher only') || lower.includes('nur lehrer') || lower.includes('lehrerbereich'))
        kind = 'teacherOnly';
    return {
        id: String(row.id || ''),
        displayName: name,
        kind
    };
}

/** @type {{ id: string, displayName: string, webUrl: string }|null} */
let cachedCentralSite = null;

/**
 * @param {string} [webUrl]
 * @returns {Promise<{ id: string, displayName: string, webUrl: string }>}
 */
export async function resolveSiteByWebUrl(webUrl) {
    const raw = String(webUrl || CENTRAL_CATALOG_SITE_WEB_URL).trim().replace(/\/+$/, '');
    if (!raw) throw new Error('Site-URL fehlt.');
    let u;
    try {
        u = new URL(raw);
    } catch {
        throw new Error('Ungültige Site-URL.');
    }
    const host = u.hostname;
    const serverRel = (u.pathname || '/').replace(/\/+$/, '') || '/';
    const token = await getToken();
    const path =
        '/sites/' + encodeURIComponent(host) + ':' + serverRel + '?$select=id,displayName,webUrl';
    const data = await graphJson('GET', path, token);
    if (!data || !data.id) throw new Error('SharePoint-Site nicht gefunden: ' + raw);
    return {
        id: String(data.id),
        displayName: String(data.displayName || ''),
        webUrl: String(data.webUrl || raw)
    };
}

/**
 * Zentrale Katalog-Site (MS365-Schultools).
 */
export async function getCentralCatalogSite() {
    if (cachedCentralSite && cachedCentralSite.id) return cachedCentralSite;
    cachedCentralSite = await resolveSiteByWebUrl(CENTRAL_CATALOG_SITE_WEB_URL);
    return cachedCentralSite;
}

/**
 * SharePoint-Sites suchen (für Notizbücher auf Sites).
 * @param {string} query
 * @returns {Promise<Array<{ id: string, displayName: string, webUrl: string }>>}
 */
export async function searchSites(query) {
    const q = String(query || '').trim();
    if (!q) throw new Error('Suchbegriff fehlt.');
    const token = await getToken();

    // Direkte Site-URL → auflösen
    if (/^https?:\/\//i.test(q)) {
        const site = await resolveSiteByWebUrl(q);
        return [site];
    }

    const path =
        '/sites?search=' +
        encodeURIComponent(q) +
        '&$select=id,displayName,webUrl&$top=25';
    const data = await graphJson('GET', path, token);
    return ((data && data.value) || [])
        .map((s) => ({
            id: String(s.id || ''),
            displayName: String(s.displayName || ''),
            webUrl: String(s.webUrl || '')
        }))
        .filter((s) => s.id);
}

/**
 * @param {string|null|undefined|{ kind: string, id?: string }} [scope] me | groupId-string | {kind,id}
 */
export async function listOnenoteNotebooks(scope) {
    const token = await getToken();
    const path =
        onenoteBase(scope, 'notebooks') +
        '?$select=id,displayName,isDefault,userRole,links,lastModifiedDateTime,lastModifiedBy&$top=100';
    const data = await graphJson('GET', path, token);
    return ((data && data.value) || []).map(mapNotebook).filter((n) => n.id);
}

function preferNotebookSort(notebooks) {
    const prefer = CENTRAL_TEMPLATE_NOTEBOOK_NAME.toLowerCase();
    notebooks.sort((a, b) => {
        const ap = a.displayName.toLowerCase() === prefer ? 0 : 1;
        const bp = b.displayName.toLowerCase() === prefer ? 0 : 1;
        if (ap !== bp) return ap - bp;
        return a.displayName.localeCompare(b.displayName, 'de');
    });
    return notebooks;
}

/**
 * Notizbücher über License-API (App-Only auf Betreiber-Site).
 * @returns {Promise<{ site: object, notebooks: Array, via: 'catalog-api' }>}
 */
async function listCentralViaCatalogApi() {
    const api = catalogClient();
    const token = await catalogToken();
    const data = await api.fetchCatalogOnenoteNotebooks(token);
    const site = data.site || {
        id: '',
        displayName: 'MS365-Katalog (API)',
        webUrl: data.siteWebUrl || CENTRAL_CATALOG_SITE_WEB_URL
    };
    const notebooks = preferNotebookSort(
        (Array.isArray(data.notebooks) ? data.notebooks : []).map(mapNotebook).filter((n) => n.id)
    );
    return {
        site: {
            id: String(site.id || ''),
            displayName: String(site.displayName || 'MS365-Katalog'),
            webUrl: String(site.webUrl || data.siteWebUrl || CENTRAL_CATALOG_SITE_WEB_URL)
        },
        notebooks,
        via: data.via === 'catalog-snapshot' ? 'catalog-snapshot' : 'catalog-api',
        preferredName: data.preferredName || CENTRAL_TEMPLATE_NOTEBOOK_NAME
    };
}

/**
 * Zentrale Vorlagen: zuerst Katalog/Snapshot (schnell, kein OneNote-429),
 * Site nur auflösen für Veröffentlichen; Live-Site-OneNote nur Fallback.
 * @returns {Promise<{ site: { id: string, displayName: string, webUrl: string }, notebooks: Array, via: string }>}
 */
export async function listCentralTemplateNotebooks() {
    let siteFromGraph = null;
    try {
        siteFromGraph = await getCentralCatalogSite();
    } catch {
        siteFromGraph = null;
    }

    // 1) Katalog/Snapshot zuerst – schont OneNote Graph (429)
    if (catalogApiConfigured()) {
        try {
            const catalog = await listCentralViaCatalogApi();
            if (catalog.notebooks && catalog.notebooks.length) {
                if (siteFromGraph && siteFromGraph.id) {
                    catalog.site = siteFromGraph;
                }
                return catalog;
            }
        } catch {
            /* Live-Graph versuchen */
        }
    }

    // 2) Live Site-OneNote (kann 429 werfen → kurze Retries im Graph-Client)
    let graphErr = null;
    if (siteFromGraph && siteFromGraph.id) {
        try {
            const notebooks = preferNotebookSort(
                await listOnenoteNotebooks({ kind: 'site', id: siteFromGraph.id })
            );
            if (notebooks.length) {
                return { site: siteFromGraph, notebooks, via: 'site-graph' };
            }
            graphErr = new Error('Site-Graph: keine OneNote-Notizbücher auf der Katalog-Site.');
        } catch (e) {
            graphErr = e;
        }
    } else {
        graphErr = new Error('Katalog-Site nicht erreichbar (anderer Tenant oder fehlende Rechte).');
    }

    if (!catalogApiConfigured()) {
        throw graphErr || new Error('Zentrale Vorlagen nicht erreichbar.');
    }

    try {
        const catalog = await listCentralViaCatalogApi();
        if (siteFromGraph && siteFromGraph.id) catalog.site = siteFromGraph;
        return catalog;
    } catch (catalogErr) {
        const gmsg = String((graphErr && graphErr.message) || graphErr || '');
        const foreignTenant = /invalid hostname|tenancy|nicht gefunden|anderer Tenant/i.test(gmsg);
        const st = catalogErr.status ? Number(catalogErr.status) : 0;
        const detail =
            catalogErr.message ||
            (catalogErr.payload && (catalogErr.payload.error || catalogErr.payload.message)) ||
            'Katalog-API abgelehnt';
        const code = catalogErr.code || (catalogErr.payload && catalogErr.payload.code) || '';

        if (foreignTenant || /429|Too Many Requests/i.test(gmsg)) {
            let tip =
                'Details im Protokoll; oft fehlt License.Access-Consent oder der Snapshot.';
            if (/429|Too Many Requests/i.test(gmsg)) {
                tip =
                    'Microsoft OneNote-API überlastet (429). Kurz warten oder Snapshot nutzen – erneut „Vorlagen laden“.';
            } else if (
                code === 'onenote_app_only_blocked' ||
                /app-only|Snapshot|Für Schulen veröffentlichen|Kein OneNote-Snapshot/i.test(
                    String(detail)
                )
            ) {
                tip =
                    'Snapshot fehlt oder war noch nicht fertig. In kurtrocks veröffentlichen, Toast abwarten, hier Seite neu laden.';
            } else if (
                code === 'missing_license_scope' ||
                code === 'graph_token' ||
                code === 'wrong_audience'
            ) {
                tip =
                    'Consent „Lizenz-API und Katalog lesen“ fehlt. Abmelden, neu anmelden, Dialog bestätigen.';
            } else if (st === 403) {
                tip = 'Schul-Tenant in der Lizenzliste freischalten (Status trial oder active).';
            } else if (st === 401) {
                tip =
                    'License.Access-Token fehlt oder ungültig. Abmelden, neu anmelden, Consent bestätigen.';
            }
            throw new Error(detail + (code ? ' [' + code + ']' : '') + ' · ' + tip);
        }

        throw new Error(
            'Katalog-API: ' +
                ((catalogErr && catalogErr.message) || catalogErr) +
                ' · Site-Graph: ' +
                gmsg
        );
    }
}

async function listOnenoteSectionGroups(notebookId, scope) {
    const nid = String(notebookId || '').trim();
    if (!nid) throw new Error('Notizbuch-ID fehlt.');
    const token = await getToken();
    const path =
        onenoteBase(scope, 'notebooks/' + encodeURIComponent(nid) + '/sectionGroups') +
        '?$select=id,displayName&$top=100';
    const data = await graphJson('GET', path, token);
    return ((data && data.value) || []).map(mapSectionGroup).filter((g) => g.id);
}

async function listOnenoteNotebookSections(notebookId, scope) {
    const nid = String(notebookId || '').trim();
    if (!nid) throw new Error('Notizbuch-ID fehlt.');
    const token = await getToken();
    const path =
        onenoteBase(scope, 'notebooks/' + encodeURIComponent(nid) + '/sections') +
        '?$select=id,displayName&$top=100';
    const data = await graphJson('GET', path, token);
    return ((data && data.value) || []).map(mapSection).filter((s) => s.id);
}

async function listOnenoteSectionGroupSections(sectionGroupId, scope) {
    const sid = String(sectionGroupId || '').trim();
    if (!sid) throw new Error('Abschnittsgruppen-ID fehlt.');
    const token = await getToken();
    const path =
        onenoteBase(scope, 'sectionGroups/' + encodeURIComponent(sid) + '/sections') +
        '?$select=id,displayName&$top=100';
    const data = await graphJson('GET', path, token);
    return ((data && data.value) || []).map(mapSection).filter((s) => s.id);
}

/**
 * @param {string} notebookId
 * @param {string|null|undefined|{ kind: string, id?: string }} [scope]
 */
export async function loadOnenoteNotebookTree(notebookId, scope) {
    const nid = String(notebookId || '').trim();
    if (!nid) throw new Error('Notizbuch-ID fehlt.');
    const sc = normalizeScope(scope);
    if (sc.kind === 'catalog') {
        const api = catalogClient();
        const token = await catalogToken();
        const data = await api.fetchCatalogOnenoteTree(token, nid);
        return {
            sections: (Array.isArray(data.sections) ? data.sections : []).map(mapSection).filter((s) => s.id),
            groups: (Array.isArray(data.groups) ? data.groups : []).map((g) =>
                Object.assign({}, mapSectionGroup(g), {
                    sections: (Array.isArray(g.sections) ? g.sections : [])
                        .map(mapSection)
                        .filter((s) => s.id)
                })
            )
        };
    }
    // Sequentiell: parallele OneNote-Calls triggern schnell 429
    const sections = await listOnenoteNotebookSections(nid, sc);
    await G().sleep(250);
    const groups = await listOnenoteSectionGroups(nid, sc);
    const withSections = [];
    for (const g of groups) {
        let childSections = [];
        try {
            await G().sleep(200);
            childSections = await listOnenoteSectionGroupSections(g.id, sc);
        } catch {
            childSections = [];
        }
        withSections.push(Object.assign({}, g, { sections: childSections }));
    }
    return { sections, groups: withSections };
}

async function pollOnenoteOperation(operationUrl, token, onProgress) {
    const api = G();
    const started = Date.now();
    const maxMs = 120000;
    let delay = 1500;
    while (Date.now() - started < maxMs) {
        const data = await api.graphJson('GET', operationUrl, token);
        const status = String((data && data.status) || '').toLowerCase();
        if (onProgress) onProgress(data);
        if (status === 'completed') return data;
        if (status === 'failed') {
            const err =
                (data && data.error && (data.error.message || JSON.stringify(data.error))) ||
                'OneNote-Kopieren fehlgeschlagen.';
            throw new Error(err);
        }
        await api.sleep(delay);
        delay = Math.min(5000, delay + 500);
    }
    throw new Error('OneNote-Kopieren: Zeitüberschreitung (Status wird asynchron weiterlaufen).');
}

/** Max. Bildgröße beim Einbetten (Bytes); darüber Platzhalter. */
const SNAP_MAX_IMAGE_BYTES = 900_000;
/** Max. Dateianhang beim Einbetten (Bytes). */
const SNAP_MAX_FILE_BYTES = 450_000;
/** Max. HTML-Länge einer Snapshot-Seite nach Medien-Inlining. */
const SNAP_MAX_PAGE_HTML = 4_000_000;

const EMBED_PLACEHOLDER_HTML =
    '<p style="margin:0.75em 0;padding:0.65em 0.85em;border:1px solid #c9a227;background:#fff8e1;color:#333;font-family:Segoe UI,sans-serif;font-size:14px;">' +
    '<strong>Eingebetteter Inhalt nicht übernommen</strong> ' +
    '(Forms, Learning Activity, Stream o. Ä.). Bitte in OneNote manuell neu einfügen.' +
    '</p>';

/**
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function bytesToBase64(bytes) {
    let binary = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
}

/**
 * @param {string} url
 * @returns {boolean}
 */
function isOnenoteResourceUrl(url) {
    const u = String(url || '').trim();
    if (!u || u.startsWith('data:')) return false;
    try {
        const parsed = new URL(u);
        const host = parsed.hostname.toLowerCase();
        if (host === 'graph.microsoft.com') {
            return /\/onenote\/resources\//i.test(parsed.pathname) || /\/resources\//i.test(parsed.pathname);
        }
        if (host === 'www.onenote.com' || host.endsWith('.onenote.com')) {
            return /\/resources\//i.test(parsed.pathname) || /\/\$value/i.test(parsed.pathname);
        }
    } catch {
        return false;
    }
    return false;
}

/**
 * OneNote-HTML liefert oft /siteCollections/… und unkodierte „!“ in Resource-IDs → Graph 400.
 * @param {string} url
 * @returns {string[]}
 */
function onenoteResourceUrlCandidates(url) {
    const raw = String(url || '').trim();
    if (!raw) return [];
    const out = [];
    const push = (u) => {
        if (u && !out.includes(u)) out.push(u);
    };
    push(raw);
    try {
        const u = new URL(raw);
        // siteCollections → sites (Graph v1)
        let path = u.pathname.replace(/\/siteCollections\//gi, '/sites/');
        // Resource-ID-Segment korrekt enkodieren (! → %21)
        path = path.replace(/\/onenote\/resources\/([^/]+)(\/|$)/i, (_, id, tail) => {
            let decoded = id;
            try {
                decoded = decodeURIComponent(id);
            } catch {
                /* keep */
            }
            return '/onenote/resources/' + encodeURIComponent(decoded) + (tail || '');
        });
        u.pathname = path;
        push(u.toString());

        const m = path.match(/\/onenote\/resources\/([^/]+)/i);
        const resourceId = m ? m[1] : '';
        if (resourceId) {
            push(
                'https://graph.microsoft.com/v1.0/me/onenote/resources/' + resourceId + '/$value'
            );
            const siteMatch = path.match(/\/sites\/([^/]+)\//i);
            if (siteMatch) {
                let siteId = siteMatch[1];
                try {
                    siteId = decodeURIComponent(siteId);
                } catch {
                    /* keep */
                }
                push(
                    'https://graph.microsoft.com/v1.0/sites/' +
                        encodeURIComponent(siteId) +
                        '/onenote/resources/' +
                        resourceId +
                        '/$value'
                );
            }
            if (cachedCentralSite && cachedCentralSite.id) {
                push(
                    'https://graph.microsoft.com/v1.0/sites/' +
                        encodeURIComponent(cachedCentralSite.id) +
                        '/onenote/resources/' +
                        resourceId +
                        '/$value'
                );
            }
        }
    } catch {
        /* ignore */
    }
    return out;
}

/**
 * @param {string} src
 * @param {string} [tagLower]
 */
function isNonCopyableEmbed(src, tagLower) {
    const s = String(src || '').toLowerCase();
    const t = String(tagLower || '').toLowerCase();
    return (
        /forms\.office\.com|forms\.microsoft\.com/.test(s) ||
        /education\.microsoft\.com|learningtools|learning.?activit/.test(s + t) ||
        /microsoftstream|web\.microsoftstream|stream\.microsoft|stream\.azure/.test(s) ||
        /data-plugin|officeforms|onenote\.embedded/.test(t) ||
        (/iframe|object/.test(t) && /embed|player|widget/.test(s) && /microsoft|office|forms|stream/.test(s))
    );
}

/** @type {Map<string, { bytes: Uint8Array, contentType: string }|null>} */
const resourceFetchCache = new Map();

/**
 * @param {string} url
 * @param {string} token
 * @param {number} maxBytes
 * @returns {Promise<{ bytes: Uint8Array, contentType: string }|null>}
 */
async function fetchOnenoteResource(url, token, maxBytes) {
    if (!isOnenoteResourceUrl(url)) return null;
    const cacheKey = String(url || '').trim();
    if (resourceFetchCache.has(cacheKey)) return resourceFetchCache.get(cacheKey);

    const candidates = onenoteResourceUrlCandidates(url);
    let packed = null;
    for (const candidate of candidates) {
        try {
            const res = await fetch(candidate, {
                method: 'GET',
                headers: { Authorization: 'Bearer ' + token }
            });
            if (!res.ok) continue;
            const buf = new Uint8Array(await res.arrayBuffer());
            if (!buf.length || buf.length > maxBytes) continue;
            let contentType = String(res.headers.get('Content-Type') || '')
                .split(';')[0]
                .trim()
                .toLowerCase();
            if (!contentType || contentType === 'application/octet-stream') {
                contentType = 'image/jpeg';
            }
            packed = { bytes: buf, contentType };
            break;
        } catch {
            /* next candidate */
        }
    }
    resourceFetchCache.set(cacheKey, packed);
    // gleiche Resource-ID unter anderen URLs ebenfalls merken
    for (const c of candidates) {
        if (!resourceFetchCache.has(c)) resourceFetchCache.set(c, packed);
    }
    return packed;
}

/**
 * Snapshot-HTML: OneNote-Ressourcen (Bilder/Dateien) per Publisher-Token einbetten;
 * Forms/Learning Activities → sichtbarer Platzhalter.
 * @param {string} html
 * @param {string} token
 * @returns {Promise<{ html: string, stats: { imagesInlined: number, imagesSkipped: number, filesInlined: number, embedsReplaced: number } }>}
 */
async function prepareSnapshotPageHtml(html, token) {
    const stats = {
        imagesInlined: 0,
        imagesSkipped: 0,
        filesInlined: 0,
        embedsReplaced: 0
    };
    let out = String(html || '');
    if (!out) return { html: out, stats };

    // iframe / object-Embeds die Graph nicht sinnvoll kopiert
    out = out.replace(/<(iframe|object)\b[^>]*>[\s\S]*?<\/\1>/gi, (full) => {
        const lower = full.toLowerCase();
        const srcMatch =
            full.match(/\bsrc\s*=\s*["']([^"']+)["']/i) ||
            full.match(/\bdata\s*=\s*["']([^"']+)["']/i);
        const src = srcMatch ? srcMatch[1] : '';
        if (isNonCopyableEmbed(src, lower) || /<(iframe)\b/i.test(full)) {
            // Nur verdächtige / typische Embeds ersetzen; reine OneNote-object-Dateien unten behandeln
            if (
                isNonCopyableEmbed(src, lower) ||
                (/iframe/i.test(full) && !isOnenoteResourceUrl(src))
            ) {
                stats.embedsReplaced++;
                return EMBED_PLACEHOLDER_HTML;
            }
        }
        return full;
    });

    // Bilder: data-fullres-src bevorzugen, sonst src
    const imgTagRe = /<img\b[^>]*>/gi;
    const imgTags = out.match(imgTagRe) || [];
    const cache = new Map();

    for (const tag of imgTags) {
        if (out.length > SNAP_MAX_PAGE_HTML) {
            stats.imagesSkipped++;
            continue;
        }
        const fullRes = (tag.match(/\bdata-fullres-src\s*=\s*["']([^"']+)["']/i) || [])[1] || '';
        const src = (tag.match(/\bsrc\s*=\s*["']([^"']+)["']/i) || [])[1] || '';
        const candidate = fullRes || src;
        if (!candidate || candidate.startsWith('data:')) continue;
        if (!isOnenoteResourceUrl(candidate)) {
            // Fremd-URL belassen (öffentliche https://…-Bilder funktionieren oft beim Recreate)
            continue;
        }

        let packed = cache.get(candidate);
        if (packed === undefined) {
            try {
                packed = await fetchOnenoteResource(candidate, token, SNAP_MAX_IMAGE_BYTES);
            } catch {
                packed = null;
            }
            cache.set(candidate, packed);
            await G().sleep(80);
        }
        if (!packed) {
            stats.imagesSkipped++;
            const broken =
                tag.replace(/\bsrc\s*=\s*["'][^"']*["']/i, 'src=""') +
                '<!-- Snapshot: Bildressource nicht ladbar -->';
            out = out.replace(tag, broken);
            continue;
        }
        const dataUri = 'data:' + packed.contentType + ';base64,' + bytesToBase64(packed.bytes);
        let next = tag.replace(/\bsrc\s*=\s*["'][^"']*["']/i, 'src="' + dataUri + '"');
        if (!/\bsrc\s*=/i.test(tag)) {
            next = tag.replace(/<img\b/i, '<img src="' + dataUri + '"');
        }
        next = next.replace(/\s*data-fullres-src\s*=\s*["'][^"']*["']/gi, '');
        next = next.replace(/\s*data-fullres-src-type\s*=\s*["'][^"']*["']/gi, '');
        out = out.replace(tag, next);
        stats.imagesInlined++;
    }

    // Dateianhänge (<object data="…/resources/…">)
    const objTagRe = /<object\b[^>]*(?:\/>|>[\s\S]*?<\/object>)/gi;
    const objTags = out.match(objTagRe) || [];
    for (const tag of objTags) {
        if (out.length > SNAP_MAX_PAGE_HTML) break;
        const dataUrl = (tag.match(/\bdata\s*=\s*["']([^"']+)["']/i) || [])[1] || '';
        if (!isOnenoteResourceUrl(dataUrl)) continue;
        let packed = cache.get(dataUrl);
        if (packed === undefined) {
            try {
                packed = await fetchOnenoteResource(dataUrl, token, SNAP_MAX_FILE_BYTES);
            } catch {
                packed = null;
            }
            cache.set(dataUrl, packed);
            await G().sleep(80);
        }
        if (!packed) {
            stats.embedsReplaced++;
            out = out.replace(
                tag,
                EMBED_PLACEHOLDER_HTML.replace(
                    'Eingebetteter Inhalt nicht übernommen',
                    'Dateianhang nicht übernommen'
                )
            );
            continue;
        }
        const dataUri = 'data:' + packed.contentType + ';base64,' + bytesToBase64(packed.bytes);
        const next = tag.replace(/\bdata\s*=\s*["'][^"']*["']/i, 'data="' + dataUri + '"');
        out = out.replace(tag, next);
        stats.filesInlined++;
    }

    if (out.length > SNAP_MAX_PAGE_HTML) {
        out = out.slice(0, SNAP_MAX_PAGE_HTML);
    }
    return { html: out, stats };
}

/**
 * data:-URIs aus HTML in Multipart-Parts zerlegen (zuverlässiger als riesige data-URIs im Body).
 * @param {string} html
 * @returns {{ html: string, parts: Array<{ name: string, contentType: string, bytes: Uint8Array }> }}
 */
function extractDataUrisForMultipart(html) {
    const parts = [];
    let n = 0;
    const nextHtml = String(html || '').replace(
        /\b(src|data)\s*=\s*["']data:([^;"']+);base64,([A-Za-z0-9+/=\s]+)["']/gi,
        (full, attr, mime, b64) => {
            n++;
            const name = 'mediaBlock' + n;
            const clean = String(b64).replace(/\s+/g, '');
            let bytes;
            try {
                const bin = atob(clean);
                bytes = new Uint8Array(bin.length);
                for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
            } catch {
                return full;
            }
            parts.push({
                name,
                contentType: String(mime || 'application/octet-stream').trim(),
                bytes
            });
            return attr + '="name:' + name + '"';
        }
    );
    return { html: nextHtml, parts };
}

/**
 * HTML-Seite im Ziel-Abschnitt anlegen (Schul-Token).
 * Mit eingebetteten data:-Medien → multipart (Graph-Empfehlung).
 * @param {string} groupId
 * @param {string} sectionId
 * @param {string} html
 * @param {string} token
 */
async function createPageHtmlInGroupSection(groupId, sectionId, html, token) {
    const url =
        'https://graph.microsoft.com/v1.0/groups/' +
        encodeURIComponent(groupId) +
        '/onenote/sections/' +
        encodeURIComponent(sectionId) +
        '/pages';

    const extracted = extractDataUrisForMultipart(html);
    let body;
    let contentType;

    if (extracted.parts.length) {
        const boundary = 'OneNoteSnap' + Date.now() + 'x' + Math.random().toString(36).slice(2, 8);
        const chunks = [];
        const enc = new TextEncoder();
        const pushStr = (s) => chunks.push(enc.encode(s));
        pushStr('--' + boundary + '\r\n');
        pushStr('Content-Disposition: form-data; name="Presentation"\r\n');
        pushStr('Content-Type: text/html\r\n\r\n');
        pushStr(extracted.html);
        pushStr('\r\n');
        for (const part of extracted.parts) {
            pushStr('--' + boundary + '\r\n');
            pushStr('Content-Disposition: form-data; name="' + part.name + '"\r\n');
            pushStr('Content-Type: ' + part.contentType + '\r\n\r\n');
            chunks.push(part.bytes);
            pushStr('\r\n');
        }
        pushStr('--' + boundary + '--\r\n');
        let total = 0;
        for (const c of chunks) total += c.length;
        const merged = new Uint8Array(total);
        let off = 0;
        for (const c of chunks) {
            merged.set(c, off);
            off += c.length;
        }
        body = merged;
        contentType = 'multipart/form-data; boundary=' + boundary;
    } else {
        body = html;
        contentType = 'text/html';
    }

    const res = await fetch(url, {
        method: 'POST',
        headers: {
            Authorization: 'Bearer ' + token,
            'Content-Type': contentType,
            Accept: 'application/json'
        },
        body
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
        // Fallback: nochmal als reines HTML mit data-URIs (manche Tenants zickt bei Multipart)
        if (extracted.parts.length && res.status >= 400) {
            const retry = await fetch(url, {
                method: 'POST',
                headers: {
                    Authorization: 'Bearer ' + token,
                    'Content-Type': 'text/html',
                    Accept: 'application/json'
                },
                body: html
            });
            const retryText = await retry.text();
            let retryData = null;
            try {
                retryData = JSON.parse(retryText);
            } catch {
                retryData = { raw: retryText };
            }
            if (retry.ok) return retryData;
            const retryMsg =
                retryData && retryData.error && retryData.error.message
                    ? retryData.error.message
                    : retryText || 'HTTP ' + retry.status;
            throw new Error('Seite anlegen: ' + retryMsg);
        }
        throw new Error('Seite anlegen: ' + msg);
    }
    return data;
}

/**
 * Abschnitt aus Katalog-Export im Ziel neu anlegen (cross-tenant).
 * @param {string} sourceSectionId
 * @param {{ sectionGroupId: string, groupId: string, renameAs?: string }} dest
 * @param {(info: object) => void} [onProgress]
 */
async function recreateSectionFromCatalog(sourceSectionId, dest, onProgress) {
    const sectionId = String(sourceSectionId || '').trim();
    const sectionGroupId = String((dest && dest.sectionGroupId) || '').trim();
    const groupId = String((dest && dest.groupId) || '').trim();
    if (!sectionId) throw new Error('Quell-Abschnitt fehlt.');
    if (!sectionGroupId) throw new Error('Ziel-Abschnittsgruppe fehlt.');
    if (!groupId) throw new Error('Ziel-Team/Gruppe fehlt.');

    const renameAs = dest && dest.renameAs ? String(dest.renameAs).trim() : '';
    if (onProgress) onProgress({ status: 'export' });

    const api = catalogClient();
    const licToken = await catalogToken();
    const exported = await api.fetchCatalogOnenoteSectionExport(licToken, sectionId);
    const pages = Array.isArray(exported.pages) ? exported.pages : [];
    const okPages = pages.filter((p) => p && p.html);
    if (!okPages.length) {
        const firstErr = pages.find((p) => p && p.error);
        throw new Error(
            firstErr && firstErr.error
                ? 'Katalog-Export: ' +
                      firstErr.error +
                      ' Hinweis: In kurtrocks das Notizbuch erneut „Veröffentlichen“ (Seiten mit Bildern/Karten brauchen einen frischen Snapshot).'
                : 'Katalog-Export: keine Seiten-HTML erhalten. In kurtrocks erneut veröffentlichen.'
        );
    }

    if (onProgress) onProgress({ status: 'createSection' });
    const graphToken = await getToken();
    const displayName = renameAs || 'Vorlage';
    const created = await graphJson(
        'POST',
        '/groups/' +
            encodeURIComponent(groupId) +
            '/onenote/sectionGroups/' +
            encodeURIComponent(sectionGroupId) +
            '/sections',
        graphToken,
        { displayName }
    );
    const newSectionId = created && created.id ? String(created.id) : '';
    if (!newSectionId) throw new Error('Ziel-Abschnitt konnte nicht angelegt werden.');

    const sleep = G().sleep;
    let done = 0;
    for (const page of okPages) {
        done++;
        if (onProgress) {
            onProgress({
                status: 'createPage ' + done + '/' + okPages.length,
                title: page.title || ''
            });
        }
        await createPageHtmlInGroupSection(groupId, newSectionId, page.html, graphToken);
        if (done < okPages.length) await sleep(400);
    }

    const skipped = pages.length - okPages.length;
    return {
        status: 'completed',
        via: 'catalog-recreate',
        sectionId: newSectionId,
        pageCount: okPages.length,
        skipped
    };
}

/**
 * Abschnitt (aus /me, Site oder Katalog) in eine Abschnittsgruppe eines Teams kopieren.
 * Katalog: HTML-Export → Abschnitt + Seiten im Schul-Tenant neu anlegen.
 * @param {string} sourceSectionId
 * @param {{ sectionGroupId: string, groupId: string, renameAs?: string }} dest
 * @param {(info: object) => void} [onProgress]
 * @param {string|null|undefined|{ kind: string, id?: string }} [sourceScope]
 */
export async function copySectionToGroupSectionGroup(
    sourceSectionId,
    dest,
    onProgress,
    sourceScope
) {
    const sc = normalizeScope(sourceScope);
    if (sc.kind === 'catalog') {
        return recreateSectionFromCatalog(sourceSectionId, dest, onProgress);
    }

    const sectionId = String(sourceSectionId || '').trim();
    const sectionGroupId = String((dest && dest.sectionGroupId) || '').trim();
    const groupId = String((dest && dest.groupId) || '').trim();
    if (!sectionId) throw new Error('Quell-Abschnitt fehlt.');
    if (!sectionGroupId) throw new Error('Ziel-Abschnittsgruppe fehlt.');
    if (!groupId) throw new Error('Ziel-Team/Gruppe fehlt.');

    const token = await getToken();
    const api = G();
    const path = onenoteBase(sc, 'sections/' + encodeURIComponent(sectionId) + '/copyToSectionGroup');
    const body = {
        id: sectionGroupId,
        groupId: groupId
    };
    const renameAs = dest && dest.renameAs ? String(dest.renameAs).trim() : '';
    if (renameAs) body.renameAs = renameAs;

    const res = await api.graphRequest('POST', path, token, body);
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = { raw: text };
        }
    }
    if (res.status !== 202 && !res.ok) {
        const msg =
            data && data.error && data.error.message
                ? data.error.message
                : text || 'HTTP ' + res.status;
        throw new Error('copyToSectionGroup: ' + msg);
    }
    const opUrl =
        res.headers.get('Operation-Location') ||
        res.headers.get('Location') ||
        (data && data['@odata.id']) ||
        '';
    if (!opUrl) {
        return { status: 'accepted', message: 'Kopieren gestartet (ohne Operation-Location).' };
    }
    return pollOnenoteOperation(opUrl, token, onProgress);
}

/** @deprecated Alias – nutzt /me als Quelle */
export async function copyMySectionToGroupSectionGroup(sourceSectionId, dest, onProgress) {
    return copySectionToGroupSectionGroup(sourceSectionId, dest, onProgress, { kind: 'me' });
}

/**
 * Seiten eines Abschnitts (für Vorschau-Liste).
 * @param {string} sectionId
 * @param {string|null|undefined|{ kind: string, id?: string }} [scope]
 */
export async function listSectionPages(sectionId, scope) {
    const sid = String(sectionId || '').trim();
    if (!sid) throw new Error('Abschnitt-ID fehlt.');
    const sc = normalizeScope(scope);
    if (sc.kind === 'catalog') {
        const api = catalogClient();
        const token = await catalogToken();
        const data = await api.fetchCatalogOnenoteSectionPages(token, sid);
        return (Array.isArray(data.pages) ? data.pages : [])
            .map((p) => ({
                id: String(p.id || ''),
                title: String(p.title || 'Ohne Titel'),
                webUrl: p.webUrl || ''
            }))
            .filter((p) => p.id);
    }
    const token = await getToken();
    const path =
        onenoteBase(sc, 'sections/' + encodeURIComponent(sid) + '/pages') +
        '?$top=25&$select=id,title,createdDateTime,lastModifiedDateTime,links&$orderby=lastModifiedDateTime desc';
    const data = await graphJson('GET', path, token);
    return ((data && data.value) || [])
        .map((p) => ({
            id: String(p.id || ''),
            title: String(p.title || 'Ohne Titel'),
            webUrl:
                (p.links && p.links.oneNoteWebUrl && p.links.oneNoteWebUrl.href) ||
                ''
        }))
        .filter((p) => p.id);
}

/**
 * Leichte Text-/Bildvorschau einer Seite.
 * @param {string} pageId
 * @param {string|null|undefined|{ kind: string, id?: string }} [scope]
 * @returns {Promise<{ previewText: string, previewImageUrl: string, links?: object }>}
 */
export async function getPagePreview(pageId, scope) {
    const pid = String(pageId || '').trim();
    if (!pid) throw new Error('Seiten-ID fehlt.');
    const sc = normalizeScope(scope);
    if (sc.kind === 'catalog') {
        const api = catalogClient();
        const token = await catalogToken();
        const data = await api.fetchCatalogOnenotePagePreview(token, pid);
        return {
            previewText: String((data && data.previewText) || '').trim(),
            previewImageUrl: String((data && data.previewImageUrl) || '').trim(),
            links: null
        };
    }
    const token = await getToken();
    const path = onenoteBase(sc, 'pages/' + encodeURIComponent(pid) + '/preview');
    const data = await graphJson('GET', path, token);
    const img =
        (data &&
            data.links &&
            data.links.previewImageUrl &&
            (data.links.previewImageUrl.href || data.links.previewImageUrl)) ||
        '';
    return {
        previewText: String((data && data.previewText) || '').trim(),
        previewImageUrl: String(img || '').trim(),
        links: (data && data.links) || null
    };
}

/**
 * HTML-Inhalt einer Seite (für reichhaltige Vorschau inkl. Tabellen).
 * @param {string} pageId
 * @param {string|null|undefined|{ kind: string, id?: string }} [scope]
 * @returns {Promise<{ html: string }>}
 */
export async function getPageContent(pageId, scope) {
    const pid = String(pageId || '').trim();
    if (!pid) throw new Error('Seiten-ID fehlt.');
    const sc = normalizeScope(scope);
    if (sc.kind === 'catalog') {
        const api = catalogClient();
        if (typeof api.fetchCatalogOnenotePageContent !== 'function') {
            throw new Error('Katalog-Content-Client fehlt.');
        }
        const token = await catalogToken();
        const data = await api.fetchCatalogOnenotePageContent(token, pid);
        return { html: String((data && data.html) || '') };
    }
    const token = await getToken();
    const path = onenoteBase(sc, 'pages/' + encodeURIComponent(pid) + '/content');
    const url = 'https://graph.microsoft.com/v1.0' + path;
    const res = await fetch(url, {
        method: 'GET',
        headers: {
            Authorization: 'Bearer ' + token,
            Accept: 'text/html'
        }
    });
    const html = await res.text();
    if (!res.ok) {
        let msg = html || 'HTTP ' + res.status;
        try {
            const data = JSON.parse(html);
            if (data && data.error && data.error.message) msg = data.error.message;
        } catch {
            /* ignore */
        }
        throw new Error(msg);
    }
    return { html };
}

/**
 * Kursnotizbuch heuristisch wählen (Team-Name / „Notizbuch“).
 * @param {Array<{ id: string, displayName: string }>} notebooks
 * @param {string} [teamName]
 */
export function pickClassNotebook(notebooks, teamName) {
    const list = Array.isArray(notebooks) ? notebooks : [];
    if (!list.length) return null;
    const tn = String(teamName || '')
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .trim();
    const tip = tn.slice(0, Math.min(18, tn.length));
    let best = null;
    let bestScore = -1;
    for (const n of list) {
        const d = String(n.displayName || '').toLowerCase();
        let s = 0;
        if (d.includes('notizbuch') || d.includes('notebook') || d.includes('class notebook')) s += 4;
        if (d.includes('kurs')) s += 1;
        if (tip && tip.length >= 4 && d.includes(tip)) s += 5;
        if (s > bestScore) {
            bestScore = s;
            best = n;
        }
    }
    return bestScore > 0 ? best : list[0];
}

/**
 * Abschnittsgruppe nach Art wählen (Inhaltsbibliothek bevorzugt).
 * @param {Array<{ id: string, displayName: string, kind: string }>} groups
 * @param {'contentLibrary'|'teacherOnly'|'collaboration'|'any'} [prefer]
 */
export function pickSectionGroup(groups, prefer) {
    const list = Array.isArray(groups) ? groups : [];
    if (!list.length) return null;
    const want = prefer || 'contentLibrary';
    if (want !== 'any') {
        const hit = list.find((g) => g.kind === want);
        if (hit) return hit;
    }
    return (
        list.find((g) => g.kind === 'contentLibrary') ||
        list.find((g) => g.kind === 'teacherOnly') ||
        list[0]
    );
}

/**
 * Welche Notizbücher bereits im Schul-Snapshot stehen (Katalog-API).
 * @returns {Promise<Array<{ id: string, displayName: string, lastModifiedDateTime?: string, publishedAt?: string, publishedBy?: string }>>}
 */
export async function listPublishedSnapshotNotebooks() {
    if (!catalogApiConfigured()) return [];
    try {
        const api = catalogClient();
        const token = await catalogToken();
        const data = await api.fetchCatalogOnenoteNotebooks(token);
        return (Array.isArray(data.notebooks) ? data.notebooks : [])
            .map((n) => ({
                id: String(n.id || ''),
                displayName: String(n.displayName || ''),
                lastModifiedDateTime: String(n.lastModifiedDateTime || data.updatedAt || '').trim(),
                publishedAt: String(n.publishedAt || data.updatedAt || '').trim(),
                publishedBy: String(n.publishedBy || data.publishedBy || '').trim()
            }))
            .filter((n) => n.id);
    } catch {
        return [];
    }
}

/**
 * Ein oder mehrere zentrale Notizbücher (Site-Graph) als Snapshot für Schulen veröffentlichen.
 * @param {Array|{ id: string, displayName: string }} notebooks
 * @param {{ kind: string, id?: string }} scope
 * @param {(info: { phase: string, detail?: string }) => void} [onProgress]
 */
export async function publishCentralNotebookSnapshot(notebooks, scope, onProgress) {
    const list = (Array.isArray(notebooks) ? notebooks : [notebooks]).filter((n) => n && n.id);
    if (!list.length) throw new Error('Notizbuch fehlt.');
    const sc = normalizeScope(scope);
    if (sc.kind === 'catalog') {
        throw new Error('Snapshot braucht Site-/Graph-Zugriff (kurtrocks), nicht den Katalog-Pfad.');
    }
    const api = catalogClient();
    if (typeof api.publishCatalogOnenoteSnapshot !== 'function') {
        throw new Error('Snapshot-Publish-Client fehlt.');
    }

    resourceFetchCache.clear();
    try {
        await getCentralCatalogSite();
    } catch {
        /* optional für Resource-URL-Fallback */
    }

    const trees = {};
    const sectionsPayload = {};
    const notebookMeta = [];
    let sectionTotal = 0;
    let sectionDone = 0;

    for (const nb of list) {
        if (onProgress) onProgress({ phase: 'tree', detail: nb.displayName || nb.id });
        const tree = await loadOnenoteNotebookTree(nb.id, sc);
        trees[nb.id] = {
            sections: tree.sections || [],
            groups: tree.groups || []
        };
        const sectionList = [];
        (tree.sections || []).forEach((s) => sectionList.push(s));
        (tree.groups || []).forEach((g) => {
            (g.sections || []).forEach((s) => sectionList.push(s));
        });
        sectionTotal += sectionList.length;

        for (const sec of sectionList) {
            sectionDone++;
            if (onProgress) {
                onProgress({
                    phase: 'section',
                    detail:
                        (nb.displayName || 'Buch') +
                        ' · ' +
                        sectionDone +
                        '/' +
                        sectionTotal +
                        ' · ' +
                        (sec.displayName || '')
                });
            }
            const pages = await listSectionPages(sec.id, sc);
            const pageRows = [];
            const graphToken = await getToken();
            for (const p of pages) {
                let html = '';
                let previewText = '';
                let mediaStats = null;
                for (let attempt = 0; attempt < 2 && !html; attempt++) {
                    try {
                        if (attempt) await G().sleep(500);
                        const content = await getPageContent(p.id, sc);
                        html = (content && content.html) || '';
                        if (html) {
                            const prepared = await prepareSnapshotPageHtml(html, graphToken);
                            html = prepared.html;
                            mediaStats = prepared.stats;
                        }
                    } catch (e) {
                        previewText = String((e && e.message) || e);
                    }
                }
                if (html && html.length > SNAP_MAX_PAGE_HTML) {
                    html = html.slice(0, SNAP_MAX_PAGE_HTML);
                }
                if (!html) {
                    previewText =
                        previewText ||
                        'Kein Seiten-HTML von Graph – Seite wird ohne Inhalt markiert.';
                }
                pageRows.push({
                    id: p.id,
                    title: p.title,
                    html,
                    previewText: previewText || html.replace(/<[^>]+>/g, ' ').slice(0, 400),
                    webUrl: p.webUrl || '',
                    media: mediaStats || undefined
                });
                await G().sleep(200);
            }
            sectionsPayload[sec.id] = {
                displayName: sec.displayName || '',
                pages: pageRows
            };
        }

        notebookMeta.push({
            id: nb.id,
            displayName: nb.displayName || '',
            lastModifiedDateTime: nb.lastModifiedDateTime || new Date().toISOString(),
            lastModifiedByName: nb.lastModifiedByName || ''
        });
    }

    if (onProgress) onProgress({ phase: 'upload' });
    const licToken = await catalogToken();
    const publisher =
        (typeof window.ms365AuthGetAccountLabel === 'function' &&
            window.ms365AuthGetAccountLabel()) ||
        '';
    const mediaTotals = { imagesInlined: 0, imagesSkipped: 0, filesInlined: 0, embedsReplaced: 0 };
    let pagesWithHtml = 0;
    let pagesWithoutHtml = 0;
    Object.values(sectionsPayload).forEach((sec) => {
        (sec.pages || []).forEach((p) => {
            if (p && p.html) pagesWithHtml++;
            else pagesWithoutHtml++;
            const m = p && p.media;
            if (!m) return;
            mediaTotals.imagesInlined += m.imagesInlined || 0;
            mediaTotals.imagesSkipped += m.imagesSkipped || 0;
            mediaTotals.filesInlined += m.filesInlined || 0;
            mediaTotals.embedsReplaced += m.embedsReplaced || 0;
        });
    });

    const notebooksPayload = notebookMeta.map((n) =>
        Object.assign({}, n, {
            lastModifiedByName: n.lastModifiedByName || publisher
        })
    );

    // 1) Index + Bäume
    let result = await api.publishCatalogOnenoteSnapshot(licToken, {
        notebooks: notebooksPayload,
        trees,
        sections: {},
        publishedBy: publisher
    });

    // 2) Abschnitte einzeln (große Bild-Seiten sprengen sonst Request/Datei-Limits)
    const sectionEntries = Object.entries(sectionsPayload);
    let sectionUpload = 0;
    for (const [sid, sec] of sectionEntries) {
        sectionUpload++;
        if (onProgress) {
            onProgress({
                phase: 'upload',
                detail:
                    'Abschnitt ' +
                    sectionUpload +
                    '/' +
                    sectionEntries.length +
                    ' · ' +
                    (sec.displayName || sid)
            });
        }
        const htmlCount = (sec.pages || []).filter((p) => p && p.html).length;
        if (!htmlCount && (sec.pages || []).length) {
            // trotzdem Meta schreiben, aber im Log sichtbar
            if (onProgress) {
                onProgress({
                    phase: 'upload',
                    detail:
                        'Warnung: „' +
                        (sec.displayName || '') +
                        '“ ohne HTML – später erneut veröffentlichen'
                });
            }
        }
        result = await api.publishCatalogOnenoteSnapshot(licToken, {
            notebooks: notebooksPayload,
            trees: {},
            sections: { [sid]: sec },
            publishedBy: publisher
        });
    }

    return Object.assign({}, result, {
        media: mediaTotals,
        pagesWithHtml,
        pagesWithoutHtml,
        sectionCount: sectionEntries.length
    });
}
