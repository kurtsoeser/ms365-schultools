/**
 * SharePoint-Listen: Zeilenanzahl per Microsoft Graph ($count / @odata.count).
 * Listennamen aus Stammdaten (intranetListTitles) und letztem Sync.
 */
import {
    DEFAULT_INTRANET_LIST_TITLES,
    resolveIntranetListTitle
} from '../../shared/intranet-list-title-logic.js';
import { readIntranetListLinks } from '../../shared/tenant-intranet-list-links.js';
import { resolveSchoolToolSiteUrl } from '../../shared/school-tool-sites.js';
import { LIST_TITLES as SA_LIST_TITLES } from '../schularbeiten-planer/schularbeiten-planer-schema.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
];

const FREISTELLUNG_SETUP_KEY = 'ms365-freistellung-setup-v1';

/**
 * @typedef {{
 *   countKey: string,
 *   defaultListTitle: string,
 *   siteKey: 'intranet'|'schularbeiten'|'projektwochen'|'freistellung'|'aktivitaeten',
 *   intranetKind?: string,
 *   resolveTitle?: (setup: object) => string
 * }} SpoListProbe
 */

/** @type {SpoListProbe[]} */
export const SPO_LIST_PROBES = [
    {
        countKey: 'spoListKlassen',
        defaultListTitle: DEFAULT_INTRANET_LIST_TITLES.klassen,
        siteKey: 'intranet',
        intranetKind: 'klassen'
    },
    {
        countKey: 'spoListFaecher',
        defaultListTitle: DEFAULT_INTRANET_LIST_TITLES.faecher,
        siteKey: 'intranet',
        intranetKind: 'faecher'
    },
    {
        countKey: 'spoListSchuelerinnen',
        defaultListTitle: DEFAULT_INTRANET_LIST_TITLES.schueler,
        siteKey: 'intranet',
        intranetKind: 'schueler'
    },
    {
        countKey: 'spoListLehrerPublic',
        defaultListTitle: DEFAULT_INTRANET_LIST_TITLES.lehrer,
        siteKey: 'intranet',
        intranetKind: 'lehrer'
    },
    {
        countKey: 'spoListSapSchularbeiten',
        defaultListTitle: SA_LIST_TITLES.schularbeiten,
        siteKey: 'schularbeiten'
    },
    {
        countKey: 'spoListPwAngebote',
        defaultListTitle: 'PW-Angebote',
        siteKey: 'projektwochen'
    },
    {
        countKey: 'spoListPwAktionen',
        defaultListTitle: 'PW-Aktionen',
        siteKey: 'projektwochen'
    },
    {
        countKey: 'spoListFreistellungen',
        defaultListTitle: 'Freistellungen',
        siteKey: 'freistellung',
        resolveTitle: () => readFreistellungListTitle()
    },
    {
        countKey: 'spoListAktivitaeten',
        defaultListTitle: 'Schulaktivitaeten',
        siteKey: 'aktivitaeten'
    }
];

/** Abwärtskompatibel: früher `listTitle` in Tests. */
SPO_LIST_PROBES.forEach(function (p) {
    Object.defineProperty(p, 'listTitle', {
        get: function () {
            return p.defaultListTitle;
        },
        enumerable: true
    });
});

function spoGraph() {
    const g = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!g) throw new Error('SharePoint-Graph (spo-graph-shared.js) nicht geladen.');
    return g;
}

function localStorageTrim(key) {
    try {
        return String(localStorage.getItem(key) || '').trim();
    } catch {
        return '';
    }
}

function readFreistellungListTitle() {
    try {
        const raw = localStorage.getItem(FREISTELLUNG_SETUP_KEY);
        const o = raw ? JSON.parse(raw) : {};
        const n = o && o.listName ? String(o.listName).trim() : '';
        return n || 'Freistellungen';
    } catch {
        return 'Freistellungen';
    }
}

function normListTitleKey(title) {
    return String(title || '')
        .trim()
        .toLowerCase()
        .normalize('NFC');
}

/**
 * @param {SpoListProbe} probe
 * @param {object} setup
 */
export function resolveSpoProbeListTitle(probe, setup) {
    if (!probe) return '';
    if (typeof probe.resolveTitle === 'function') {
        const t = String(probe.resolveTitle(setup || {}) || '').trim();
        if (t) return t;
    }
    const st = setup && typeof setup === 'object' ? setup : {};
    if (probe.intranetKind) {
        const titles =
            st.intranetListTitles && typeof st.intranetListTitles === 'object' ? st.intranetListTitles : {};
        return resolveIntranetListTitle(probe.intranetKind, titles[probe.intranetKind]);
    }
    return String(probe.defaultListTitle || '').trim();
}

/**
 * @param {object} setup
 */
export function resolveSiteUrls(setup) {
    const intranet = String(setup?.intranetSiteUrl || '').trim();
    const planerDefault = String(setup?.planerDefaultSiteUrl || '').trim() || intranet;
    return {
        intranet,
        schularbeiten: resolveSchoolToolSiteUrl('schularbeiten', setup) || planerDefault,
        projektwochen: resolveSchoolToolSiteUrl('projektwochen', setup) || planerDefault,
        freistellung: resolveSchoolToolSiteUrl('freistellung', setup) || planerDefault,
        aktivitaeten: resolveSchoolToolSiteUrl('aktivitaeten', setup) || planerDefault
    };
}

/** @param {string} countKey */
function cachedIntranetMetric(countKey) {
    const map = {
        spoListKlassen: 'klassen',
        spoListFaecher: 'faecher',
        spoListSchuelerinnen: 'schueler',
        spoListLehrerPublic: 'lehrer'
    };
    const kind = map[countKey];
    if (!kind) return null;
    const links = readIntranetListLinks();
    const row = links[kind];
    if (!row) return null;
    const count = row.count != null ? Number(row.count) : NaN;
    const title = row.title ? String(row.title).trim() : '';
    if (Number.isFinite(count) && count >= 0) {
        return {
            value: count,
            hint: title ? `Letzter Intranet-Sync · „${title}"` : 'Letzter Intranet-Sync'
        };
    }
    if (row.url) {
        return {
            value: '✓',
            hint: title ? `Liste verknüpft („${title}") – Zähler nur nach Anmeldung` : 'Liste verknüpft – Zähler nur nach Anmeldung'
        };
    }
    return null;
}

async function graphGet(token, path, version) {
    const G = spoGraph();
    const v = version === 'beta' ? 'beta' : 'v1.0';
    const url =
        String(path || '').indexOf('http') === 0
            ? path
            : 'https://graph.microsoft.com/' + v + (path.indexOf('/') === 0 ? path : '/' + path);
    const res = await fetch(url, {
        method: 'GET',
        headers: {
            Authorization: 'Bearer ' + token,
            ConsistencyLevel: 'eventual'
        }
    });
    const text = await res.text();
    let data = null;
    if (text) {
        try {
            data = JSON.parse(text);
        } catch {
            data = text;
        }
    }
    return { ok: res.ok, status: res.status, data, text };
}

/**
 * @param {string} token
 * @param {string} siteId
 * @param {string} listTitle
 */
export async function findGraphListOnSite(token, siteId, listTitle) {
    const G = spoGraph();
    const title = String(listTitle || '').trim();
    if (!title || !siteId) return null;
    const want = normListTitleKey(title);

    try {
        const filterPath =
            G.graphPathSite(siteId) +
            '/lists?$filter=' +
            encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
            '&$select=id,displayName,webUrl';
        const filtered = await G.graphJson('GET', filterPath, token, undefined, 'v1.0');
        const hit = ((filtered && filtered.value) || [])[0];
        if (hit && hit.id) return hit;
    } catch {
        /* Filter schlägt bei Sonderzeichen/Indexing manchmal fehl */
    }

    let path = G.graphPathSite(siteId) + '/lists?$select=id,displayName,webUrl&$top=200';
    while (path) {
        const data = await G.graphJson(
            'GET',
            path.indexOf('http') === 0 ? path : path,
            token,
            undefined,
            'v1.0'
        );
        const batch = (data && data.value) || [];
        for (let i = 0; i < batch.length; i++) {
            const row = batch[i];
            if (normListTitleKey(row.displayName) === want) return row;
        }
        path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
    }
    return null;
}

function parseODataCount(payload) {
    if (!payload || typeof payload !== 'object') return NaN;
    const c = payload['@odata.count'];
    return typeof c === 'number' ? c : parseInt(String(c), 10);
}

function graphErrorText(raw, fallback) {
    if (raw && raw.data && raw.data.error && raw.data.error.message) {
        return String(raw.data.error.message);
    }
    if (raw && raw.text) return String(raw.text).slice(0, 200);
    if (raw && raw.status) return String(raw.status);
    return fallback || 'Zähler fehlgeschlagen';
}

/**
 * Zeilen zählen: $count → @odata.count → Seitenweise ids (wenn Advanced Queries/ $count blockiert).
 * @param {string} token
 * @param {string} siteId
 * @param {string} listId
 * @returns {Promise<{ ok: boolean, count: number, via?: string, truncated?: boolean, error?: string }>}
 */
export async function fetchListItemCount(token, siteId, listId) {
    const G = spoGraph();
    const id = String(listId || '').trim();
    if (!id) return { ok: false, count: -1, error: 'listId fehlt' };

    const base = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(id) + '/items';
    const errors = [];

    // 1) Dedizierter $count-Endpunkt (braucht ConsistencyLevel: eventual)
    const rawCount = await graphGet(token, base + '/$count', 'v1.0');
    if (rawCount.ok) {
        const n = parseInt(String(rawCount.text || '').trim(), 10);
        if (Number.isFinite(n) && n >= 0) return { ok: true, count: n, via: '$count' };
    } else {
        errors.push(graphErrorText(rawCount, '$count'));
    }

    // 2) @odata.count – $top=0 wird von manchen Tenants abgelehnt → $top=1
    const odataPath = base + '?$select=id&$top=1&$count=true';
    const odata = await graphGet(token, odataPath, 'v1.0');
    if (odata.ok) {
        const n = parseODataCount(odata.data);
        if (Number.isFinite(n) && n >= 0) return { ok: true, count: n, via: '@odata.count' };
    } else {
        errors.push(graphErrorText(odata, '@odata.count'));
    }

    // 3) Fallback: sichtbare Items seitenweise zählen (ohne Advanced Queries)
    const PAGE = 200;
    const MAX = 5000;
    let total = 0;
    let path = base + '?$select=id&$top=' + PAGE;
    let pageOk = false;
    while (path && total < MAX) {
        const page = await graphGet(token, path.indexOf('http') === 0 ? path : path, 'v1.0');
        if (!page.ok) {
            errors.push(graphErrorText(page, 'paging'));
            break;
        }
        pageOk = true;
        const batch = (page.data && page.data.value) || [];
        total += batch.length;
        if (batch.length < PAGE) {
            return { ok: true, count: total, via: 'page' };
        }
        path = page.data && page.data['@odata.nextLink'] ? page.data['@odata.nextLink'] : '';
        if (!path) {
            return { ok: true, count: total, via: 'page' };
        }
    }
    if (pageOk && total >= MAX) {
        return { ok: true, count: MAX, via: 'page', truncated: true };
    }
    if (pageOk) {
        return { ok: true, count: total, via: 'page' };
    }

    return { ok: false, count: -1, error: errors.filter(Boolean)[0] || 'Zähler fehlgeschlagen' };
}

/**
 * @param {object} setup app-data setup
 * @returns {Promise<{ metrics: Record<string, { value: number|string, hint?: string }>, errors: string[] }>}
 */
export async function fetchSharePointListMetrics(setup) {
    const sites = resolveSiteUrls(setup || {});
    const errors = [];
    /** @type {Record<string, { value: number|string, hint?: string }>} */
    const metrics = {};

    if (!sites.intranet && !sites.schularbeiten) {
        return {
            metrics: {},
            errors: ['Keine Intranet-/Planer-Site-URL – in den Stammdaten unter Einrichtung setzen.']
        };
    }

    let token;
    try {
        token = await spoGraph().getGraphToken(SCOPES);
    } catch (e) {
        const cachedOnly = {};
        SPO_LIST_PROBES.forEach(function (probe) {
            const c = cachedIntranetMetric(probe.countKey);
            if (c) cachedOnly[probe.countKey] = c;
        });
        return {
            metrics: cachedOnly,
            errors: [e && e.message ? e.message : 'Anmeldung für SharePoint-Zähler fehlgeschlagen.']
        };
    }

    /** @type {Map<string, string>} */
    const siteIds = new Map();

    async function siteIdFor(key) {
        const url = sites[key];
        if (!url) return '';
        if (siteIds.has(key)) return siteIds.get(key);
        try {
            const site = await spoGraph().resolveSiteFromWebUrl(token, url);
            const id = site && site.id ? String(site.id) : '';
            siteIds.set(key, id);
            return id;
        } catch (e) {
            errors.push(`${key}: ${e && e.message ? e.message : e}`);
            siteIds.set(key, '');
            return '';
        }
    }

    for (const probe of SPO_LIST_PROBES) {
        const listTitle = resolveSpoProbeListTitle(probe, setup);
        const siteUrl = sites[probe.siteKey];
        if (!siteUrl) {
            metrics[probe.countKey] = { value: '—', hint: 'Site-URL fehlt (Planer oder Intranet)' };
            continue;
        }
        try {
            const siteId = await siteIdFor(probe.siteKey);
            if (!siteId) {
                metrics[probe.countKey] = { value: '—', hint: 'Site nicht erreichbar' };
                continue;
            }
            const list = await findGraphListOnSite(token, siteId, listTitle);
            if (!list || !list.id) {
                const cached = cachedIntranetMetric(probe.countKey);
                if (cached) {
                    metrics[probe.countKey] = {
                        value: cached.value,
                        hint: `Graph: Liste „${listTitle}" nicht gefunden · ${cached.hint || ''}`
                    };
                } else {
                    metrics[probe.countKey] = {
                        value: '—',
                        hint: `Liste „${listTitle}" auf Site nicht gefunden (Name in den Stammdaten prüfen)`
                    };
                    errors.push(`${listTitle}: Liste auf ${probe.siteKey} nicht gefunden`);
                }
                continue;
            }
            const cnt = await fetchListItemCount(token, siteId, list.id);
            if (!cnt.ok || cnt.count < 0) {
                // Liste ist da – kein „?“ (wirkt wie Fehler). Cache oder ✓ nutzen.
                const cached = cachedIntranetMetric(probe.countKey);
                if (cached && (typeof cached.value === 'number' || cached.value === '✓')) {
                    metrics[probe.countKey] = {
                        value: cached.value,
                        hint:
                            (typeof cached.value === 'number'
                                ? 'Zeilenzahl aus letztem Sync (Graph-Zähler gesperrt)'
                                : 'Liste gefunden, Zeilenzahl nicht lesbar') +
                            (cached.hint ? ' · ' + cached.hint : '')
                    };
                } else {
                    metrics[probe.countKey] = {
                        value: '✓',
                        hint:
                            'Liste „' +
                            (list.displayName || listTitle) +
                            '“ vorhanden – Zeilenzahl nicht lesbar' +
                            (cnt.error ? ' (' + String(cnt.error).slice(0, 100) + ')' : '')
                    };
                }
                continue;
            }
            const label = list.displayName || listTitle;
            metrics[probe.countKey] = {
                value: cnt.count,
                hint: cnt.truncated
                    ? `SharePoint · „${label}" (≥${cnt.count}, Zähler gekappt)`
                    : `SharePoint · „${label}"` + (cnt.via && cnt.via !== '$count' ? ` · via ${cnt.via}` : '')
            };
        } catch (e) {
            metrics[probe.countKey] = { value: '—', hint: 'Fehler beim Lesen' };
            errors.push(`${listTitle}: ${e && e.message ? e.message : e}`);
        }
    }

    return { metrics, errors };
}
