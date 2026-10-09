/**
 * SharePoint-Listen: Zeilenanzahl per Microsoft Graph ($count / @odata.count).
 * Listennamen aus Stammdaten (intranetListTitles) und letztem Sync.
 */
import {
    DEFAULT_INTRANET_LIST_TITLES,
    resolveIntranetListTitle
} from '../../shared/intranet-list-title-logic.js';
import { readIntranetListLinks } from '../../shared/tenant-intranet-list-links.js';
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
    return {
        intranet,
        schularbeiten: localStorageTrim('ms365-sa-site-url') || intranet,
        projektwochen: localStorageTrim('ms365-pw-site-url') || intranet,
        freistellung: localStorageTrim('ms365-freistellung-planer-site-v1') || intranet,
        aktivitaeten: localStorageTrim('ms365-akt-planer-site-v1') || intranet
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

/**
 * @param {string} token
 * @param {string} siteId
 * @param {string} listId
 */
export async function fetchListItemCount(token, siteId, listId) {
    const G = spoGraph();
    const id = String(listId || '').trim();
    if (!id) return { ok: false, count: -1, error: 'listId fehlt' };

    const countPath = G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(id) + '/items/$count';
    const raw = await graphGet(token, countPath, 'v1.0');
    if (raw.ok) {
        const n = parseInt(String(raw.text || '').trim(), 10);
        if (Number.isFinite(n) && n >= 0) return { ok: true, count: n, via: '$count' };
    }

    const odataPath =
        G.graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(id) +
        '/items?$select=id&$top=0&$count=true';
    const odata = await graphGet(token, odataPath, 'v1.0');
    if (odata.ok && odata.data && typeof odata.data === 'object') {
        const c = odata.data['@odata.count'];
        const n = typeof c === 'number' ? c : parseInt(String(c), 10);
        if (Number.isFinite(n) && n >= 0) return { ok: true, count: n, via: '@odata.count' };
    }

    const err =
        (odata.data && odata.data.error && odata.data.error.message) ||
        (raw.data && raw.data.error && raw.data.error.message) ||
        raw.text ||
        String(raw.status || odata.status || '');
    return { ok: false, count: -1, error: err };
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
                const cached = cachedIntranetMetric(probe.countKey);
                if (cached && typeof cached.value === 'number') {
                    metrics[probe.countKey] = {
                        value: cached.value,
                        hint: `Zähler per Graph nicht lesbar · ${cached.hint || ''}`
                    };
                } else {
                    metrics[probe.countKey] = {
                        value: '?',
                        hint:
                            'Zähler nicht lesbar (Berechtigung Sites.Read.All oder Listenrechte). ' +
                            (cnt.error ? String(cnt.error).slice(0, 120) : '')
                    };
                    errors.push(`${listTitle}: ${cnt.error || 'Zähler fehlgeschlagen'}`);
                }
                continue;
            }
            metrics[probe.countKey] = {
                value: cnt.count,
                hint: `SharePoint · „${list.displayName || listTitle}"`
            };
        } catch (e) {
            metrics[probe.countKey] = { value: '—', hint: 'Fehler beim Lesen' };
            errors.push(`${listTitle}: ${e && e.message ? e.message : e}`);
        }
    }

    return { metrics, errors };
}
