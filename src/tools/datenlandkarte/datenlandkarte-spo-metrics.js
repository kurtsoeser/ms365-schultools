/**
 * SharePoint-Listen: Zeilenanzahl per Microsoft Graph ($count).
 */

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
];

/**
 * @typedef {{ countKey: string, listTitle: string, siteKey: 'intranet'|'schularbeiten'|'projektwochen'|'freistellung'|'aktivitaeten' }} SpoListProbe
 */

/** @type {SpoListProbe[]} */
export const SPO_LIST_PROBES = [
    { countKey: 'spoListKlassen', listTitle: 'Klassen', siteKey: 'intranet' },
    { countKey: 'spoListFaecher', listTitle: 'Fächer', siteKey: 'intranet' },
    { countKey: 'spoListSchuelerinnen', listTitle: 'Schülerinnen', siteKey: 'intranet' },
    { countKey: 'spoListLehrerPublic', listTitle: 'Lehrerinnen', siteKey: 'intranet' },
    { countKey: 'spoListSapSchularbeiten', listTitle: 'SAP-Schularbeiten', siteKey: 'schularbeiten' },
    { countKey: 'spoListPwAngebote', listTitle: 'PW-Angebote', siteKey: 'projektwochen' },
    { countKey: 'spoListPwAktionen', listTitle: 'PW-Aktionen', siteKey: 'projektwochen' },
    { countKey: 'spoListFreistellungen', listTitle: 'Freistellungen', siteKey: 'freistellung' },
    { countKey: 'spoListAktivitaeten', listTitle: 'Schulaktivitaeten', siteKey: 'aktivitaeten' }
];

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

async function findListByTitle(token, siteId, listTitle) {
    const G = spoGraph();
    const title = String(listTitle || '').trim();
    const path =
        G.graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName';
    const data = await G.graphJson('GET', path, token, undefined, 'v1.0');
    return ((data && data.value) || [])[0] || null;
}

async function fetchListItemCount(token, siteId, listId) {
    const G = spoGraph();
    const path =
        G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/items/$count';
    const url = 'https://graph.microsoft.com/v1.0' + path;
    const res = await fetch(url, {
        method: 'GET',
        headers: {
            Authorization: 'Bearer ' + token,
            ConsistencyLevel: 'eventual'
        }
    });
    const text = await res.text();
    if (!res.ok) return { ok: false, count: -1, error: text || String(res.status) };
    const n = parseInt(String(text).trim(), 10);
    return { ok: true, count: Number.isFinite(n) ? n : -1 };
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
            errors: ['Keine Intranet-/Planer-Site-URL – Einrichtung oder Planer-URL setzen.']
        };
    }

    let token;
    try {
        token = await spoGraph().getGraphToken(SCOPES);
    } catch (e) {
        return {
            metrics: {},
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
        const siteUrl = sites[probe.siteKey];
        if (!siteUrl) {
            metrics[probe.countKey] = { value: '—', hint: 'Site-URL fehlt' };
            continue;
        }
        try {
            const siteId = await siteIdFor(probe.siteKey);
            if (!siteId) {
                metrics[probe.countKey] = { value: '—', hint: 'Site nicht erreichbar' };
                continue;
            }
            const list = await findListByTitle(token, siteId, probe.listTitle);
            if (!list || !list.id) {
                metrics[probe.countKey] = {
                    value: '—',
                    hint: `Liste „${probe.listTitle}" nicht auf Site`
                };
                continue;
            }
            const cnt = await fetchListItemCount(token, siteId, list.id);
            if (!cnt.ok || cnt.count < 0) {
                metrics[probe.countKey] = { value: '?', hint: 'Zähler nicht lesbar' };
                continue;
            }
            metrics[probe.countKey] = {
                value: cnt.count,
                hint: `SharePoint · ${probe.listTitle}`
            };
        } catch (e) {
            metrics[probe.countKey] = { value: '—', hint: 'Fehler beim Lesen' };
            errors.push(`${probe.listTitle}: ${e && e.message ? e.message : e}`);
        }
    }

    return { metrics, errors };
}
