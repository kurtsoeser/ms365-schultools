/**
 * Nach Login: IT-Sicherungsbibliothek per Graph verknüpfen (Drive-ID),
 * wenn Site + Bibliotheksname lokal bekannt sind, aber driveId fehlt –
 * oder per Tenant-Discovery / expliziter Site-URL.
 */
import {
    IT_LIBRARY_TITLE,
    isItLibraryConfigured,
    collectItLibraryLinkHints,
    normalizeItLibraryMeta
} from './stammdaten-sharepoint-sync-logic.js';
import {
    loadItMeta,
    saveItMeta,
    readItLibraryFormDraft
} from './stammdaten-sharepoint-sync-api.js';
import { discoverItLibraryInTenant } from './stammdaten-sharepoint-it-library-discover.js';

const SCOPES_GRAPH = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/Group.Read.All'
];

function getG() {
    const G = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!G) throw new Error('SharePoint-Graph-Helfer nicht geladen (spo-graph-shared.js).');
    return G;
}

function readSetup() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
            return window.ms365AppDataV2.getSetup() || {};
        }
    } catch {
        /* ignore */
    }
    return {};
}

async function findListByTitle(token, siteId, listTitle) {
    const G = getG();
    const title = String(listTitle || '').trim();
    const path =
        G.graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await G.graphJson('GET', path, token, undefined, 'v1.0');
    const list = (data && data.value) || [];
    return list[0] || null;
}

async function getListDrive(token, siteId, listId) {
    const G = getG();
    return G.graphJson(
        'GET',
        G.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/drive',
        token,
        undefined,
        'v1.0'
    );
}

/**
 * @param {object} current
 * @param {{ siteUrl: string, listTitle: string, listId: string, driveId: string, webUrl?: string, itGroupId?: string, itGroupMail?: string, autoLinked?: boolean }} hit
 */
function persistLinkedMeta(current, hit) {
    const meta = normalizeItLibraryMeta(
        Object.assign({}, current, {
            listTitle: hit.listTitle || IT_LIBRARY_TITLE,
            listId: String(hit.listId),
            driveId: String(hit.driveId),
            webUrl: String(hit.webUrl || ''),
            siteUrl: String(hit.siteUrl || '').replace(/\/$/, ''),
            itGroupId: hit.itGroupId || current.itGroupId || '',
            itGroupMail: hit.itGroupMail || current.itGroupMail || '',
            linkedAt: new Date().toISOString(),
            autoLinked: hit.autoLinked !== false
        })
    );
    saveItMeta(meta);
    return meta;
}

/**
 * @param {string} token
 * @param {string} siteUrl
 * @param {string} listTitle
 * @param {object} current
 * @param {{ itGroupId?: string, itGroupMail?: string }} extras
 */
async function linkFromKnownSite(token, siteUrl, listTitle, current, extras) {
    const G = getG();
    const url = String(siteUrl || '').trim().replace(/\/$/, '');
    const title = String(listTitle || IT_LIBRARY_TITLE).trim() || IT_LIBRARY_TITLE;
    const extra = extras || {};

    let site;
    try {
        site = await G.resolveSiteFromWebUrl(token, url);
    } catch (e) {
        return { linked: false, skipped: 'site-unresolved', error: e && e.message ? e.message : String(e) };
    }
    if (!site || !site.id) {
        return { linked: false, skipped: 'site-missing' };
    }

    let list = null;
    try {
        list = await findListByTitle(token, site.id, title);
    } catch (e) {
        return { linked: false, skipped: 'list-search-failed', error: e && e.message ? e.message : String(e) };
    }
    if (!list) {
        return { linked: false, skipped: 'library-not-found' };
    }

    const listId = list.id || list.Id;
    if (!listId) {
        return { linked: false, skipped: 'list-id-missing' };
    }

    let drive;
    try {
        drive = await getListDrive(token, site.id, listId);
    } catch (e) {
        return { linked: false, skipped: 'drive-unreadable', error: e && e.message ? e.message : String(e) };
    }
    if (!drive || !drive.id) {
        return { linked: false, skipped: 'drive-missing' };
    }

    const meta = persistLinkedMeta(current, {
        siteUrl: url,
        listTitle: title,
        listId: String(listId),
        driveId: String(drive.id),
        webUrl: String(list.webUrl || drive.webUrl || ''),
        itGroupId: extra.itGroupId || '',
        itGroupMail: extra.itGroupMail || '',
        autoLinked: true
    });
    return { linked: true, meta: meta, via: 'site-url' };
}

/**
 * @param {{ siteUrl?: string, listTitle?: string, allowDiscover?: boolean }} [opts]
 * @returns {Promise<{ linked?: boolean, skipped?: string, meta?: object, error?: string, via?: string }>}
 */
export async function tryAutoLinkItLibrary(opts) {
    const options = opts || {};
    const allowDiscover = options.allowDiscover !== false;
    const current = normalizeItLibraryMeta(loadItMeta());
    if (isItLibraryConfigured(current)) {
        return { linked: false, skipped: 'already-configured' };
    }

    const hints = collectItLibraryLinkHints({
        itMeta: current,
        setup: readSetup(),
        formDraft: readItLibraryFormDraft()
    });

    const overrideSite = String(options.siteUrl || '').trim().replace(/\/$/, '');
    const listTitle =
        String(options.listTitle || hints.listTitle || IT_LIBRARY_TITLE).trim() || IT_LIBRARY_TITLE;
    const siteUrls = overrideSite
        ? [overrideSite]
        : Array.isArray(hints.siteUrls) && hints.siteUrls.length
          ? hints.siteUrls.slice()
          : hints.siteUrl
            ? [hints.siteUrl]
            : [];
    const siteUrl = siteUrls[0] || '';

    const G = getG();
    let token;
    try {
        token = await G.getGraphToken(SCOPES_GRAPH);
    } catch (e) {
        return { linked: false, skipped: 'no-token', error: e && e.message ? e.message : String(e) };
    }

    let lastSiteResult = null;
    for (let i = 0; i < siteUrls.length; i++) {
        const fromSite = await linkFromKnownSite(token, siteUrls[i], listTitle, current, {
            itGroupId: hints.itGroupId,
            itGroupMail: hints.itGroupMail
        });
        if (fromSite.linked) return fromSite;
        lastSiteResult = fromSite;
        if (
            fromSite.skipped !== 'library-not-found' &&
            fromSite.skipped !== 'site-unresolved' &&
            fromSite.skipped !== 'site-missing'
        ) {
            if (!allowDiscover) return fromSite;
            /* harte Fehler (Token/Drive) nicht mit nächster URL weiterprobieren */
            if (
                fromSite.skipped === 'no-token' ||
                fromSite.skipped === 'drive-unreadable' ||
                fromSite.skipped === 'drive-missing' ||
                fromSite.skipped === 'list-search-failed'
            ) {
                return fromSite;
            }
        }
    }

    if (!allowDiscover) {
        return lastSiteResult || { linked: false, skipped: siteUrl ? 'library-not-found' : 'no-hints' };
    }

    let discovered = null;
    try {
        discovered = await discoverItLibraryInTenant({
            token: token,
            listTitle: listTitle,
            preferSiteUrl: siteUrl || ''
        });
    } catch (e) {
        return {
            linked: false,
            skipped: (lastSiteResult && lastSiteResult.skipped) || (siteUrl ? 'library-not-found' : 'no-hints'),
            error: e && e.message ? e.message : String(e)
        };
    }

    if (!discovered) {
        return (
            lastSiteResult || {
                linked: false,
                skipped: siteUrl ? 'library-not-found' : 'no-hints'
            }
        );
    }

    const meta = persistLinkedMeta(current, {
        siteUrl: discovered.siteUrl,
        listTitle: discovered.listTitle || listTitle,
        listId: discovered.listId,
        driveId: discovered.driveId,
        webUrl: discovered.webUrl,
        itGroupId: hints.itGroupId,
        itGroupMail: hints.itGroupMail,
        autoLinked: true
    });
    return { linked: true, meta: meta, via: discovered.via || 'discover' };
}

export default { tryAutoLinkItLibrary };
