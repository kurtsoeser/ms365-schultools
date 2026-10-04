/**
 * Nach Login: IT-Sicherungsbibliothek per Graph verknüpfen (Drive-ID),
 * wenn Site + Bibliotheksname lokal bekannt sind, aber driveId fehlt.
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
 * @returns {Promise<{ linked?: boolean, skipped?: string, meta?: object, error?: string }>}
 */
export async function tryAutoLinkItLibrary() {
    const current = normalizeItLibraryMeta(loadItMeta());
    if (isItLibraryConfigured(current)) {
        return { linked: false, skipped: 'already-configured' };
    }

    const hints = collectItLibraryLinkHints({
        itMeta: current,
        setup: readSetup(),
        formDraft: readItLibraryFormDraft()
    });
    if (!hints.hasMinimum) {
        return { linked: false, skipped: 'no-hints' };
    }

    const G = getG();
    let token;
    try {
        token = await G.getGraphToken(SCOPES_GRAPH);
    } catch (e) {
        return { linked: false, skipped: 'no-token', error: e && e.message ? e.message : String(e) };
    }

    let site;
    try {
        site = await G.resolveSiteFromWebUrl(token, hints.siteUrl);
    } catch (e) {
        return { linked: false, skipped: 'site-unresolved', error: e && e.message ? e.message : String(e) };
    }
    if (!site || !site.id) {
        return { linked: false, skipped: 'site-missing' };
    }

    let list = null;
    try {
        list = await findListByTitle(token, site.id, hints.listTitle);
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

    const meta = normalizeItLibraryMeta(
        Object.assign({}, current, {
            listTitle: hints.listTitle || IT_LIBRARY_TITLE,
            listId: String(listId),
            driveId: String(drive.id),
            webUrl: String(list.webUrl || drive.webUrl || ''),
            siteUrl: hints.siteUrl,
            itGroupId: hints.itGroupId || current.itGroupId || '',
            itGroupMail: hints.itGroupMail || current.itGroupMail || '',
            linkedAt: new Date().toISOString(),
            autoLinked: true
        })
    );
    saveItMeta(meta);
    return { linked: true, meta: meta };
}

export default { tryAutoLinkItLibrary };
