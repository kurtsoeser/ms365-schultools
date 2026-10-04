/**
 * SharePoint-Listen des Schularbeiten-Planers auflösen (SAP-Namen + Legacy).
 */
import { LIST_TITLES, titlesForListKey, LIST_DESCRIPTIONS } from './schularbeiten-planer-schema.js';

/**
 * @param {(tok: string, siteId: string, title: string) => Promise<object|null>} findListByTitle
 */
export async function resolvePlanerList(tok, siteId, listKey, findListByTitle) {
    const titles = titlesForListKey(listKey);
    const canonical = LIST_TITLES[listKey];
    for (let i = 0; i < titles.length; i++) {
        const title = titles[i];
        const list = await findListByTitle(tok, siteId, title);
        if (list && list.id) {
            const displayName = String(list.displayName || title).trim();
            return {
                list,
                listKey,
                displayName,
                canonicalTitle: canonical,
                isLegacyTitle: displayName !== canonical
            };
        }
    }
    return null;
}

/**
 * @param {object} graphApi ms365SpoGraph
 * @param {string} token
 * @param {string} siteId
 * @param {string} listId
 * @param {string} newTitle
 */
export async function renameListDisplayName(graphApi, token, siteId, listId, newTitle) {
    const title = String(newTitle || '').trim();
    if (!title) throw new Error('Neuer Listenname fehlt.');
    await graphApi.graphJson(
        'PATCH',
        graphApi.graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId),
        token,
        { displayName: title },
        'v1.0'
    );
}

/**
 * @param {object} graphApi
 * @param {string} token
 * @param {string} siteId
 * @param {string} listKey
 * @param {(tok: string, siteId: string, title: string) => Promise<object|null>} findListByTitle
 * @param {(msg: string) => void} [write]
 */
export async function ensureCanonicalListTitle(graphApi, token, siteId, listKey, findListByTitle, write) {
    const resolved = await resolvePlanerList(token, siteId, listKey, findListByTitle);
    if (!resolved || !resolved.isLegacyTitle) return resolved;
    const log = typeof write === 'function' ? write : () => {};
    log(
        'Benenne Legacy-Liste „' +
            resolved.displayName +
            '" → „' +
            resolved.canonicalTitle +
            '" …'
    );
    await renameListDisplayName(
        graphApi,
        token,
        siteId,
        String(resolved.list.id),
        resolved.canonicalTitle
    );
    const list = await findListByTitle(token, siteId, resolved.canonicalTitle);
    return {
        list: list || { ...resolved.list, displayName: resolved.canonicalTitle },
        listKey,
        displayName: resolved.canonicalTitle,
        canonicalTitle: resolved.canonicalTitle,
        isLegacyTitle: false
    };
}

export { LIST_DESCRIPTIONS };
