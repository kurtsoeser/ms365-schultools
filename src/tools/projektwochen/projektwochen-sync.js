/**
 * SPO Live Sync Meta für Projektwochen.
 * CRUD bleibt in projektwochen-graph.js; Bookings bleibt getrennt.
 * localStorage/Demo = Cache, SharePoint = System of Record.
 */

/**
 * @returns {{
 *   ok: boolean,
 *   source: 'live'|'local'|'demo'|'none',
 *   connected: boolean,
 *   error: string,
 *   siteUrl: string,
 *   lists: string,
 *   lastPullAt: string,
 *   lastPushAt: string
 * }}
 */
export function emptyPwSpoSyncMeta() {
    return {
        ok: false,
        source: 'none',
        connected: false,
        error: '',
        siteUrl: '',
        lists: '',
        lastPullAt: '',
        lastPushAt: ''
    };
}

/**
 * Lokale Demo-/Offline-IDs dürfen nicht per PATCH nach SPO.
 * @param {{ itemId?: string }|string|null|undefined} itemOrId
 */
export function isPwLocalOnlyItem(itemOrId) {
    const id =
        typeof itemOrId === 'string'
            ? itemOrId
            : String((itemOrId && itemOrId.itemId) || '');
    const s = String(id || '').trim();
    if (!s) return true;
    return s.indexOf('local-') === 0 || s.indexOf('demo-') === 0;
}

/**
 * Demo local-first oder fehlender Context → kein SPO-Write.
 * Hinweis: `local-*`-IDs bei bestehendem Context werden vom Caller per Create promotet
 * (nicht hier übersprungen).
 * @param {{ localDemoOnly?: boolean, ctx?: object|null }} stateLike
 */
export function shouldSkipPwSpoPush(stateLike) {
    return !stateLike || !!stateLike.localDemoOnly || !stateLike.ctx;
}

/**
 * @param {object|null|undefined} prev
 * @param {Partial<ReturnType<typeof emptyPwSpoSyncMeta>>} partial
 */
export function mergePwSpoSyncMeta(prev, partial) {
    return { ...emptyPwSpoSyncMeta(), ...(prev || {}), ...(partial || {}) };
}

/**
 * Kurztext für System-UI.
 * @param {ReturnType<typeof emptyPwSpoSyncMeta>} meta
 */
export function formatPwSpoSyncLine(meta) {
    const m = meta || emptyPwSpoSyncMeta();
    const parts = ['Quelle: ' + (m.source || 'none')];
    if (m.lists) parts.push('Listen ' + m.lists);
    if (m.connected) parts.push('verbunden');
    else parts.push('nicht verbunden');
    if (m.lastPullAt) parts.push('Pull ' + formatPwSyncTime(m.lastPullAt));
    if (m.lastPushAt) parts.push('Push ' + formatPwSyncTime(m.lastPushAt));
    if (m.error) parts.push(m.error);
    return parts.join(' · ');
}

/**
 * @param {string} iso
 */
export function formatPwSyncTime(iso) {
    const s = String(iso || '').trim();
    if (!s) return '';
    try {
        const d = new Date(s);
        if (Number.isNaN(d.getTime())) return s;
        return d.toLocaleString('de-AT', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch {
        return s;
    }
}
