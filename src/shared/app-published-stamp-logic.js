/**
 * Anzeige des letzten Veröffentlichungszeitpunkts (Build/Deploy).
 */

export const STAMP_ELEMENT_ID = 'ms365AppPublishedStamp';

/**
 * @param {unknown} payload
 * @returns {string}
 */
export function parsePublishedAt(payload) {
    if (!payload || typeof payload !== 'object') return '';
    return String(payload.publishedAt || '').trim();
}

/**
 * @param {string} iso
 * @param {string} [timeZone]
 * @returns {string}
 */
export function formatPublishedStamp(iso, timeZone = 'Europe/Vienna') {
    const value = String(iso || '').trim();
    if (!value) return '';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return '';
    const formatted = new Intl.DateTimeFormat('de-AT', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        timeZone
    }).format(date);
    return `Stand: ${formatted}`;
}

/**
 * @param {string} label
 * @returns {HTMLParagraphElement | null}
 */
export function createPublishedStampElement(label, doc = document) {
    const text = String(label || '').trim();
    if (!text || !doc || !doc.createElement) return null;
    const el = doc.createElement('p');
    el.id = STAMP_ELEMENT_ID;
    el.className = 'app-published-stamp';
    el.textContent = text;
    el.title = 'Zeitpunkt der letzten Veröffentlichung dieser App';
    return el;
}
