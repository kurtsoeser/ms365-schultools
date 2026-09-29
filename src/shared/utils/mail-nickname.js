/**
 * Einheitliche Mail-Nickname-Normalisierung für Unified Groups.
 * Umlaute → ae/oe/ue/ss, sonst nur `[a-z0-9-]`, max. 64 Zeichen.
 */

/**
 * @param {unknown} raw
 * @param {{ maxLen?: number }} [opts]
 * @returns {string}
 */
export function normalizeMailNickname(raw, opts) {
    const maxLen = opts && typeof opts.maxLen === 'number' && opts.maxLen > 0 ? opts.maxLen : 64;
    let s = String(raw ?? '')
        .trim()
        .toLowerCase()
        .replace(/ä/g, 'ae')
        .replace(/ö/g, 'oe')
        .replace(/ü/g, 'ue')
        .replace(/ß/g, 'ss')
        .replace(/\s+/g, '-')
        .replace(/[^a-z0-9-]/g, '')
        .replace(/-+/g, '-')
        .replace(/^-|-$/g, '');
    if (s.length > maxLen) s = s.slice(0, maxLen).replace(/-$/g, '');
    return s;
}
