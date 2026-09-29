/**
 * Kanonisches Schuljahr (AT/DE): Start am 1. September.
 * Jan–Aug → noch das vorherige Schuljahr; Sep–Dez → neues Schuljahr.
 *
 * Beispiel: 15.03.2026 → `2025/26`, 01.09.2026 → `2026/27`.
 */

/**
 * @param {Date|number|string} [now]
 * @returns {Date}
 */
function asDate(now) {
    if (now instanceof Date && !isNaN(now.getTime())) return now;
    if (typeof now === 'number' && isFinite(now)) {
        const d = new Date(now);
        if (!isNaN(d.getTime())) return d;
    }
    if (typeof now === 'string' && now.trim()) {
        const d = new Date(now);
        if (!isNaN(d.getTime())) return d;
    }
    return new Date();
}

/**
 * Anfangsjahr des laufenden Schuljahrs (Sep–Aug).
 * @param {Date|number|string} [now]
 * @returns {number}
 */
export function schoolYearStartYear(now) {
    const d = asDate(now);
    const y = d.getFullYear();
    // getMonth(): 0=Jan … 7=Aug → noch Vorjahr; ab Sep (8) neues Jahr
    return d.getMonth() < 8 ? y - 1 : y;
}

/**
 * Label `"YYYY/YY"` für das laufende Schuljahr.
 * @param {Date|number|string} [now]
 * @returns {string}
 */
export function currentSchoolYearLabel(now) {
    const y = schoolYearStartYear(now);
    return String(y) + '/' + String(y + 1).slice(2);
}

/**
 * Extrahiert das Anfangsjahr aus `"2025/26"` / `"2025/2026"`.
 * @param {unknown} label
 * @returns {number}
 */
export function parseSchoolYearStartYear(label) {
    const m = String(label || '')
        .trim()
        .match(/^(\d{4})\s*\/\s*(\d{2}|\d{4})/);
    return m ? parseInt(m[1], 10) : NaN;
}

/**
 * Nächstes Schuljahr zu `cur`. Ungültig → aktuelles Label.
 * @param {unknown} cur
 * @param {Date|number|string} [now]
 * @returns {string}
 */
export function nextSchoolYearLabel(cur, now) {
    const y = parseSchoolYearStartYear(cur);
    if (!isFinite(y)) return currentSchoolYearLabel(now);
    return String(y + 1) + '/' + String(y + 2).slice(2);
}

/**
 * @param {unknown} label
 * @returns {boolean}
 */
export function isSchoolYearLabel(label) {
    return isFinite(parseSchoolYearStartYear(label));
}

/** Für Legacy-IIFE-Module (app-data-v2, tenant-settings-ui, demo-mode). */
if (typeof window !== 'undefined') {
    window.ms365SchoolYear = {
        schoolYearStartYear,
        currentSchoolYearLabel,
        parseSchoolYearStartYear,
        nextSchoolYearLabel,
        isSchoolYearLabel
    };
}
