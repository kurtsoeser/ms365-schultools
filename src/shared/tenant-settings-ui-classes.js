/**
 * Klassen-Tab Helfer für tenant-settings-ui (Analyse 02 Phase C).
 * Pure Zeilen-Konvertierung – Render bleibt in der UI-Closure.
 */
import { normStr, normCode } from './utils/strings.js';

/**
 * @param {Array<{ code?: string, year?: string, name?: string, headName?: string, headEmail?: string }>} rows
 * @returns {string}
 */
export function classesToLines(rows) {
    return (rows || [])
        .map((x) => {
            const y = normStr(x.year || '');
            const year = /^\d{4}$/.test(y) ? y : '';
            return `${normCode(x.code)};${year};${normStr(x.name || '')};${normStr(x.headName || '')};${normStr(x.headEmail || '').toLowerCase()}`.trim();
        })
        .filter(Boolean)
        .join('\n');
}
