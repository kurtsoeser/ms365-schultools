/**
 * Klassen-Tab Helfer für tenant-settings-ui (Analyse 02 Phase C).
 * Pure Zeilen-Konvertierung – Render bleibt in der UI-Closure.
 */
import { normStr, normCode } from './utils/strings.js';

export const CLASS_LIST_XLSX_FILENAME = 'Klassenliste-Vorlage.xlsx';
export const CLASS_LIST_SHEET_NAME = 'Klassen';

/** Spalten der Excel-Vorlage – kompatibel zum CSV/XLSX-Import. */
export const CLASS_LIST_TEMPLATE_HEADERS = ['Kürzel', 'Abschlussjahr', 'Anzeigename', 'KV-Name', 'KV-E-Mail'];

/** @returns {string[][]} */
export function classListTemplateAoa() {
    return [
        CLASS_LIST_TEMPLATE_HEADERS.slice(),
        ['1AK', '2030', '1A-Klasse', 'Max Mustermann', 'max.mustermann@schule.de'],
        ['2BK', '2029', '2B-Klasse', 'Anna Beispiel', 'anna.beispiel@schule.de']
    ];
}

/**
 * @param {Array<{ code?: string, year?: string, name?: string, headName?: string, headEmail?: string }>} rows
 * @returns {string[][]}
 */
export function classesRowsToAoa(rows) {
    const aoa = [CLASS_LIST_TEMPLATE_HEADERS.slice()];
    (rows || []).forEach((x) => {
        const y = normStr(x.year || '');
        const year = /^\d{4}$/.test(y) ? y : '';
        aoa.push([
            normCode(x.code),
            year,
            normStr(x.name || ''),
            normStr(x.headName || ''),
            normStr(x.headEmail || '').toLowerCase()
        ]);
    });
    return aoa;
}

function csvEscapeField(v) {
    const s = String(v ?? '');
    if (/[;"\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
}

/** @param {string} filename @param {object[]} rows */
export function downloadClassesCsv(filename, rows) {
    try {
        const aoa = classesRowsToAoa(rows);
        const body = aoa.map((line) => line.map(csvEscapeField).join(';')).join('\r\n');
        const blob = new Blob(['\ufeff' + body], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename || 'Klassenliste.csv';
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 250);
        return true;
    } catch {
        return false;
    }
}

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
