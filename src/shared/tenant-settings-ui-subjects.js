/**
 * Fächer-Tab: Excel-Vorlage und Import-Spalten (Schulregister).
 * Eine Quelle für Tab „Fächer“, Gesamt-Vorlage (Blatt Faecher) und Tests.
 */
import { normStr, normCode, normHeaderKey } from './utils/strings.js';

export const SUBJECT_LIST_XLSX_FILENAME = 'Faecherliste-Vorlage.xlsx';
export const SUBJECT_LIST_SHEET_NAME = 'Faecher';

/** Spalten der offiziellen Vorlage – müssen vom Import erkannt werden. */
export const SUBJECT_LIST_TEMPLATE_HEADERS = ['Kürzel', 'Name'];

const CODE_FIELDS = ['kürzel', 'kuerzel', 'code', 'fach', 'abk', 'abkuerzung', 'abbrev', 'abbreviation'];
const NAME_FIELDS = ['name', 'fachname', 'bezeichnung', 'subject', 'subjectname', 'displayname'];

function getField(row, candidates) {
    if (!row || typeof row !== 'object') return '';
    const map = new Map();
    Object.keys(row).forEach((k) => map.set(normHeaderKey(k), row[k]));
    for (const c of candidates) {
        const v = map.get(normHeaderKey(c));
        if (v != null && String(v).trim() !== '') return String(v).trim();
    }
    return '';
}

/** @returns {string[][]} */
export function subjectListTemplateAoa() {
    return [
        SUBJECT_LIST_TEMPLATE_HEADERS.slice(),
        ['D', 'Deutsch'],
        ['M', 'Mathematik'],
        ['E', 'Englisch']
    ];
}

/**
 * Erste Arbeitsblatt-Zeilen aus Excel/CSV (sheet_to_json).
 * @param {object[]} jsonRows
 * @returns {Array<{ code: string, name: string }>}
 */
export function subjectsFromSpreadsheetJsonRows(jsonRows) {
    const out = [];
    (jsonRows || []).forEach((r) => {
        const code = getField(r, CODE_FIELDS);
        const name = getField(r, NAME_FIELDS);
        const c = normCode(code);
        if (!c) return;
        out.push({ code: c, name: normStr(name) });
    });
    return out;
}

/**
 * Textfeld-Format (eine Zeile pro Fachgruppe): Kürzel;Name
 * @param {Array<{ code?: string, name?: string }>} rows
 */
export function subjectsToLines(rows) {
    return (rows || [])
        .map((x) => `${normCode(x.code)};${normStr(x.name || '')}`.trim())
        .filter(Boolean)
        .join('\n');
}
