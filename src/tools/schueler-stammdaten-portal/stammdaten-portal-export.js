/**
 * Mapping Payload → Sokrates Einschreiber V20 (67 Spalten).
 */

import { normalizePayload, validateSvnr } from './stammdaten-portal-logic.js';

/** Kurz-Header wie in Musterdatei Blatt Einschreiber (erste Zeile, ohne Kommentar). */
export const EINSCHREIBER_HEADERS = [
    'Familienname',
    'erster Vorname',
    'weitere Vornamen',
    'Geschlecht bitte wählen',
    'Geburtsdatum',
    'Geburtsort',
    'Geburtsstaat',
    'Staatsbürgerschaft',
    'Religionsbekenntnis',
    'Alltagssprache',
    'Erstsprache',
    'SVNR kurz genau 4 Ziffern',
    'SVNr m. Formel',
    'SV Anstalt',
    'BilDok-ErsatzNr',
    'Anmerkung',
    'Mailadresse',
    '2. Mailadresse',
    'Mobiltelefon',
    'Telefonnummer 2',
    'Telefonnummer 3',
    'Adressart',
    'Anrede',
    'Titel',
    'Akad.Grad',
    'Vorname',
    '2. Vorname',
    'Familienname',
    'Akad. Grad nach',
    'Geburtsdatum',
    'Land',
    'PLZ',
    'Ort',
    'Gemeindenummer',
    'Straße',
    'Hausnummer',
    'Mobiltelefon',
    'Telefonnummer 2',
    'Telefonnummer 3',
    'Mailadresse',
    '2. Mailadresse',
    'Erziehungsberechtigt',
    'Zahlungspflichtig',
    'Entscheide an',
    'Adressart',
    'Anrede',
    'Titel',
    'Akad.Grad',
    'Vorname',
    'Vorname 2',
    'Familienname',
    'Akad. Grad nach',
    'Geburtsdatum',
    'Land',
    'PLZ',
    'Ort',
    'Gemeindenummer',
    'Straße',
    'Hausnummer',
    'Mobiltelefon',
    'Telefonnummer 2',
    'Telefonnummer 3',
    'Mailadresse',
    '2. Mailadresse',
    'Erziehungsberechtigt',
    'Zahlungspflichtig',
    'Entscheide an'
];

function fmtDate(value) {
    const s = String(value || '').trim();
    if (!s) return '';
    if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) return s;
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[3]}.${m[2]}.${m[1]}`;
    try {
        const d = new Date(s);
        if (!Number.isNaN(d.getTime())) {
            const dd = String(d.getDate()).padStart(2, '0');
            const mm = String(d.getMonth() + 1).padStart(2, '0');
            return `${dd}.${mm}.${d.getFullYear()}`;
        }
    } catch {
        /* ignore */
    }
    return s;
}

function guardianCells(g, fallbackAddress) {
    const a = g || {};
    const addr = {
        land: a.land || fallbackAddress.land || 'A',
        plz: a.plz || fallbackAddress.plz || '',
        ort: a.ort || fallbackAddress.ort || '',
        gemeindenummer: a.gemeindenummer || fallbackAddress.gemeindenummer || '',
        strasse: a.strasse || fallbackAddress.strasse || '',
        hausnummer: a.hausnummer || fallbackAddress.hausnummer || ''
    };
    return [
        a.adressart || '',
        a.anrede || '',
        a.titel || '',
        a.akadGrad || '',
        a.vorname || '',
        a.vorname2 || '',
        a.familienname || '',
        a.akadGradNach || '',
        fmtDate(a.geburtsdatum),
        addr.land,
        addr.plz,
        addr.ort,
        addr.gemeindenummer,
        addr.strasse,
        addr.hausnummer,
        a.mobil || '',
        a.tel2 || '',
        a.tel3 || '',
        a.mail || '',
        a.mail2 || '',
        a.erziehungsberechtigt || '',
        a.zahlungspflichtig || '',
        a.entscheideAn || ''
    ];
}

/**
 * @param {object} payload
 * @returns {string[]} eine Datenzeile (67 Zellen)
 */
export function payloadToEinschreiberRow(payload) {
    const data = normalizePayload(payload);
    const s = data.schueler;
    const w = data.wohnadresse;
    const sv = validateSvnr(s.svnr);
    const g0 = data.erziehungsberechtigte[0] || {};
    const g1 = data.erziehungsberechtigte[1] || {};

    return [
        s.familienname || '',
        s.vorname || '',
        s.weitereVornamen || '',
        s.geschlecht || '',
        fmtDate(s.geburtsdatum),
        s.geburtsort || '',
        s.geburtsstaat || '',
        s.staatsbuergerschaft || '',
        s.religion || '',
        s.alltagssprache || '',
        s.erstsprache || '',
        sv.kurz || (s.svnr || '').slice(0, 4),
        sv.ok ? sv.svnr : s.svnr || '',
        '',
        '',
        '',
        s.mail || '',
        s.mail2 || '',
        s.mobil || '',
        s.tel2 || '',
        s.tel3 || '',
        ...guardianCells(g0, w),
        ...guardianCells(g1, w)
    ];
}

/**
 * @param {Array<object>} records – { payload, title? }
 * @returns {(string|number)[][]} AOA inkl. Header
 */
export function recordsToEinschreiberAoa(records) {
    const rows = [EINSCHREIBER_HEADERS.slice()];
    for (const rec of records || []) {
        rows.push(payloadToEinschreiberRow(rec.payload || rec));
    }
    return rows;
}

/**
 * Baut eine XLSX-Datei als Blob (Browser, SheetJS global `XLSX`).
 * @param {Array<object>} records
 * @param {{ XLSX?: any, sheetName?: string }} [opts]
 */
export function buildEinschreiberWorkbookBlob(records, opts = {}) {
    const XLSX = opts.XLSX || (typeof globalThis !== 'undefined' ? globalThis.XLSX : null);
    if (!XLSX) throw new Error('SheetJS (XLSX) nicht geladen.');
    const aoa = recordsToEinschreiberAoa(records);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, opts.sheetName || 'Einschreiber');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    return new Blob([out], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
}
