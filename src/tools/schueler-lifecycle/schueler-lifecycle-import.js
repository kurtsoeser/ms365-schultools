/**
 * Schülerlisten aus CSV/XLSX (inkl. WebUntis-Export) für Lifecycle-Diff.
 */

export function normalizeLifecycleStudent(row) {
    const r = row && typeof row === 'object' ? row : {};
    return {
        name: String(r.name || '').trim(),
        email: String(r.email || '').trim().toLowerCase(),
        klasse: String(r.klasse || r.class || '').trim()
    };
}

function headerIndexMap(headerRow) {
    const map = {};
    (headerRow || []).forEach(function (cell, i) {
        const k = String(cell == null ? '' : cell)
            .trim()
            .toLowerCase()
            .replace(/\s+/g, '');
        if (k) map[k] = i;
    });
    return map;
}

function findCol(map, patterns) {
    const keys = Object.keys(map || {});
    for (let p = 0; p < patterns.length; p++) {
        const pat = patterns[p];
        for (let k = 0; k < keys.length; k++) {
            const key = keys[k];
            if (pat.test(key)) return map[key];
        }
    }
    return -1;
}

function cell(row, idx) {
    if (idx < 0 || !row) return '';
    return String(row[idx] == null ? '' : row[idx]).trim();
}

/**
 * Generische Tabelle (erste Zeile = Header) → Lifecycle-Schüler.
 * @param {unknown[][]} aoa
 * @returns {Array<{ name: string, email: string, klasse: string }>}
 */
export function parseStudentsTableAoa(aoa) {
    const rows = Array.isArray(aoa) ? aoa : [];
    if (rows.length < 2) return [];

    const wu =
        typeof globalThis !== 'undefined' && globalThis.ms365WebuntisExportImport
            ? globalThis.ms365WebuntisExportImport
            : typeof window !== 'undefined'
              ? window.ms365WebuntisExportImport
              : null;
    if (wu && typeof wu.detectExportKindFromAoa === 'function' && typeof wu.parseStudentsAoa === 'function') {
        const kind = wu.detectExportKindFromAoa(rows);
        if (kind === 'student') {
            return wu
                .parseStudentsAoa(rows, { includeExited: false })
                .map(normalizeLifecycleStudent)
                .filter(function (s) {
                    return s.email || s.name;
                });
        }
    }

    const map = headerIndexMap(rows[0]);
    const iFore = findCol(map, [/vorname/, /forename/, /firstname/]);
    const iLong = findCol(map, [/nachname/, /familienname/, /lastname/, /longname/]);
    const iName = findCol(map, [/^name$/, /schüler/, /schueler/]);
    const iEmail = findCol(map, [/mail/, /email/, /upn/]);
    const iKlasse = findCol(map, [/klasse/, /class/, /jahrgang/]);

    const out = [];
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !row.length) continue;
        const fore = cell(row, iFore);
        const long = cell(row, iLong);
        const name =
            fore && long
                ? fore + ' ' + long
                : cell(row, iName) || fore || long;
        const email = cell(row, iEmail).toLowerCase();
        const klasse = cell(row, iKlasse);
        if (!email && !name) continue;
        out.push({ name: name, email: email, klasse: klasse });
    }
    return out;
}

/**
 * @param {ArrayBuffer} buf
 * @returns {unknown[][]|null}
 */
export function aoaFromXlsxArrayBuffer(buf) {
    const X = typeof globalThis !== 'undefined' ? globalThis.XLSX : typeof window !== 'undefined' ? window.XLSX : null;
    if (!X || !X.read || !X.utils) return null;
    try {
        const wb = X.read(buf, { type: 'array' });
        const sn = wb.SheetNames && wb.SheetNames[0];
        if (!sn) return [];
        const sheet = wb.Sheets[sn];
        return sheet ? X.utils.sheet_to_json(sheet, { header: 1, defval: '' }) : [];
    } catch {
        return null;
    }
}
