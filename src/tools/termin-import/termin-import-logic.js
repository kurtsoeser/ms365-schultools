/**
 * CSV/Excel-Zeilen → Schultermine-Import.
 * Kein DOM, kein Graph.
 */

function normStr(v) {
    return String(v == null ? '' : v).trim();
}

function parseBool(v) {
    const s = normStr(v).toLowerCase();
    return s === '1' || s === 'true' || s === 'ja' || s === 'yes' || s === 'x' || s === 'ganztägig' || s === 'ganztagig';
}

/**
 * Ziel: school | teachers | both
 * @param {unknown} raw
 */
export function parseTerminTarget(raw) {
    const s = normStr(raw).toLowerCase();
    if (!s) return 'school';
    if (s === 'l' || s === 'lehrer' || s === 'teachers' || s === 'kollegium' || s === 'lehrerkalender') {
        return 'teachers';
    }
    if (s === 'b' || s === 'beide' || s === 'both' || s === 's+l' || s === 'sl') return 'both';
    if (s === 's' || s === 'schule' || s === 'school' || s === 'schulkalender') return 'school';
    return 'school';
}

function parseDateLoose(raw) {
    const s = normStr(raw);
    if (!s) return '';
    // ISO
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s;
    // dd.mm.yyyy or dd.mm.yyyy hh:mm
    const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
    if (m) {
        const dd = m[1].padStart(2, '0');
        const mm = m[2].padStart(2, '0');
        const yyyy = m[3];
        if (m[4] != null) {
            return yyyy + '-' + mm + '-' + dd + 'T' + m[4].padStart(2, '0') + ':' + m[5] + ':00';
        }
        return yyyy + '-' + mm + '-' + dd;
    }
    return s;
}

/**
 * @param {Record<string, string>} row
 */
export function normalizeTerminRow(row) {
    const r = row || {};
    const pick = function () {
        for (let i = 0; i < arguments.length; i++) {
            const k = arguments[i];
            if (r[k] != null && normStr(r[k])) return normStr(r[k]);
        }
        return '';
    };
    const title = pick('title', 'titel', 'betreff', 'name', 'termin', 'subject');
    const start = parseDateLoose(pick('start', 'beginn', 'von', 'startzeit', 'datum'));
    const end = parseDateLoose(pick('end', 'ende', 'bis', 'endzeit')) || start;
    const allDay = parseBool(pick('allday', 'ganztägig', 'ganztagig', 'ganztaegig'));
    const category = pick('kategorie', 'category', 'art') || 'sonstiges';
    const info = pick('info', 'notiz', 'beschreibung', 'description', 'notes');
    const target = parseTerminTarget(pick('ziel', 'target', 'kalender', 'für', 'fuer'));
    const zeitraum = pick('zeitraum', 'zeitraumtext', 'zeit');

    const issues = [];
    if (!title) issues.push('Titel fehlt');
    if (!start) issues.push('Beginn fehlt');

    return {
        title,
        start,
        end,
        allDay,
        category,
        info,
        target,
        zeitraumText: zeitraum,
        ok: issues.length === 0,
        issues
    };
}

/**
 * Header-Zeile + Datenzeilen (Arrays) → Termine.
 * @param {string[]} headers
 * @param {string[][]} dataRows
 */
export function parseTerminTable(headers, dataRows) {
    const hdr = (Array.isArray(headers) ? headers : []).map((h) =>
        normStr(h)
            .toLowerCase()
            .replace(/[\s_]+/g, '')
    );
    const rows = [];
    (Array.isArray(dataRows) ? dataRows : []).forEach(function (cells) {
        if (!cells || !cells.length) return;
        const obj = {};
        for (let i = 0; i < hdr.length; i++) {
            if (!hdr[i]) continue;
            obj[hdr[i]] = cells[i] != null ? String(cells[i]) : '';
        }
        // also keep original keys lightly
        rows.push(normalizeTerminRow(obj));
    });
    const summary = {
        total: rows.length,
        ok: rows.filter((r) => r.ok).length,
        bad: rows.filter((r) => !r.ok).length,
        school: rows.filter((r) => r.ok && (r.target === 'school' || r.target === 'both')).length,
        teachers: rows.filter((r) => r.ok && (r.target === 'teachers' || r.target === 'both')).length
    };
    return { rows, summary };
}

/**
 * Semikolon/CSV-Text parsen (erste Zeile = Header).
 * @param {string} text
 */
export function parseTermineCsvText(text) {
    const raw = String(text || '').replace(/^\uFEFF/, '');
    const lines = raw.split(/\r?\n/).filter((l) => normStr(l));
    if (!lines.length) return { rows: [], summary: { total: 0, ok: 0, bad: 0, school: 0, teachers: 0 } };
    const sep = lines[0].includes(';') ? ';' : ',';
    const split = function (line) {
        // simple split – ausreichend für Schul-Exporte ohne verschachtelte Quotes-Komplexität
        return line.split(sep).map((c) => c.replace(/^"|"$/g, '').trim());
    };
    const headers = split(lines[0]);
    const data = lines.slice(1).map(split);
    return parseTerminTable(headers, data);
}

/**
 * SharePoint List Item fields für Schultermine.
 * @param {ReturnType<typeof normalizeTerminRow>} termin
 */
export function terminToListFields(termin) {
    const t = termin || {};
    const start = t.start || '';
    const end = t.end || start;
    return {
        Title: t.title || '',
        Beginn: start,
        Ende: end,
        Kategorie: t.category || 'sonstiges',
        Info: t.info || '',
        ZeitraumText: t.zeitraumText || '',
        AllDay: !!t.allDay,
        SyncStatus: t.target === 'teachers' ? 'pending' : 'manual',
        SyncError: t.target === 'teachers' || t.target === 'both' ? 'Ziel Lehrerkalender: per Power Automate / Outlook nachziehen' : ''
    };
}

export default {
    parseTerminTarget,
    normalizeTerminRow,
    parseTerminTable,
    parseTermineCsvText,
    terminToListFields
};
