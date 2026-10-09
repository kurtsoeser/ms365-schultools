/**
 * Lehrkraft-Kürzel, Name und E-Mail für Unterrichtsbelegung ergänzen (Stammdaten / M365-Besitzer).
 */

function normStr(v) {
    return String(v ?? '').trim();
}

function normCode(v) {
    return normStr(v).toUpperCase();
}

function normMail(v) {
    return normStr(v).toLowerCase();
}

/**
 * @param {Array<{ code?: string, email?: string, name?: string }>} teachers
 */
export function buildTeacherLookupFromList(teachers) {
    const byCode = new Map();
    const byEmail = new Map();
    (Array.isArray(teachers) ? teachers : []).forEach((t) => {
        if (!t || typeof t !== 'object') return;
        const code = normCode(t.code);
        const email = normMail(t.email);
        const name = normStr(t.name);
        const entry = { code, email, name };
        if (code) byCode.set(code, entry);
        if (email) byEmail.set(email, entry);
    });
    return { byCode, byEmail };
}

/**
 * @param {Map<string, { email?: string, name?: string }>|null|undefined} teacherByCode
 */
export function buildTeacherLookupFromCodeMap(teacherByCode) {
    const teachers = [];
    if (teacherByCode && typeof teacherByCode.forEach === 'function') {
        teacherByCode.forEach((v, code) => {
            teachers.push({
                code,
                email: v && v.email,
                name: v && v.name
            });
        });
    }
    return buildTeacherLookupFromList(teachers);
}

/**
 * @param {object} row
 * @param {{ byCode?: Map<string, object>, byEmail?: Map<string, object> }} lookup
 * @param {{ ownerEmail?: string }} [options]
 */
export function enrichBelegungRow(row, lookup, options) {
    const r = row && typeof row === 'object' ? Object.assign({}, row) : {};
    const opts = options && typeof options === 'object' ? options : {};
    const ownerEmail = normMail(opts.ownerEmail);
    if (ownerEmail && !normMail(r.lehrerEmail)) {
        r.lehrerEmail = ownerEmail;
    }

    const byCode = lookup && lookup.byCode;
    const byEmail = lookup && lookup.byEmail;

    const code = normCode(r.lehrerCode);
    if (code && byCode && typeof byCode.get === 'function') {
        const t = byCode.get(code);
        if (t) {
            if (!normMail(r.lehrerEmail) && t.email) r.lehrerEmail = normMail(t.email);
            if (!normStr(r.lehrerName) && t.name) r.lehrerName = normStr(t.name);
        }
    }

    const mail = normMail(r.lehrerEmail);
    if (mail && byEmail && typeof byEmail.get === 'function') {
        const t = byEmail.get(mail);
        if (t) {
            if (!normCode(r.lehrerCode) && t.code) r.lehrerCode = normCode(t.code);
            if (!normStr(r.lehrerName) && t.name) r.lehrerName = normStr(t.name);
            if (!normMail(r.lehrerEmail) && t.email) r.lehrerEmail = normMail(t.email);
        }
    }

    return r;
}

/**
 * @param {object[]} rows
 * @param {{ byCode?: Map<string, object>, byEmail?: Map<string, object> }} lookup
 * @param {Map<string, string>|Record<string, string>} [ownerEmailByGroupId]
 */
export function enrichBelegungRows(rows, lookup, ownerEmailByGroupId) {
    const owners =
        ownerEmailByGroupId instanceof Map
            ? ownerEmailByGroupId
            : ownerEmailByGroupId && typeof ownerEmailByGroupId === 'object'
              ? new Map(Object.entries(ownerEmailByGroupId))
              : new Map();

    return (Array.isArray(rows) ? rows : []).map((row) => {
        const gid = normStr(row && row.graphGroupId);
        const ownerEmail = gid && owners.has(gid) ? owners.get(gid) : '';
        return enrichBelegungRow(row, lookup, { ownerEmail });
    });
}

/**
 * @param {object[]} rows
 * @param {{ byCode?: Map<string, object>, byEmail?: Map<string, object> }} lookup
 * @returns {{ rows: object[], changed: number }}
 */
export function enrichBelegungRowsInPlace(rows, lookup) {
    let changed = 0;
    const next = (Array.isArray(rows) ? rows : []).map((row) => {
        const enriched = enrichBelegungRow(row, lookup, {});
        if (
            normStr(enriched.lehrerName) !== normStr(row.lehrerName) ||
            normMail(enriched.lehrerEmail) !== normMail(row.lehrerEmail) ||
            normCode(enriched.lehrerCode) !== normCode(row.lehrerCode)
        ) {
            changed++;
        }
        return enriched;
    });
    return { rows: next, changed };
}
