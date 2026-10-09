/**
 * Bestehende Kursteams aus Microsoft 365 (Graph) erkennen, parsen und in die
 * Unterrichtsbelegung mergen (DisplayName + mailNickname nach Schulschema).
 */

function normStr(v) {
    return String(v ?? '').trim();
}

function normCode(v) {
    return normStr(v).toUpperCase();
}

function normMailNick(v) {
    return normStr(v).toLowerCase();
}

/**
 * Drittes Segment im Anzeigenamen: „BW-FRECH“ → Fach + Lehrkraft-Kürzel.
 * @param {string} segment
 */
export function parseFachLehrerSegment(segment) {
    const s = normStr(segment);
    if (!s) return { fach: '', lehrerCode: '' };
    const idx = s.lastIndexOf('-');
    if (idx <= 0) return { fach: s, lehrerCode: '' };
    const fach = s.slice(0, idx);
    const lehrerCode = normCode(s.slice(idx + 1));
    return { fach, lehrerCode };
}

/**
 * Anzeigename: „SJ26-27 | 1A | BW-FRECH“
 * @param {string} displayName
 * @returns {{ ok: boolean, yearPrefix?: string, klasse?: string, fach?: string, lehrerCode?: string, tail?: string }}
 */
export function parseKursteamDisplayName(displayName) {
    const dn = normStr(displayName);
    const m = dn.match(/^(.+?)\s\|\s(.+?)\s\|\s(.+)$/);
    if (!m) return { ok: false };
    const yearPrefix = normStr(m[1]);
    const klasse = normStr(m[2]);
    const tail = normStr(m[3]);
    const { fach, lehrerCode } = parseFachLehrerSegment(tail);
    return {
        ok: true,
        yearPrefix,
        klasse,
        fach,
        lehrerCode,
        tail
    };
}

/**
 * Mögliche Nickname-Präfixe für ein Schuljahr-Label (z. B. „DEMO SJ26-27“ → demo-sj26-27 und demosj26-27).
 * @param {string} yearPrefix
 * @returns {string[]}
 */
export function kursteamMailNickPrefixVariants(yearPrefix) {
    const yp = normStr(yearPrefix);
    if (!yp) return [];
    const lower = yp.toLowerCase();
    const compact = lower.replace(/\s+/g, '');
    const hyphen = lower.replace(/\s+/g, '-').replace(/-+/g, '-');
    const out = [];
    if (compact) out.push(compact);
    if (hyphen && hyphen !== compact) out.push(hyphen);
    return out;
}

/**
 * @param {string} displayName
 * @param {string} yearPrefix
 */
export function displayNameMatchesKursteamYearPrefix(displayName, yearPrefix) {
    const dn = normStr(displayName);
    const yp = normStr(yearPrefix);
    if (!dn || !yp) return false;
    const d = dn.toLowerCase();
    const p = yp.toLowerCase();
    return d === p || d.startsWith(p + ' |') || d.startsWith(p + '|');
}

/**
 * @param {string} mailNickname
 * @param {{ yearPrefix?: string, requirePipeInDisplayName?: boolean }} [options]
 */
export function mailNicknameMatchesKursteamFilter(mailNickname, options) {
    const nick = normMailNick(mailNickname);
    if (!nick) return false;
    const yp = normStr(options && options.yearPrefix);
    if (yp) {
        const variants = kursteamMailNickPrefixVariants(yp);
        if (!variants.some((prefix) => nick.startsWith(prefix))) return false;
    } else if (!/^sj\d{2}-\d{2}/i.test(nick) && !/^demo-sj\d{2}-\d{2}/i.test(nick)) {
        return false;
    }
    return true;
}

/**
 * @param {{ id?: string, displayName?: string, mailNickname?: string }} group
 * @param {{ yearPrefix?: string, teacherByCode?: Map<string, { email?: string }>, linkedAt?: string }} [options]
 */
export function graphGroupToBelegungRow(group, options) {
    const g = group && typeof group === 'object' ? group : {};
    const mailNickname = normMailNick(g.mailNickname);
    const displayName = normStr(g.displayName);
    const graphGroupId = normStr(g.id);
    if (!mailNickname || !graphGroupId) return null;

    const opts = options && typeof options === 'object' ? options : {};
    const nickOk = mailNicknameMatchesKursteamFilter(mailNickname, { yearPrefix: opts.yearPrefix });
    const dnOk = displayNameMatchesKursteamYearPrefix(displayName, opts.yearPrefix);
    if (!nickOk && !(dnOk && opts.yearPrefix)) {
        if (opts.yearPrefix) return null;
        if (!nickOk) return null;
    }
    if (!opts.yearPrefix && !nickOk) return null;

    const parsed = parseKursteamDisplayName(displayName);
    let klasse = '';
    let fach = '';
    let lehrerCode = '';
    let yearPrefix = normStr(opts.yearPrefix);
    let parseOk = false;

    if (parsed.ok) {
        parseOk = true;
        klasse = parsed.klasse;
        fach = parsed.fach;
        lehrerCode = parsed.lehrerCode;
        if (!yearPrefix) yearPrefix = parsed.yearPrefix;
    }

    let lehrerEmail = '';
    let lehrerName = '';
    const byCode = opts.teacherByCode;
    if (lehrerCode && byCode && typeof byCode.get === 'function') {
        const t = byCode.get(normCode(lehrerCode));
        if (t && t.email) lehrerEmail = normStr(t.email).toLowerCase();
        if (t && t.name) lehrerName = normStr(t.name);
    }
    const ownerEmail = normStr(opts.ownerEmail).toLowerCase();
    if (ownerEmail && !lehrerEmail) lehrerEmail = ownerEmail;

    const linkedAt = normStr(opts.linkedAt) || new Date().toISOString();

    return {
        klasse,
        lehrerCode,
        lehrerEmail,
        lehrerName,
        fach,
        gruppe: '',
        teamName: displayName || mailNickname,
        gruppenmail: mailNickname,
        graphGroupId,
        linkedAt,
        parseOk
    };
}

/**
 * @param {Array<{ id?: string, displayName?: string, mailNickname?: string }>} groups
 * @param {{ yearPrefix?: string, teacherByCode?: Map<string, { email?: string }> }} [options]
 */
export function buildBelegungRowsFromGraphGroups(groups, options) {
    const out = [];
    const seen = new Set();
    (Array.isArray(groups) ? groups : []).forEach((g) => {
        const row = graphGroupToBelegungRow(g, options);
        if (!row) return;
        if (seen.has(row.gruppenmail)) return;
        seen.add(row.gruppenmail);
        out.push(row);
    });
    out.sort((a, b) => {
        const c = a.klasse.localeCompare(b.klasse, 'de');
        if (c !== 0) return c;
        const f = a.fach.localeCompare(b.fach, 'de');
        if (f !== 0) return f;
        return a.gruppenmail.localeCompare(b.gruppenmail, 'de');
    });
    return out;
}

/**
 * @param {object|null} existingSnapshot
 * @param {object[]} importRows
 * @param {{ yearPrefix?: string, source?: string }} [meta]
 */
export function mergeBelegungWithGraphImport(existingSnapshot, importRows, meta) {
    const m = meta && typeof meta === 'object' ? meta : {};
    const yearPrefix = normStr(m.yearPrefix);
    const source = normStr(m.source) || 'graph-import';

    const prev =
        existingSnapshot && typeof existingSnapshot === 'object' && Array.isArray(existingSnapshot.rows)
            ? existingSnapshot.rows
            : [];

    const byMail = new Map();
    prev.forEach((r) => {
        const nick = normMailNick(r && r.gruppenmail);
        if (!nick) return;
        byMail.set(nick, Object.assign({}, r));
    });

    let added = 0;
    let updated = 0;

    (Array.isArray(importRows) ? importRows : []).forEach((row) => {
        const nick = normMailNick(row && row.gruppenmail);
        if (!nick) return;
        const ex = byMail.get(nick);
        if (!ex) {
            byMail.set(nick, Object.assign({}, row));
            added++;
            return;
        }
        const next = Object.assign({}, ex);
        let changed = false;
        if (row.graphGroupId && next.graphGroupId !== row.graphGroupId) {
            next.graphGroupId = row.graphGroupId;
            changed = true;
        }
        if (row.linkedAt && !next.linkedAt) {
            next.linkedAt = row.linkedAt;
            changed = true;
        }
        if (!next.teamName && row.teamName) {
            next.teamName = row.teamName;
            changed = true;
        }
        if (!next.klasse && row.klasse) {
            next.klasse = row.klasse;
            changed = true;
        }
        if (!next.fach && row.fach) {
            next.fach = row.fach;
            changed = true;
        }
        if (!next.lehrerCode && row.lehrerCode) {
            next.lehrerCode = row.lehrerCode;
            changed = true;
        }
        if (!next.lehrerEmail && row.lehrerEmail) {
            next.lehrerEmail = row.lehrerEmail;
            changed = true;
        }
        if (!next.lehrerName && row.lehrerName) {
            next.lehrerName = row.lehrerName;
            changed = true;
        }
        if (changed) {
            byMail.set(nick, next);
            updated++;
        }
    });

    const rows = Array.from(byMail.values());
    rows.sort((a, b) => {
        const c = String(a.klasse || '').localeCompare(String(b.klasse || ''), 'de');
        if (c !== 0) return c;
        const f = String(a.fach || '').localeCompare(String(b.fach || ''), 'de');
        if (f !== 0) return f;
        return String(a.gruppenmail || '').localeCompare(String(b.gruppenmail || ''), 'de');
    });

    const classSet = new Set();
    const teacherSet = new Set();
    rows.forEach((r) => {
        if (r.klasse) classSet.add(r.klasse);
        if (r.lehrerEmail) teacherSet.add(r.lehrerEmail);
        else if (r.lehrerCode) teacherSet.add(r.lehrerCode);
    });

    const linkedCount = rows.filter((r) => normStr(r.graphGroupId)).length;

    return {
        snapshot: {
            updatedAt: new Date().toISOString(),
            yearPrefix: yearPrefix || normStr(existingSnapshot && existingSnapshot.yearPrefix),
            source,
            teamsCount: rows.length,
            classCount: classSet.size,
            teacherCount: teacherSet.size,
            rows
        },
        stats: { added, updated, total: rows.length, linkedCount }
    };
}
