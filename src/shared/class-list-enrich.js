/**
 * Fehlende Abschlussjahre / Klassenvorstände aus lokalen Verknüpfungen ergänzen
 * (classTeams, classGroupMatchByKey, andere Schuljahre).
 */
import { normStr, normCode } from './utils/strings.js';
import { findClassTeamForClass } from './membership-hygiene.js';

export function abschlussjahrFromMailNickname(nickRaw) {
    const n = String(nickRaw || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '');
    const m = n.match(/^jg(\d{4})/);
    return m ? m[1] : '';
}

export function classRowMatchKey(row) {
    return normCode(row && row.code) || normStr(row && row.name).toUpperCase();
}

/**
 * Klassenzeile aus einer Entra-Gruppe ableiten (Mail-Nickname / Anzeigename).
 * @param {{ mailNickname?: string, displayName?: string }} g
 * @returns {{ code: string, name: string, year: string }}
 */
export function deriveClassEntryFromGroup(g) {
    const nick = normStr(g && g.mailNickname);
    const dn = normStr(g && g.displayName);
    let year = abschlussjahrFromMailNickname(nick);
    let code = '';
    let name = dn;

    const nickLc = nick.toLowerCase();
    const jg = nickLc.match(/^([a-z]{1,8})(\d{4})[-_]?(.*)$/);
    if (jg && (jg[1] === 'jg' || jg[1].startsWith('jg'))) {
        if (!year) year = jg[2];
        const tail = normStr(jg[3]).replace(/^[-_]+/, '');
        if (tail) {
            const compact = tail.replace(/[^0-9a-z]/gi, '');
            code = normCode(compact) || normCode(tail);
        }
    }

    if (!code && year && nickLc.indexOf('jg' + year) === 0) {
        const tail = nickLc.slice(('jg' + year).length).replace(/^[-_]+/, '');
        if (tail) code = normCode(tail.replace(/[^0-9a-z]/gi, ''));
    }

    const classPrefixes = ['klasse-', 'klasse.', 'class-', 'k-'];
    for (let i = 0; i < classPrefixes.length; i++) {
        const p = classPrefixes[i];
        if (nickLc.startsWith(p)) {
            const rest = nickLc.slice(p.length);
            code = normCode(rest.replace(/[^0-9a-z]/gi, ''));
            break;
        }
    }

    if (!code && dn) {
        const spaced = dn.match(/^(\d+)\s+([A-Za-zÄÖÜäöüß]+)\s*$/);
        if (spaced) code = normCode(spaced[1] + spaced[2]);
        const kl = dn.match(/^Klasse\s+([0-9A-Za-z]+)/i);
        if (!code && kl) code = normCode(kl[1]);
        const trailing = dn.match(/^([0-9]+[A-Za-z]+)\s*[-–]\s*Klasse/i);
        if (!code && trailing) code = normCode(trailing[1]);
        const bare = dn.match(/^([0-9]+[A-Za-z]+)$/);
        if (!code && bare) code = normCode(bare[1]);
    }

    if (!code && nick) {
        const stripped = nick.replace(/[^0-9A-Za-z]/g, '');
        if (/^\d+[A-Za-z]+/.test(stripped)) code = normCode(stripped);
    }

    if (!name) name = code;
    return { code, name, year };
}

/**
 * @param {object|null|undefined} yearsByLabel app-data-v2 years.byLabel
 * @returns {Map<string, { year?: string, headName?: string, headEmail?: string }>}
 */
export function collectPriorClassFieldsByCode(yearsByLabel) {
    const map = new Map();
    const by = yearsByLabel && typeof yearsByLabel === 'object' ? yearsByLabel : {};
    Object.keys(by).forEach(function (yearLabel) {
        const bucket = by[yearLabel];
        const list = bucket && Array.isArray(bucket.classes) ? bucket.classes : [];
        list.forEach(function (cl) {
            const code = normCode(cl && cl.code);
            if (!code) return;
            const prev = map.get(code) || {};
            const y = normStr(cl && cl.year);
            if (!prev.year && /^\d{4}$/.test(y)) prev.year = y;
            const hn = normStr(cl && cl.headName);
            if (!prev.headName && hn) prev.headName = hn;
            const he = normStr(cl && cl.headEmail).toLowerCase();
            if (!prev.headEmail && he && he.indexOf('@') !== -1) prev.headEmail = he;
            map.set(code, prev);
        });
    });
    return map;
}

function pickYearForRow(row, team, groupMatch, prior) {
    const existing = normStr(row && row.year);
    if (/^\d{4}$/.test(existing)) return existing;
    if (team) {
        const ty = normStr(team.abschlussJahr || team.year);
        if (/^\d{4}$/.test(ty)) return ty;
        const fromNick = abschlussjahrFromMailNickname(
            team.mailNickname || team.stableMailNickname || team.graphMailNickname
        );
        if (fromNick) return fromNick;
    }
    if (groupMatch) {
        const fromMatch = abschlussjahrFromMailNickname(groupMatch.mailNickname || groupMatch.mail);
        if (fromMatch) return fromMatch;
    }
    const code = normCode(row && row.code);
    if (code && prior && prior.has(code)) {
        const p = prior.get(code);
        if (p && /^\d{4}$/.test(normStr(p.year))) return normStr(p.year);
    }
    return '';
}

function pickKvForRow(row, prior) {
    let headName = normStr(row && row.headName);
    let headEmail = normStr(row && row.headEmail).toLowerCase();
    if (headName && headEmail && headEmail.indexOf('@') !== -1) {
        return { headName, headEmail, changed: false };
    }
    const code = normCode(row && row.code);
    if (code && prior && prior.has(code)) {
        const p = prior.get(code);
        if (p) {
            if (!headName && p.headName) headName = normStr(p.headName);
            if ((!headEmail || headEmail.indexOf('@') === -1) && p.headEmail) {
                headEmail = normStr(p.headEmail).toLowerCase();
            }
        }
    }
    const changed =
        headName !== normStr(row && row.headName) ||
        headEmail !== normStr(row && row.headEmail).toLowerCase();
    return { headName, headEmail, changed };
}

/**
 * @param {object[]} classes
 * @param {{ classTeams?: object[], classGroupMatchByKey?: object, priorByCode?: Map, deriveStableMailNickname?: function }} ctx
 */
export function enrichClassesFromLinkedGroups(classes, ctx) {
    const list = Array.isArray(classes) ? classes.slice() : [];
    const o = ctx && typeof ctx === 'object' ? ctx : {};
    const teams = Array.isArray(o.classTeams) ? o.classTeams : [];
    const matchMap =
        o.classGroupMatchByKey && typeof o.classGroupMatchByKey === 'object' ? o.classGroupMatchByKey : {};
    const prior = o.priorByCode instanceof Map ? o.priorByCode : new Map();
    const deriveNick =
        typeof o.deriveStableMailNickname === 'function' ? o.deriveStableMailNickname : null;

    let changed = false;
    let yearFilled = 0;
    let kvFilled = 0;

    const out = list.map(function (row) {
        const next = Object.assign({}, row);
        const key = classRowMatchKey(row);
        const team = findClassTeamForClass(row, teams);
        const groupMatch = key && matchMap[key] ? matchMap[key] : null;

        const year = pickYearForRow(row, team, groupMatch, prior);
        if (year && year !== normStr(row.year)) {
            next.year = year;
            changed = true;
            yearFilled++;
        }

        const kv = pickKvForRow(next, prior);
        if (kv.changed) {
            if (kv.headName) next.headName = kv.headName;
            if (kv.headEmail && kv.headEmail.indexOf('@') !== -1) next.headEmail = kv.headEmail;
            changed = true;
            kvFilled++;
        }

        if (deriveNick && next.year && next.code && !normStr(next.stableMailNickname)) {
            const nick = deriveNick(next.year, next.code, next);
            if (nick) next.stableMailNickname = nick;
        }

        return next;
    });

    return { classes: out, changed, stats: { yearFilled, kvFilled } };
}

/**
 * Ergänzt KV aus Graph-Gruppenbesitzern (Lehrkraft-E-Mail in Stammdaten).
 * @param {object[]} classes
 * @param {{ classTeams?: object[], teachers?: object[], fetchGroupOwners?: function, getGraphToken?: function }} ctx
 */
export async function enrichClassesKvFromGraphOwners(classes, ctx) {
    const o = ctx && typeof ctx === 'object' ? ctx : {};
    const teams = Array.isArray(o.classTeams) ? o.classTeams : [];
    const teachers = Array.isArray(o.teachers) ? o.teachers : [];
    const teacherByEmail = new Map();
    teachers.forEach(function (t) {
        const em = normStr(t && t.email).toLowerCase();
        if (em && em.indexOf('@') !== -1) teacherByEmail.set(em, t);
    });
    const fetchOwners = o.fetchGroupOwners;
    const getToken = o.getGraphToken;
    if (typeof fetchOwners !== 'function' || typeof getToken !== 'function') {
        return { classes: classes, changed: false, stats: { kvFilled: 0, skipped: 0 } };
    }

    let token;
    try {
        token = await getToken();
    } catch {
        return { classes: classes, changed: false, stats: { kvFilled: 0, skipped: 0, error: 'token' } };
    }

    let changed = false;
    let kvFilled = 0;
    let skipped = 0;

    const out = [];
    for (let i = 0; i < (classes || []).length; i++) {
        const row = classes[i];
        const next = Object.assign({}, row);
        const hasKv =
            normStr(row.headEmail).indexOf('@') !== -1 ||
            (normStr(row.headName) && normStr(row.headEmail).indexOf('@') !== -1);
        if (hasKv) {
            out.push(next);
            continue;
        }

        const team = findClassTeamForClass(row, teams);
        const gid = team && team.graphGroupId ? String(team.graphGroupId).trim() : '';
        if (!gid) {
            skipped++;
            out.push(next);
            continue;
        }

        try {
            const owners = await fetchOwners(token, gid);
            let picked = null;
            (owners || []).forEach(function (u) {
                if (picked) return;
                const em = normStr(u.mail || u.userPrincipalName).toLowerCase();
                if (!em || em.indexOf('@') === -1) return;
                const t = teacherByEmail.get(em);
                if (t) {
                    picked = { headName: normStr(t.name) || normStr(u.displayName), headEmail: em };
                    return;
                }
                if (!picked && teacherByEmail.size === 0) {
                    picked = { headName: normStr(u.displayName), headEmail: em };
                }
            });
            if (picked) {
                next.headName = picked.headName;
                next.headEmail = picked.headEmail;
                changed = true;
                kvFilled++;
            } else {
                skipped++;
            }
        } catch {
            skipped++;
        }
        out.push(next);
    }

    return { classes: out, changed, stats: { kvFilled, skipped } };
}
