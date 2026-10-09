/**
 * Freie Verwaltungs-Zielgruppen (M365-Sammelgruppen) + Mitgliedschaften pro E-Mail.
 * Builtin: schulleitung, verwaltung. Custom: id va-* mit eigenem graphGroupId im Setup.
 */

import {
    ADMIN_TIER_SCHULLEITUNG,
    ADMIN_TIER_VERWALTUNG,
    inferAdminTierForPersonRow,
    inferAdminTierForRole,
    normalizeAdminTier
} from './administration-audience-logic.js';

export const BUILTIN_AUDIENCE_SCHULLEITUNG = 'schulleitung';
export const BUILTIN_AUDIENCE_VERWALTUNG = 'verwaltung';

export function isBuiltinAudienceGroupId(id) {
    const g = String(id || '').trim();
    return g === BUILTIN_AUDIENCE_SCHULLEITUNG || g === BUILTIN_AUDIENCE_VERWALTUNG;
}

/**
 * @returns {Array<{ id: string, label: string, builtin: boolean, order: number }>}
 */
export function defaultVerwaltungAudienceGroups() {
    return [
        { id: BUILTIN_AUDIENCE_SCHULLEITUNG, label: 'Schulleitung', builtin: true, order: 0 },
        { id: BUILTIN_AUDIENCE_VERWALTUNG, label: 'Verwaltung (Personal)', builtin: true, order: 1 }
    ];
}

function normEmail(v) {
    return String(v ?? '').trim().toLowerCase();
}

function newCustomGroupId() {
    return 'va-' + Math.random().toString(36).slice(2, 10);
}

/**
 * @param {unknown} groupsIn
 */
export function normalizeVerwaltungAudienceGroups(groupsIn) {
    const builtins = defaultVerwaltungAudienceGroups();
    const byId = new Map();
    builtins.forEach(function (g) {
        byId.set(g.id, Object.assign({}, g));
    });
    (Array.isArray(groupsIn) ? groupsIn : []).forEach(function (raw, idx) {
        if (!raw || typeof raw !== 'object') return;
        const id = String(raw.id || '').trim();
        if (!id || isBuiltinAudienceGroupId(id)) {
            if (id && byId.has(id) && raw.label) {
                byId.set(id, Object.assign({}, byId.get(id), { label: String(raw.label).trim() || byId.get(id).label }));
            }
            return;
        }
        if (!id.startsWith('va-')) return;
        const label = String(raw.label || raw.name || '').trim();
        if (!label) return;
        const order = typeof raw.order === 'number' ? raw.order : 100 + idx;
        byId.set(id, {
            id: id,
            label: label,
            builtin: false,
            order: order,
            mailNick: raw.mailNick ? String(raw.mailNick).trim() : '',
            newDisplayName: raw.newDisplayName ? String(raw.newDisplayName).trim() : ''
        });
    });
    return Array.from(byId.values()).sort(function (a, b) {
        return (a.order || 0) - (b.order || 0);
    });
}

/**
 * @param {unknown} membershipsIn
 * @returns {Array<{ groupId: string, email: string, name?: string }>}
 */
export function normalizeAdminAudienceMemberships(membershipsIn) {
    const seen = new Set();
    const out = [];
    (Array.isArray(membershipsIn) ? membershipsIn : []).forEach(function (row) {
        if (!row || typeof row !== 'object') return;
        const groupId = String(row.groupId || row.audienceGroupId || '').trim();
        const email = normEmail(row.email);
        if (!groupId || !email || email.indexOf('@') === -1) return;
        const key = groupId + '\u0001' + email;
        if (seen.has(key)) return;
        seen.add(key);
        const name = String(row.name || '').trim();
        const rec = { groupId: groupId, email: email };
        if (name) rec.name = name;
        out.push(rec);
    });
    return out;
}

/**
 * @param {string} groupIdField
 * @param {string} [matchedSchulleitung]
 * @param {string} [matchedVerwaltung]
 * @param {Record<string, string>} [customMap]
 */
export function resolveAudienceGroupGraphId(groupId, matchedSchulleitung, matchedVerwaltung, customMap) {
    const id = String(groupId || '').trim();
    if (id === BUILTIN_AUDIENCE_SCHULLEITUNG) return String(matchedSchulleitung || '').trim() || null;
    if (id === BUILTIN_AUDIENCE_VERWALTUNG) return String(matchedVerwaltung || '').trim() || null;
    const map = customMap && typeof customMap === 'object' ? customMap : {};
    const gid = map[id] ? String(map[id]).trim() : '';
    return gid || null;
}

/**
 * @param {Array<{ groupId: string, email: string, name?: string }>} memberships
 * @param {string} groupId
 * @returns {string[]}
 */
/**
 * @param {Array<{ code?: string, name?: string, personName?: string, email?: string }>} displayRows
 * @param {string} emailRaw
 */
export function findAdminDisplayRowByEmail(displayRows, emailRaw) {
    const em = normEmail(emailRaw);
    if (!em) return null;
    const rows = Array.isArray(displayRows) ? displayRows : [];
    for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        if (r && normEmail(r.email) === em) return r;
    }
    return null;
}

export function collectEmailsForAudienceGroup(memberships, groupId) {
    const want = String(groupId || '').trim();
    if (!want) return [];
    const seen = new Set();
    const out = [];
    (Array.isArray(memberships) ? memberships : []).forEach(function (m) {
        if (!m || String(m.groupId) !== want) return;
        const em = normEmail(m.email);
        if (!em || seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

/**
 * @param {Array<{ role?: string, name?: string, email?: string, defaultKey?: string }>} adminRows
 * @param {Array<{ name?: string, code?: string, tier?: string }>} roleCatalog
 * @param {Array<{ groupId: string, email: string }>} [memberships]
 */
export function splitAdminEmailsByAudienceGroups(adminRows, roleCatalog, memberships) {
    const mem = normalizeAdminAudienceMemberships(memberships);
    if (mem.length) {
        return {
            schulleitung: collectEmailsForAudienceGroup(mem, BUILTIN_AUDIENCE_SCHULLEITUNG),
            verwaltung: collectEmailsForAudienceGroup(mem, BUILTIN_AUDIENCE_VERWALTUNG)
        };
    }
    const schulleitung = [];
    const verwaltung = [];
    const seenS = new Set();
    const seenV = new Set();
    (Array.isArray(adminRows) ? adminRows : []).forEach(function (row) {
        const em = normEmail(row && row.email);
        if (!em || em.indexOf('@') === -1) return;
        const tier = inferAdminTierForPersonRow(row, roleCatalog);
        if (tier === ADMIN_TIER_SCHULLEITUNG) {
            if (!seenS.has(em)) {
                seenS.add(em);
                schulleitung.push(em);
            }
        } else {
            if (!seenV.has(em)) {
                seenV.add(em);
                verwaltung.push(em);
            }
        }
    });
    return { schulleitung, verwaltung };
}

/**
 * Erzeugt Mitgliedschaften aus Rollen-Tier (einmalig / wenn noch leer).
 * @param {Array<{ role?: string, name?: string, email?: string, defaultKey?: string }>} adminRows
 * @param {Array<{ name?: string, code?: string, tier?: string }>} roleCatalog
 */
export function buildMembershipsFromRoleTiers(adminRows, roleCatalog) {
    const out = [];
    const seen = new Set();
    (Array.isArray(adminRows) ? adminRows : []).forEach(function (row) {
        const em = normEmail(row && row.email);
        if (!em || em.indexOf('@') === -1) return;
        const tier = inferAdminTierForPersonRow(row, roleCatalog);
        const groupId = tier === ADMIN_TIER_SCHULLEITUNG ? BUILTIN_AUDIENCE_SCHULLEITUNG : BUILTIN_AUDIENCE_VERWALTUNG;
        const key = groupId + '\u0001' + em;
        if (seen.has(key)) return;
        seen.add(key);
        const name = String(row && row.name || '').trim();
        const rec = { groupId: groupId, email: em };
        if (name) rec.name = name;
        out.push(rec);
    });
    return out;
}

/**
 * @param {string} email
 * @param {Array<{ groupId: string, email: string }>} memberships
 * @param {Array<{ id: string, label: string }>} groups
 */
export function audienceGroupIdsForEmail(email, memberships, groups) {
    const em = normEmail(email);
    if (!em) return [];
    const ids = [];
    (Array.isArray(memberships) ? memberships : []).forEach(function (m) {
        if (normEmail(m.email) === em) ids.push(String(m.groupId));
    });
    const catalog = Array.isArray(groups) ? groups : [];
    return ids.filter(function (id) {
        return catalog.some(function (g) {
            return g && g.id === id;
        });
    });
}

/**
 * @param {string} email
 * @param {string[]} groupIds
 * @param {string} [name]
 * @param {Array<{ groupId: string, email: string, name?: string }>} memberships
 */
export function setAudienceGroupsForEmail(email, groupIds, name, memberships) {
    const em = normEmail(email);
    const base = normalizeAdminAudienceMemberships(memberships).filter(function (m) {
        return normEmail(m.email) !== em;
    });
    const want = new Set(
        (Array.isArray(groupIds) ? groupIds : [])
            .map(function (id) {
                return String(id || '').trim();
            })
            .filter(Boolean)
    );
    want.forEach(function (gid) {
        const rec = { groupId: gid, email: em };
        if (name) rec.name = String(name).trim();
        base.push(rec);
    });
    return normalizeAdminAudienceMemberships(base);
}

/**
 * @param {string} label
 * @param {Array<{ id: string, label: string, builtin?: boolean }>} groups
 */
export function addCustomVerwaltungAudienceGroup(label, groups) {
    const next = normalizeVerwaltungAudienceGroups(groups);
    const clean = String(label || '').trim();
    if (!clean) return { groups: next, id: null };
    const id = newCustomGroupId();
    next.push({ id: id, label: clean, builtin: false, order: 100 + next.length });
    return { groups: next, id: id };
}

/**
 * @param {string} groupId
 * @param {Array<{ id: string, builtin?: boolean }>} groups
 */
export function removeCustomVerwaltungAudienceGroup(groupId, groups, memberships) {
    const id = String(groupId || '').trim();
    if (!id || isBuiltinAudienceGroupId(id)) {
        return { groups: normalizeVerwaltungAudienceGroups(groups), memberships: normalizeAdminAudienceMemberships(memberships) };
    }
    const nextG = normalizeVerwaltungAudienceGroups(groups).filter(function (g) {
        return g.id !== id;
    });
    const nextM = normalizeAdminAudienceMemberships(memberships).filter(function (m) {
        return m.groupId !== id;
    });
    return { groups: nextG, memberships: nextM };
}

/**
 * 5. Textfeld: schulleitung | verwaltung | va-xxx | kommagetrennt
 * @param {string} raw
 * @returns {string[]}
 */
export function parseAudienceGroupIdsField(raw) {
    const s = String(raw || '').trim();
    if (!s) return [];
    return s
        .split(/[,;|]/)
        .map(function (p) {
            const t = p.trim().toLowerCase();
            if (!t) return '';
            const tier = normalizeAdminTier(t);
            if (tier) return tier;
            if (t.startsWith('va-')) return t;
            return '';
        })
        .filter(Boolean);
}

export function formatAudienceGroupIdsField(groupIds) {
    return (Array.isArray(groupIds) ? groupIds : [])
        .map(function (id) {
            return String(id || '').trim();
        })
        .filter(Boolean)
        .join(',');
}

/**
 * @param {object} settings
 */
export function ensureAdminAudienceOnSettings(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const groups = normalizeVerwaltungAudienceGroups(s.verwaltungAudienceGroups);
    let memberships = normalizeAdminAudienceMemberships(s.adminAudienceMemberships);
    const admin = Array.isArray(s.admin) ? s.admin : [];
    const roles = Array.isArray(s.adminRoles) ? s.adminRoles : [];
    if (!memberships.length && admin.length) {
        memberships = buildMembershipsFromRoleTiers(admin, roles);
    }
    return { groups, memberships };
}
