/**
 * Schulleitung vs. Schulverwaltung: Rollen-Tier und E-Mail-Sammlung für Sammelgruppen.
 * Schulleitung (Tier): z. B. Direktion – eigene M365-Gruppe, höhere Planer-/Listenrechte.
 * Verwaltung (Tier): Sekretariat, Schularzt, IT-Support, Bibliothek, freie Rollen, …
 */

export const ADMIN_TIER_SCHULLEITUNG = 'schulleitung';
export const ADMIN_TIER_VERWALTUNG = 'verwaltung';

/**
 * @param {unknown} v
 * @returns {'schulleitung'|'verwaltung'|''}
 */
export function normalizeAdminTier(v) {
    const t = String(v ?? '')
        .trim()
        .toLowerCase();
    if (t === ADMIN_TIER_SCHULLEITUNG || t === 'schulfuehrung' || t === 'direktion') return ADMIN_TIER_SCHULLEITUNG;
    if (t === ADMIN_TIER_VERWALTUNG || t === 'admin' || t === 'staff') return ADMIN_TIER_VERWALTUNG;
    return '';
}

/**
 * @param {{ name?: string, code?: string, role?: string, defaultKey?: string, tier?: string }|null|undefined} roleOrRow
 */
export function isDirektionRoleLabel(roleOrRow) {
    const parts = [
        roleOrRow && roleOrRow.name,
        roleOrRow && roleOrRow.code,
        roleOrRow && roleOrRow.role,
        roleOrRow && roleOrRow.defaultKey
    ];
    for (let i = 0; i < parts.length; i++) {
        const r = String(parts[i] || '').toLowerCase();
        if (!r) continue;
        if (r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1) return true;
    }
    return false;
}

/**
 * @param {{ name?: string, code?: string, tier?: string }|null|undefined} role
 * @returns {'schulleitung'|'verwaltung'}
 */
export function inferAdminTierForRole(role) {
    const explicit = normalizeAdminTier(role && role.tier);
    if (explicit) return explicit;
    if (isDirektionRoleLabel(role)) return ADMIN_TIER_SCHULLEITUNG;
    return ADMIN_TIER_VERWALTUNG;
}

/**
 * @param {{ role?: string, name?: string, email?: string, defaultKey?: string }} row
 * @param {Array<{ name?: string, code?: string, tier?: string }>} [roleCatalog]
 * @returns {'schulleitung'|'verwaltung'}
 */
export function inferAdminTierForPersonRow(row, roleCatalog) {
    const catalog = Array.isArray(roleCatalog) ? roleCatalog : [];
    const roleName = String((row && row.role) || '').trim();
    const dk = String((row && row.defaultKey) || '').trim();
    let matched = null;
    if (roleName || dk) {
        const rl = (roleName || dk).toLowerCase();
        matched =
            catalog.find(function (r) {
                const n = String(r.name || '').toLowerCase();
                const c = String(r.code || '').toLowerCase();
                return n === rl || c === rl;
            }) || null;
    }
    if (matched) return inferAdminTierForRole(matched);
    if (isDirektionRoleLabel(row)) return ADMIN_TIER_SCHULLEITUNG;
    return ADMIN_TIER_VERWALTUNG;
}

function normEmail(v) {
    return String(v ?? '').trim().toLowerCase();
}

/**
 * @param {Array<{ role?: string, name?: string, email?: string, defaultKey?: string }>} adminRows
 * @param {Array<{ name?: string, code?: string, tier?: string }>} [roleCatalog]
 * @param {'schulleitung'|'verwaltung'} tier
 * @returns {string[]}
 */
export function collectAdminEmailsForTier(adminRows, roleCatalog, tier) {
    const want = normalizeAdminTier(tier);
    if (!want) return [];
    const seen = new Set();
    const out = [];
    (Array.isArray(adminRows) ? adminRows : []).forEach(function (row) {
        if (inferAdminTierForPersonRow(row, roleCatalog) !== want) return;
        const em = normEmail(row && row.email);
        if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

/**
 * @param {Array<{ role?: string, name?: string, email?: string, defaultKey?: string }>} adminRows
 * @param {Array<{ name?: string, code?: string, tier?: string }>} [roleCatalog]
 */
/**
 * @param {Array<{ role?: string, name?: string, email?: string, defaultKey?: string }>} adminRows
 * @param {Array<{ name?: string, code?: string, tier?: string }>} [roleCatalog]
 * @param {Array<{ groupId: string, email: string }>} [memberships]
 */
export function splitAdminEmailsByAudienceTier(adminRows, roleCatalog, memberships) {
    if (Array.isArray(memberships) && memberships.length) {
        const schulleitung = [];
        const verwaltung = [];
        const seenS = new Set();
        const seenV = new Set();
        memberships.forEach(function (m) {
            if (!m) return;
            const em = normEmail(m.email);
            if (!em || em.indexOf('@') === -1) return;
            const gid = String(m.groupId || '').trim();
            if (gid === ADMIN_TIER_SCHULLEITUNG) {
                if (!seenS.has(em)) {
                    seenS.add(em);
                    schulleitung.push(em);
                }
            } else if (gid === ADMIN_TIER_VERWALTUNG) {
                if (!seenV.has(em)) {
                    seenV.add(em);
                    verwaltung.push(em);
                }
            }
        });
        if (schulleitung.length || verwaltung.length) {
            return { schulleitung, verwaltung };
        }
    }
    return {
        schulleitung: collectAdminEmailsForTier(adminRows, roleCatalog, ADMIN_TIER_SCHULLEITUNG),
        verwaltung: collectAdminEmailsForTier(adminRows, roleCatalog, ADMIN_TIER_VERWALTUNG)
    };
}

/**
 * @param {Array<{ name?: string, code?: string, tier?: string, people?: unknown[] }>} administrationGroups
 */
export function normalizeAdministrationGroupTiers(administrationGroups) {
    return (Array.isArray(administrationGroups) ? administrationGroups : []).map(function (group) {
        const g = group && typeof group === 'object' ? group : {};
        const tier = inferAdminTierForRole({
            name: g.name,
            code: g.code,
            tier: g.tier
        });
        const next = Object.assign({}, g, { tier });
        return next;
    });
}
