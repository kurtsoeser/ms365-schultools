/**
 * Praxis-Regeln für Verwaltungsrollen: Einzelbesetzung vs. Team, optionales M365-Objekt.
 */

export const ADMIN_ROLE_SEAT_SINGLE = 'single';
export const ADMIN_ROLE_SEAT_MULTI = 'multi';

export const ADMIN_ROLE_M365_NONE = 'none';
export const ADMIN_ROLE_M365_GROUP = 'group';
export const ADMIN_ROLE_M365_SHARED_MAILBOX = 'sharedMailbox';

const SINGLE_SEAT_HINTS = ['direktion', 'direktor', 'schularzt', 'schulärztin', 'schulaerzt'];

/**
 * @param {unknown} v
 * @returns {'single'|'multi'}
 */
export function normalizeAdminRoleSeatMode(v, role) {
    const t = String(v ?? '')
        .trim()
        .toLowerCase();
    if (t === ADMIN_ROLE_SEAT_SINGLE || t === 'one' || t === '1') return ADMIN_ROLE_SEAT_SINGLE;
    if (t === ADMIN_ROLE_SEAT_MULTI || t === 'many' || t === 'team') return ADMIN_ROLE_SEAT_MULTI;
    return inferAdminRoleSeatMode(role);
}

/**
 * @param {{ name?: string, code?: string }|null|undefined} role
 * @returns {'single'|'multi'}
 */
export function inferAdminRoleSeatMode(role) {
    const parts = [role && role.name, role && role.code].map(function (p) {
        return String(p || '').toLowerCase();
    });
    for (let i = 0; i < parts.length; i++) {
        const r = parts[i];
        if (!r) continue;
        for (let j = 0; j < SINGLE_SEAT_HINTS.length; j++) {
            if (r.indexOf(SINGLE_SEAT_HINTS[j]) !== -1) return ADMIN_ROLE_SEAT_SINGLE;
        }
    }
    return ADMIN_ROLE_SEAT_MULTI;
}

/**
 * @param {unknown} v
 * @returns {'none'|'group'|'sharedMailbox'}
 */
export function normalizeAdminRoleM365Kind(v) {
    const t = String(v ?? '')
        .trim()
        .toLowerCase();
    if (t === ADMIN_ROLE_M365_GROUP || t === 'm365group' || t === 'unified') return ADMIN_ROLE_M365_GROUP;
    if (
        t === ADMIN_ROLE_M365_SHARED_MAILBOX ||
        t === 'sharedmailbox' ||
        t === 'mailbox' ||
        t === 'shared'
    ) {
        return ADMIN_ROLE_M365_SHARED_MAILBOX;
    }
    return ADMIN_ROLE_M365_NONE;
}

/**
 * @param {object|null|undefined} raw
 */
export function normalizeAdminRoleM365Resource(raw) {
    const kind = normalizeAdminRoleM365Kind(raw && (raw.m365ResourceType || raw.m365Kind));
    if (kind === ADMIN_ROLE_M365_NONE) {
        return {
            m365ResourceType: ADMIN_ROLE_M365_NONE,
            m365ResourceId: '',
            m365ResourceEmail: '',
            m365ResourceLabel: ''
        };
    }
    return {
        m365ResourceType: kind,
        m365ResourceId: String((raw && raw.m365ResourceId) || '').trim(),
        m365ResourceEmail: String((raw && raw.m365ResourceEmail) || '')
            .trim()
            .toLowerCase(),
        m365ResourceLabel: String((raw && raw.m365ResourceLabel) || '').trim()
    };
}

/**
 * @param {object|null|undefined} role
 */
export function adminRolePolicyFromRecord(role) {
    const seatMode = normalizeAdminRoleSeatMode(role && role.seatMode, role);
    const m365 = normalizeAdminRoleM365Resource(role);
    return Object.assign({ seatMode: seatMode }, m365);
}

/**
 * @param {'single'|'multi'} seatMode
 * @param {number} personCount
 */
export function canAddPersonToAdminRole(seatMode, personCount) {
    const mode = normalizeAdminRoleSeatMode(seatMode, null);
    const n = typeof personCount === 'number' && personCount >= 0 ? personCount : 0;
    if (mode === ADMIN_ROLE_SEAT_SINGLE) return n < 1;
    return true;
}

/**
 * @param {'single'|'multi'} seatMode
 */
export function adminRoleSeatModeLabel(seatMode) {
    return normalizeAdminRoleSeatMode(seatMode, null) === ADMIN_ROLE_SEAT_SINGLE ? '1 Platz' : 'Team';
}

/**
 * @param {'none'|'group'|'sharedMailbox'} kind
 */
export function adminRoleM365KindShortLabel(kind) {
    const k = normalizeAdminRoleM365Kind(kind);
    if (k === ADMIN_ROLE_M365_GROUP) return 'M365-Gruppe';
    if (k === ADMIN_ROLE_M365_SHARED_MAILBOX) return 'Postfach';
    return '';
}
