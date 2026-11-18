/**
 * Rollen für sRDP-Anmeldung (Demo-Umschalter + leichte Heuristik).
 */

export const SRDP_ROLE_ORDER = ['schueler', 'lehrer', 'admin'];

export function readDemoRoleOverrideFromUrl() {
    try {
        const p = new URLSearchParams(typeof window !== 'undefined' ? window.location.search || '' : '');
        return p.get('demoRole') === '1';
    } catch {
        return false;
    }
}

/**
 * Role-Chips nur bei explizitem ?demoRole=1 – nie in der Produktiv-Lehrerwelt.
 * @param {boolean} [_entraConfigured]
 * @param {boolean} [demoOverride]
 */
export function isSrdpDemoRoleUiEnabled(_entraConfigured, demoOverride) {
    return !!demoOverride;
}

export function canUseSrdpRoleSwitcher(planerRoles, demoUi) {
    if (demoUi) return true;
    const roles = Array.isArray(planerRoles) ? planerRoles : [];
    return roles.length > 1;
}

export function sortSrdpRoles(roles) {
    const set = new Set((Array.isArray(roles) ? roles : []).map((r) => String(r || '').toLowerCase()));
    return SRDP_ROLE_ORDER.filter((r) => set.has(r));
}

/**
 * Heuristik ohne Entra-Gruppen: Lehrer aus Stammdaten, sonst Schüler.
 * @param {object} state
 */
export function inferSrdpRolesFromStammdaten(state) {
    const teachers = (state.stammdaten && state.stammdaten.teachers) || [];
    const em = String(state.accountEmail || '')
        .trim()
        .toLowerCase();
    if (!em) return ['schueler'];
    const isTeacher = teachers.some(
        (t) =>
            String(t.email || '')
                .trim()
                .toLowerCase() === em ||
            String(t.upn || '')
                .trim()
                .toLowerCase() === em
    );
    if (isTeacher) return ['lehrer', 'admin'];
    return ['schueler'];
}

export function shellPreferredSrdpRole(roles, saved) {
    const list = Array.isArray(roles) ? roles.map((r) => String(r || '').toLowerCase()) : [];
    const s = String(saved || '').toLowerCase();
    if (s && list.includes(s)) return s;
    if (list.includes('admin')) return 'admin';
    if (list.includes('lehrer')) return 'lehrer';
    if (list.includes('schueler')) return 'schueler';
    return 'schueler';
}
