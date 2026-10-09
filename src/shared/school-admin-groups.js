/**
 * Schulleitung / Verwaltung (Personal) Sammelgruppen aus Stammdaten-Setup.
 */
import { overlaySchoolAudienceOnPermissions } from './school-audience-groups.js';

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function normId(v) {
    const id = String(v || '').trim();
    return GUID_RE.test(id) ? id : '';
}

function readFromStammdatenSetup() {
    const out = {
        schulleitungGroupId: '',
        schulleitungGroupName: '',
        verwaltungGroupId: '',
        verwaltungGroupName: ''
    };
    try {
        const g = typeof globalThis !== 'undefined' ? globalThis : typeof window !== 'undefined' ? window : {};
        const api = g.ms365AppDataV2;
        if (!api || typeof api.getSetup !== 'function') return out;
        const setup = api.getSetup() || {};
        const matched = setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
        out.schulleitungGroupId = normId(matched.schulleitungGroupId);
        out.verwaltungGroupId = normId(matched.verwaltungGroupId);
        const links = Array.isArray(setup.catalogLinks) ? setup.catalogLinks : [];
        links.forEach(function (link) {
            if (!link || link.kind !== 'sammelgruppe') return;
            const dn = String(link.displayName || '').trim();
            const mail = String(link.mailNickname || '').trim();
            const label = dn || mail;
            if (link.code === 'schulleitung' && label) out.schulleitungGroupName = label;
            if (link.code === 'verwaltung' && label) out.verwaltungGroupName = label;
        });
        if (typeof api.getCatalogLink === 'function') {
            const sl = api.getCatalogLink('sammelgruppe', 'schulleitung');
            const vw = api.getCatalogLink('sammelgruppe', 'verwaltung');
            if (sl && !out.schulleitungGroupName) {
                out.schulleitungGroupName = String(sl.displayName || sl.mailNickname || '').trim();
            }
            if (vw && !out.verwaltungGroupName) {
                out.verwaltungGroupName = String(vw.displayName || vw.mailNickname || '').trim();
            }
        }
    } catch {
        /* ignore */
    }
    return out;
}

/**
 * @returns {{
 *   schulleitungGroupId: string,
 *   schulleitungGroupName: string,
 *   verwaltungGroupId: string,
 *   verwaltungGroupName: string
 * }}
 */
export function loadSchoolAdminGroups() {
    return readFromStammdatenSetup();
}

/**
 * Planer-Spalte „Admin“ / Direktion: Schulleitung-Gruppe aus Stammdaten (nicht Personal-Sammelgruppe).
 * @param {Record<string, unknown>} base
 */
export function overlaySchoolAdminGroupsOnPermissions(base) {
    const aud = loadSchoolAdminGroups();
    const out = Object.assign({}, base || {});
    if (aud.schulleitungGroupId) {
        if (!String(out.groupAdminId || '').trim()) {
            out.groupAdminId = aud.schulleitungGroupId;
            if (aud.schulleitungGroupName) out.groupAdmin = aud.schulleitungGroupName;
        }
        if (!String(out.groupDirektionId || '').trim()) {
            out.groupDirektionId = aud.schulleitungGroupId;
            if (aud.schulleitungGroupName) out.groupDirektion = aud.schulleitungGroupName;
        }
    }
    if (aud.verwaltungGroupId) {
        if (!String(out.groupVerwaltungStaffId || '').trim()) {
            out.groupVerwaltungStaffId = aud.verwaltungGroupId;
            if (aud.verwaltungGroupName) out.groupVerwaltungStaff = aud.verwaltungGroupName;
        }
    }
    return out;
}

/**
 * Lehrer/Schüler + Schulleitung/Verwaltung aus Stammdaten.
 * @param {Record<string, unknown>} base
 */
export function overlayStammdatenAudienceOnPermissions(base) {
    return overlaySchoolAdminGroupsOnPermissions(overlaySchoolAudienceOnPermissions(base));
}

/**
 * IT-/Backup-Gruppe: Schulleitung bevorzugen, sonst Verwaltung (Personal).
 */
export function preferredItStammdatenGroupId() {
    const aud = loadSchoolAdminGroups();
    return aud.schulleitungGroupId || aud.verwaltungGroupId || '';
}
