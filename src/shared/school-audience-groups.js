/**
 * Kanonische Lehrer-/Schüler-Sammelgruppen aus Stammdaten (setup.matched + catalogLinks).
 * Fallback: legacy ms365-dashboard-audience-groups-v1 bis Stammdaten gesetzt sind.
 */

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function normId(v) {
    const id = String(v || '').trim();
    return GUID_RE.test(id) ? id : '';
}

/**
 * @typedef {{
 *   groupLehrerId?: string,
 *   groupLehrerName?: string,
 *   groupSchuelerId?: string,
 *   groupSchuelerName?: string
 * }} SchoolAudienceGroupsConfig
 */

export function normalizeSchoolAudienceGroups(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    return {
        groupLehrerId: normId(o.groupLehrerId),
        groupLehrerName: String(o.groupLehrerName || '').trim(),
        groupSchuelerId: normId(o.groupSchuelerId),
        groupSchuelerName: String(o.groupSchuelerName || '').trim()
    };
}

function readLegacyDashboardGroups() {
    try {
        const raw = localStorage.getItem('ms365-dashboard-audience-groups-v1');
        if (!raw) return normalizeSchoolAudienceGroups({});
        return normalizeSchoolAudienceGroups(JSON.parse(raw));
    } catch {
        return normalizeSchoolAudienceGroups({});
    }
}

function readFromStammdatenSetup() {
    /** @type {{ groupLehrerId: string, groupLehrerName: string, groupSchuelerId: string, groupSchuelerName: string }} */
    const out = {
        groupLehrerId: '',
        groupLehrerName: '',
        groupSchuelerId: '',
        groupSchuelerName: ''
    };
    try {
        const g = typeof globalThis !== 'undefined' ? globalThis : typeof window !== 'undefined' ? window : {};
        const api = g.ms365AppDataV2;
        if (!api || typeof api.getSetup !== 'function') return out;
        const setup = api.getSetup() || {};
        const matched = setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
        out.groupLehrerId = normId(matched.lehrerGroupId);
        out.groupSchuelerId = normId(matched.schuelerGroupId);
        const links = Array.isArray(setup.catalogLinks) ? setup.catalogLinks : [];
        links.forEach(function (link) {
            if (!link || link.kind !== 'sammelgruppe') return;
            const dn = String(link.displayName || '').trim();
            const mail = String(link.mailNickname || '').trim();
            const label = dn || mail;
            if (link.code === 'lehrer' && label) out.groupLehrerName = label;
            if (link.code === 'schueler' && label) out.groupSchuelerName = label;
        });
        if (typeof api.getCatalogLink === 'function') {
            const le = api.getCatalogLink('sammelgruppe', 'lehrer');
            const sc = api.getCatalogLink('sammelgruppe', 'schueler');
            if (le && !out.groupLehrerName) {
                out.groupLehrerName = String(le.displayName || le.mailNickname || '').trim();
            }
            if (sc && !out.groupSchuelerName) {
                out.groupSchuelerName = String(sc.displayName || sc.mailNickname || '').trim();
            }
        }
    } catch {
        /* ignore */
    }
    return out;
}

/**
 * @typedef {{
 *   groupLehrerId?: string,
 *   groupLehrerName?: string,
 *   groupSchuelerId?: string,
 *   groupSchuelerName?: string
 * }} SchoolAudienceGroupsConfig
 */

/**
 * @returns {SchoolAudienceGroupsConfig}
 */
export function loadSchoolAudienceGroups() {
    const st = readFromStammdatenSetup();
    if (st.groupLehrerId || st.groupSchuelerId) {
        return normalizeSchoolAudienceGroups(st);
    }
    return readLegacyDashboardGroups();
}

/**
 * @param {SchoolAudienceGroupsConfig} [config]
 */
export function schoolAudienceGroupsConfigured(config) {
    const c = normalizeSchoolAudienceGroups(config || loadSchoolAudienceGroups());
    return !!(c.groupLehrerId || c.groupSchuelerId);
}

/**
 * Lehrer-/Schüler-Gruppen aus Stammdaten in Planer-Berechtigungen übernehmen.
 * @param {Record<string, unknown>} base
 */
export function overlaySchoolAudienceOnPermissions(base) {
    const aud = loadSchoolAudienceGroups();
    const out = Object.assign({}, base || {});
    if (aud.groupLehrerId) {
        out.groupLehrerId = aud.groupLehrerId;
        if (aud.groupLehrerName) out.groupLehrer = aud.groupLehrerName;
    }
    if (aud.groupSchuelerId) {
        out.groupSchuelerId = aud.groupSchuelerId;
        if (aud.groupSchuelerName) out.groupSchueler = aud.groupSchuelerName;
    }
    return out;
}
