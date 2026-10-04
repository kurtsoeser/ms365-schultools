/**
 * Entra-Gruppen für Dashboard-Personas (Lehrkraft / Schüler).
 * Nur Schul-IT pflegt diese Zuordnung auf der Seite „Dashboard-Werkzeug-Zugriff“.
 */
export const DASHBOARD_AUDIENCE_GROUPS_KEY = 'ms365-dashboard-audience-groups-v1';

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * @typedef {{
 *   groupLehrerId?: string,
 *   groupLehrerName?: string,
 *   groupSchuelerId?: string,
 *   groupSchuelerName?: string
 * }} DashboardAudienceGroupsConfig
 */

export function normalizeDashboardAudienceGroups(raw) {
    const o = raw && typeof raw === 'object' ? raw : {};
    const normId = (v) => {
        const id = String(v || '').trim();
        return GUID_RE.test(id) ? id : '';
    };
    return {
        groupLehrerId: normId(o.groupLehrerId),
        groupLehrerName: String(o.groupLehrerName || '').trim(),
        groupSchuelerId: normId(o.groupSchuelerId),
        groupSchuelerName: String(o.groupSchuelerName || '').trim()
    };
}

export function loadDashboardAudienceGroups() {
    try {
        const raw = localStorage.getItem(DASHBOARD_AUDIENCE_GROUPS_KEY);
        if (!raw) return normalizeDashboardAudienceGroups({});
        return normalizeDashboardAudienceGroups(JSON.parse(raw));
    } catch {
        return normalizeDashboardAudienceGroups({});
    }
}

/**
 * @param {DashboardAudienceGroupsConfig} config
 */
export function saveDashboardAudienceGroups(config) {
    const payload = normalizeDashboardAudienceGroups(config);
    try {
        localStorage.setItem(DASHBOARD_AUDIENCE_GROUPS_KEY, JSON.stringify(payload));
    } catch {
        /* ignore */
    }
    try {
        window.dispatchEvent(new CustomEvent('ms365-dashboard-audience-groups-changed'));
    } catch {
        /* ignore */
    }
    return payload;
}

/**
 * @param {DashboardAudienceGroupsConfig} [config]
 */
export function dashboardAudienceGroupsConfigured(config) {
    const c = normalizeDashboardAudienceGroups(config || loadDashboardAudienceGroups());
    return !!(c.groupLehrerId || c.groupSchuelerId);
}
