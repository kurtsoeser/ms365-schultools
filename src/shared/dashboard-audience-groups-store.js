/**
 * Entra-Gruppen für Dashboard-Personas (Lehrkraft / Schüler).
 * Kanonische Quelle: Stammdaten (Lehrer-/Schüler-Sammelgruppe), siehe school-audience-groups.js.
 */
import { notifyAppLocalDataChanged } from './app-local-data-notify.js';
import {
    loadSchoolAudienceGroups,
    schoolAudienceGroupsConfigured,
    normalizeSchoolAudienceGroups
} from './school-audience-groups.js';

export const DASHBOARD_AUDIENCE_GROUPS_KEY = 'ms365-dashboard-audience-groups-v1';

/** @typedef {import('./school-audience-groups.js').SchoolAudienceGroupsConfig} DashboardAudienceGroupsConfig */

export const normalizeDashboardAudienceGroups = normalizeSchoolAudienceGroups;

export function loadDashboardAudienceGroups() {
    return loadSchoolAudienceGroups();
}

/**
 * Legacy: separate Dashboard-Gruppen (nur wenn Stammdaten leer). Bevorzugt Stammdaten pflegen.
 * @param {DashboardAudienceGroupsConfig} config
 */
export function saveDashboardAudienceGroups(config) {
    const payload = normalizeSchoolAudienceGroups(config);
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
    notifyAppLocalDataChanged('dashboard-audience-groups');
    return payload;
}

export { schoolAudienceGroupsConfigured as dashboardAudienceGroupsConfigured };
