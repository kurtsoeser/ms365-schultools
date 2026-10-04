/**
 * Dashboard-Persona aus Entra-Gruppen (Lehrer / Schüler).
 */
import { fetchUserMemberGroupIds } from '../tools/schularbeiten-planer/schularbeiten-planer-entra-role.js';
import {
    loadDashboardAudienceGroups,
    dashboardAudienceGroupsConfigured
} from './dashboard-audience-groups-store.js';

/**
 * @param {Set<string>|string[]} memberIds
 * @param {import('./dashboard-audience-groups-store.js').DashboardAudienceGroupsConfig} config
 * @returns {'lehrer'|'schueler'|null}
 */
export function personaFromDashboardEntraGroups(memberIds, config) {
    const cfg = config || loadDashboardAudienceGroups();
    const member =
        memberIds instanceof Set
            ? memberIds
            : new Set((memberIds || []).map((x) => String(x).toLowerCase()));
    const lehrerId = String(cfg.groupLehrerId || '').trim().toLowerCase();
    const schuelerId = String(cfg.groupSchuelerId || '').trim().toLowerCase();
    if (lehrerId && member.has(lehrerId)) return 'lehrer';
    if (schuelerId && member.has(schuelerId)) return 'schueler';
    return null;
}

/**
 * @returns {Promise<'lehrer'|'schueler'|null>}
 */
export async function resolveDashboardPersonaFromEntraGroups() {
    const cfg = loadDashboardAudienceGroups();
    if (!dashboardAudienceGroupsConfigured(cfg)) return null;
    const ids = [cfg.groupLehrerId, cfg.groupSchuelerId].filter(Boolean);
    const member = await fetchUserMemberGroupIds(ids);
    return personaFromDashboardEntraGroups(member, cfg);
}
