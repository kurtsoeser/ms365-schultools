/**
 * tenant-settings-ui-school-groups.js — thin orchestrator (R20/c6).
 * @param {Record<string, any>} ctx
 */
import { installTenantSchoolGroupsMatch } from './tenant-settings-ui-school-groups-match.js';
import { installTenantSchoolGroupsNaming } from './tenant-settings-ui-school-groups-naming.js';
import { installTenantSchoolGroupsLehrer } from './tenant-settings-ui-school-groups-lehrer.js';

export function installTenantSchoolGroupsPanel(ctx) {
    installTenantSchoolGroupsMatch(ctx);
    installTenantSchoolGroupsNaming(ctx);
    installTenantSchoolGroupsLehrer(ctx);
}
