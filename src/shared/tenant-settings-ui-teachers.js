/**
 * tenant-settings-ui-teachers.js — thin orchestrator (R20/c6).
 * @param {Record<string, any>} ctx
 */
import { installTenantTeachersAdmin } from './tenant-settings-ui-teachers-admin.js';
import { installTenantTeachersTable } from './tenant-settings-ui-teachers-table.js';

export function installTenantTeachersPanel(ctx) {
    installTenantTeachersAdmin(ctx);
    installTenantTeachersTable(ctx);
}
