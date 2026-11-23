/**
 * tenant-settings-ui-directory.js — thin orchestrator (R20/c6).
 * @param {Record<string, any>} ctx
 */
import { installTenantDirectoryUsers } from './tenant-settings-ui-directory-users.js';
import { installTenantDirectoryClasses } from './tenant-settings-ui-directory-classes.js';
import { installTenantDirectoryArges } from './tenant-settings-ui-directory-arges.js';

export function installTenantDirectoryHelpers(ctx) {
    installTenantDirectoryUsers(ctx);
    installTenantDirectoryClasses(ctx);
    installTenantDirectoryArges(ctx);
}
