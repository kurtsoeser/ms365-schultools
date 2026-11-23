/**
 * tenant-settings-ui-render.js — thin orchestrator (R20/c6).
 * @param {Record<string, any>} ctx
 */
import { installTenantRenderStorageCore } from './tenant-settings-ui-render-storage.js';
import { installTenantRenderImportRoute } from './tenant-settings-ui-render-import-route.js';
import { installTenantRenderImportRows } from './tenant-settings-ui-render-import-rows.js';

export function installTenantRenderStorage(ctx) {
    installTenantRenderStorageCore(ctx);
    installTenantRenderImportRoute(ctx);
    installTenantRenderImportRows(ctx);
}
