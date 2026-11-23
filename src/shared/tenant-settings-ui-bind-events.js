/**
 * tenant-settings-ui-bind-events.js — thin orchestrator (R20/c6).
 * @param {Record<string, any>} ctx
 */
import { installTenantBindEventsWriteback } from './tenant-settings-ui-bind-events-writeback.js';
import { installTenantBindEventsFiles } from './tenant-settings-ui-bind-events-files.js';
import { installTenantBindEventsCatalog } from './tenant-settings-ui-bind-events-catalog.js';
import { installTenantBindEventsPeople } from './tenant-settings-ui-bind-events-people.js';
import { installTenantBindEventsGroups } from './tenant-settings-ui-bind-events-groups.js';
import { installTenantBindEventsInventory } from './tenant-settings-ui-bind-events-inventory.js';
import { installTenantBindEventsWebuntis } from './tenant-settings-ui-bind-events-webuntis.js';

export function installTenantBindEvents(ctx) {
    installTenantBindEventsWriteback(ctx);
    installTenantBindEventsFiles(ctx);
    installTenantBindEventsCatalog(ctx);
    installTenantBindEventsPeople(ctx);
    installTenantBindEventsGroups(ctx);
    installTenantBindEventsInventory(ctx);
    installTenantBindEventsWebuntis(ctx);
}
