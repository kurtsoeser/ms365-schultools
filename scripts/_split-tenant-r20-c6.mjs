/**
 * R20/c6: split remaining tenant-settings-ui* modules to ≤600 lines (move-only).
 * Run once from repo root: node scripts/_split-tenant-r20-c6.mjs
 */
import fs from 'fs';
import path from 'path';
import { spawnSync } from 'child_process';

const ROOT = path.resolve('e:/MS365schule/src/shared');

function readLines(file) {
    return fs.readFileSync(path.join(ROOT, file), 'utf8').split(/\r?\n/);
}

function writeFile(file, text) {
    const p = path.join(ROOT, file);
    fs.writeFileSync(p, text.endsWith('\n') ? text : text + '\n');
    const check = spawnSync(process.execPath, ['--check', p], { encoding: 'utf8' });
    const n = text.split(/\r?\n/).length;
    console.log(file, n, check.status === 0 ? 'OK' : 'FAIL');
    if (check.status !== 0) {
        console.error(check.stderr || check.stdout);
        process.exit(1);
    }
}

function unindentBody(lines, spaces = 8) {
    return lines
        .map((l) => {
            if (!l.length) return l;
            let i = 0;
            while (i < spaces && l[i] === ' ') i++;
            return l.slice(i);
        })
        .join('\n');
}

function wrapInstall(file, exportName, bodyLines, indent = 8) {
    const body = unindentBody(bodyLines, indent);
    return `/**
 * ${file}
 * @param {Record<string, any>} ctx
 */
export function ${exportName}(ctx) {
${body
    .split('\n')
    .map((l) => (l.length ? '    ' + l : l))
    .join('\n')}
}
`;
}

// ─── bind-events ─────────────────────────────────────────────────────────────
{
    const lines = readLines('tenant-settings-ui-bind-events.js');
    // 1-based inclusive ranges of body inside installTenantBindEvents (file line numbers)
    const chunks = [
        {
            file: 'tenant-settings-ui-bind-events-writeback.js',
            exportName: 'installTenantBindEventsWriteback',
            start: 6,
            end: 96
        },
        {
            file: 'tenant-settings-ui-bind-events-files.js',
            exportName: 'installTenantBindEventsFiles',
            start: 98,
            end: 604
        },
        {
            file: 'tenant-settings-ui-bind-events-catalog.js',
            exportName: 'installTenantBindEventsCatalog',
            start: 605,
            end: 815
        },
        {
            file: 'tenant-settings-ui-bind-events-people.js',
            exportName: 'installTenantBindEventsPeople',
            start: 816,
            end: 1236
        },
        {
            file: 'tenant-settings-ui-bind-events-groups.js',
            exportName: 'installTenantBindEventsGroups',
            start: 1237,
            end: 1633
        },
        {
            file: 'tenant-settings-ui-bind-events-inventory.js',
            exportName: 'installTenantBindEventsInventory',
            start: 1634,
            end: 2209
        },
        {
            file: 'tenant-settings-ui-bind-events-webuntis.js',
            exportName: 'installTenantBindEventsWebuntis',
            start: 2211,
            end: 2440
        }
    ];

    for (const c of chunks) {
        const body = lines.slice(c.start - 1, c.end);
        writeFile(c.file, wrapInstall(c.file, c.exportName, body, 8));
    }

    const barrel = `/**
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
`;
    writeFile('tenant-settings-ui-bind-events.js', barrel);
}

// ─── school-groups ───────────────────────────────────────────────────────────
{
    const lines = readLines('tenant-settings-ui-school-groups.js');
    const chunks = [
        {
            file: 'tenant-settings-ui-school-groups-match.js',
            exportName: 'installTenantSchoolGroupsMatch',
            start: 6,
            end: 282
        },
        {
            file: 'tenant-settings-ui-school-groups-naming.js',
            exportName: 'installTenantSchoolGroupsNaming',
            start: 283,
            end: 591
        },
        {
            file: 'tenant-settings-ui-school-groups-lehrer.js',
            exportName: 'installTenantSchoolGroupsLehrer',
            start: 593,
            end: 861
        }
    ];
    for (const c of chunks) {
        writeFile(c.file, wrapInstall(c.file, c.exportName, lines.slice(c.start - 1, c.end), 8));
    }
    writeFile(
        'tenant-settings-ui-school-groups.js',
        `/**
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
`
    );
}

// ─── directory ───────────────────────────────────────────────────────────────
{
    const lines = readLines('tenant-settings-ui-directory.js');
    // indent is 4 spaces for this file
    const chunks = [
        {
            file: 'tenant-settings-ui-directory-users.js',
            exportName: 'installTenantDirectoryUsers',
            start: 6,
            end: 271
        },
        {
            file: 'tenant-settings-ui-directory-classes.js',
            exportName: 'installTenantDirectoryClasses',
            start: 272,
            end: 598
        },
        {
            file: 'tenant-settings-ui-directory-arges.js',
            exportName: 'installTenantDirectoryArges',
            start: 600,
            end: 792
        }
    ];
    for (const c of chunks) {
        writeFile(c.file, wrapInstall(c.file, c.exportName, lines.slice(c.start - 1, c.end), 4));
    }
    writeFile(
        'tenant-settings-ui-directory.js',
        `/**
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
`
    );
}

// ─── render ──────────────────────────────────────────────────────────────────
{
    const lines = readLines('tenant-settings-ui-render.js');
    const chunks = [
        {
            file: 'tenant-settings-ui-render-storage.js',
            exportName: 'installTenantRenderStorageCore',
            start: 6,
            end: 222
        },
        {
            file: 'tenant-settings-ui-render-import-route.js',
            exportName: 'installTenantRenderImportRoute',
            start: 223,
            end: 475
        },
        {
            file: 'tenant-settings-ui-render-import-rows.js',
            exportName: 'installTenantRenderImportRows',
            start: 476,
            end: 769
        }
    ];
    for (const c of chunks) {
        writeFile(c.file, wrapInstall(c.file, c.exportName, lines.slice(c.start - 1, c.end), 8));
    }
    writeFile(
        'tenant-settings-ui-render.js',
        `/**
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
`
    );
}

// ─── teachers ────────────────────────────────────────────────────────────────
{
    const lines = readLines('tenant-settings-ui-teachers.js');
    const chunks = [
        {
            file: 'tenant-settings-ui-teachers-admin.js',
            exportName: 'installTenantTeachersAdmin',
            start: 6,
            end: 413
        },
        {
            file: 'tenant-settings-ui-teachers-table.js',
            exportName: 'installTenantTeachersTable',
            start: 414,
            end: 764
        }
    ];
    for (const c of chunks) {
        writeFile(c.file, wrapInstall(c.file, c.exportName, lines.slice(c.start - 1, c.end), 8));
    }
    writeFile(
        'tenant-settings-ui-teachers.js',
        `/**
 * tenant-settings-ui-teachers.js — thin orchestrator (R20/c6).
 * @param {Record<string, any>} ctx
 */
import { installTenantTeachersAdmin } from './tenant-settings-ui-teachers-admin.js';
import { installTenantTeachersTable } from './tenant-settings-ui-teachers-table.js';

export function installTenantTeachersPanel(ctx) {
    installTenantTeachersAdmin(ctx);
    installTenantTeachersTable(ctx);
}
`
    );
}

console.log('panel splits done');
