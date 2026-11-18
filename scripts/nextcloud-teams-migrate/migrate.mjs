#!/usr/bin/env node
/**
 * Nextcloud (WebDAV) → Microsoft Teams Dateien (Graph).
 *
 * Befehle:
 *   node migrate.mjs check
 *   node migrate.mjs inventory --path "5. Gymnasium (5-12)/1. Fächer/Physik"
 *   node migrate.mjs dry-run --mapping mapping.physik.json
 *   node migrate.mjs migrate --mapping mapping.physik.json [--limit 5] [--max-bytes 52428800]
 */
import path from 'node:path';
import {
    SCRIPT_ROOT,
    getConfig,
    assertNextcloudAuth,
    loadMapping
} from './lib/config.mjs';
import {
    propfind,
    inventoryTree,
    downloadFileStream,
    formatBytes,
    davFilesRoot
} from './lib/webdav.mjs';
import {
    acquireDeviceCodeToken,
    getGroupDrive,
    resolveGroupIdByName,
    ensureFolderPath,
    uploadFile,
    withRetry
} from './lib/graph.mjs';
import { stateKey, loadState, markDone } from './lib/state.mjs';

/**
 * @param {string[]} argv
 */
function parseArgs(argv) {
    const args = argv.slice(2);
    const cmd = args[0] || 'help';
    /** @type {Record<string, string | boolean>} */
    const flags = {};
    for (let i = 1; i < args.length; i++) {
        const a = args[i];
        if (a.startsWith('--')) {
            const key = a.slice(2);
            const next = args[i + 1];
            if (!next || next.startsWith('--')) {
                flags[key] = true;
            } else {
                flags[key] = next;
                i++;
            }
        }
    }
    return { cmd, flags };
}

function printHelp() {
    console.log(`
Nextcloud → Teams Migration

  node migrate.mjs check
  node migrate.mjs inventory --path "<relativer Ordner unter Dateien>"
  node migrate.mjs dry-run --mapping mapping.physik.json
  node migrate.mjs migrate --mapping mapping.physik.json [--limit N] [--max-bytes B]

Konfiguration: scripts/nextcloud-teams-migrate/.env.local
  (Vorlage: .env.example)

WebDAV-Root:
  {NEXTCLOUD_URL}/remote.php/dav/files/{NEXTCLOUD_USER}/
`);
}

/**
 * Relativpfad einer Datei innerhalb des Quellordners.
 * @param {string} baseRel
 * @param {string} fileRel
 */
function relativeInside(baseRel, fileRel) {
    const base = String(baseRel || '').replace(/^\/+|\/+$/g, '');
    const file = String(fileRel || '').replace(/^\/+|\/+$/g, '');
    if (!base) return file;
    if (file === base) return '';
    if (file.startsWith(base + '/')) return file.slice(base.length + 1);
    return file;
}

async function cmdCheck() {
    const cfg = getConfig();
    assertNextcloudAuth(cfg);
    console.log('Nextcloud URL:', cfg.nextcloudBaseUrl);
    console.log('User:', cfg.nextcloudUser);
    console.log('WebDAV:', davFilesRoot(cfg.nextcloudBaseUrl, cfg.nextcloudUser));
    console.log('PROPFIND Depth:0 …');
    const items = await propfind({
        baseUrl: cfg.nextcloudBaseUrl,
        user: cfg.nextcloudUser,
        appPassword: cfg.nextcloudAppPassword,
        relPath: '',
        depth: '0'
    });
    console.log('OK – Root erreichbar.', items[0]?.displayName || items[0]?.relPath || '(root)');
    console.log('Graph Client-ID:', cfg.graphClientId || '(fehlt)');
    console.log('Graph Tenant:', cfg.graphTenantId);
    console.log('Ziel-Root-Ordner:', cfg.destRootFolder);
}

/**
 * @param {Record<string, string | boolean>} flags
 */
async function cmdInventory(flags) {
    const cfg = getConfig();
    assertNextcloudAuth(cfg);
    const relPath = String(flags.path || flags.p || '').trim();
    if (!relPath) {
        throw new Error('Bitte --path „…/Physik“ angeben (relativ zum Nextcloud-Benutzerroot).');
    }
    console.log('Inventar:', relPath);
    const inv = await inventoryTree({
        baseUrl: cfg.nextcloudBaseUrl,
        user: cfg.nextcloudUser,
        appPassword: cfg.nextcloudAppPassword,
        relPath
    });
    console.log('Ordner:', inv.folderCount);
    console.log('Dateien:', inv.fileCount);
    console.log('Größe:', formatBytes(inv.totalBytes), `(${inv.totalBytes} Bytes)`);
    const sample = inv.files.slice(0, 15);
    if (sample.length) {
        console.log('Beispiel-Dateien:');
        for (const f of sample) {
            console.log(`  - ${f.relPath} (${formatBytes(f.size)})`);
        }
        if (inv.files.length > sample.length) {
            console.log(`  … und ${inv.files.length - sample.length} weitere`);
        }
    }
    return inv;
}

/**
 * @param {Record<string, string | boolean>} flags
 * @param {{ doUpload: boolean }} mode
 */
async function cmdMigrate(flags, mode) {
    const cfg = getConfig();
    assertNextcloudAuth(cfg);
    const mappingPath = String(flags.mapping || flags.m || '').trim();
    if (!mappingPath) throw new Error('--mapping fehlt (z. B. mapping.physik.json)');
    const mappingAbs = path.isAbsolute(mappingPath)
        ? mappingPath
        : path.resolve(SCRIPT_ROOT, mappingPath);
    const entries = loadMapping(mappingAbs);
    const limit = flags.limit != null ? Number(flags.limit) : null;
    const maxBytes = flags['max-bytes'] != null ? Number(flags['max-bytes']) : null;

    console.log(mode.doUpload ? '=== MIGRATE ===' : '=== DRY-RUN ===');
    console.log('Mapping:', mappingAbs);
    console.log('Einträge:', entries.length);
    if (limit) console.log('Limit Dateien:', limit);
    if (maxBytes) console.log('Max Bytes:', formatBytes(maxBytes));

    /** @type {string|null} */
    let token = null;
    if (mode.doUpload) {
        const auth = await acquireDeviceCodeToken(cfg.graphTenantId, cfg.graphClientId);
        token = auth.accessToken;
        console.log('Graph-Token OK.');
    }

    for (const entry of entries) {
        console.log('');
        console.log(`--- ${entry.id}: ${entry.nextcloudPath} → ${entry.teamName || entry.groupId} ---`);

        const inv = await inventoryTree({
            baseUrl: cfg.nextcloudBaseUrl,
            user: cfg.nextcloudUser,
            appPassword: cfg.nextcloudAppPassword,
            relPath: entry.nextcloudPath
        });
        console.log(
            `Quelle: ${inv.fileCount} Dateien, ${inv.folderCount} Ordner, ${formatBytes(inv.totalBytes)}`
        );

        let groupId = entry.groupId;
        /** @type {any} */
        let drive = null;
        if (mode.doUpload && token) {
            if (!groupId) {
                groupId = await resolveGroupIdByName(token, entry.teamName);
            }
            console.log('Ziel Group-ID:', groupId);
            drive = await getGroupDrive(token, groupId);
            console.log('Ziel Drive:', drive.name || drive.id);
        } else {
            console.log(
                'Ziel (geplant):',
                entry.teamName || groupId,
                '→ Ordner',
                JSON.stringify(
                    [cfg.destRootFolder, entry.destSubfolder || path.basename(entry.nextcloudPath)]
                        .filter(Boolean)
                        .join('/')
                )
            );
        }

        const key = stateKey(entry.id, entry.nextcloudPath, groupId || entry.teamName || 'pending');
        const state = mode.doUpload ? loadState(cfg.stateDir, key) : { done: {} };
        const already = Object.keys(state.done).length;
        if (already) console.log('Resume: bereits erledigt:', already);

        let uploaded = 0;
        let skipped = 0;
        let bytesPlanned = 0;
        let bytesDone = 0;
        let hitLimit = false;

        const destBaseParts = [cfg.destRootFolder];
        if (entry.destSubfolder) destBaseParts.push(entry.destSubfolder);
        else destBaseParts.push(path.basename(entry.nextcloudPath) || 'Import');
        const destBase = destBaseParts.filter(Boolean).join('/');

        for (const file of inv.files) {
            const relInside = relativeInside(entry.nextcloudPath, file.relPath);
            if (!relInside) continue;

            if (state.done[relInside]) {
                skipped++;
                continue;
            }
            if (limit != null && uploaded >= limit) {
                hitLimit = true;
                break;
            }
            if (maxBytes != null && bytesPlanned + (file.size || 0) > maxBytes && uploaded > 0) {
                hitLimit = true;
                break;
            }

            bytesPlanned += file.size || 0;
            const parentRel = relInside.includes('/')
                ? relInside.split('/').slice(0, -1).join('/')
                : '';
            const fileName = relInside.split('/').pop() || file.displayName || 'file';
            const destFolder = parentRel ? `${destBase}/${parentRel}` : destBase;

            if (!mode.doUpload) {
                console.log(`[dry] ${relInside} → ${destFolder}/${fileName} (${formatBytes(file.size)})`);
                uploaded++;
                continue;
            }

            process.stdout.write(
                `[${uploaded + skipped + 1}/${inv.fileCount}] ${relInside} (${formatBytes(file.size)}) … `
            );
            await withRetry(
                async () => {
                    if (!token || !drive) throw new Error('Kein Graph-Token/Drive');
                    const parent = await ensureFolderPath(token, drive.id, destFolder);
                    const dl = await downloadFileStream({
                        baseUrl: cfg.nextcloudBaseUrl,
                        user: cfg.nextcloudUser,
                        appPassword: cfg.nextcloudAppPassword,
                        relPath: file.relPath
                    });
                    await uploadFile(
                        token,
                        drive.id,
                        parent.id,
                        fileName,
                        dl.stream,
                        file.size || dl.contentLength
                    );
                },
                { label: relInside, retries: 4 }
            );
            markDone(cfg.stateDir, key, relInside, file.size || 0, state);
            bytesDone += file.size || 0;
            uploaded++;
            console.log('OK');
        }

        console.log(
            `Fertig ${entry.id}: neu=${uploaded}, übersprungen=${skipped}, bytes=${formatBytes(
                mode.doUpload ? bytesDone : bytesPlanned
            )}${hitLimit ? ' (Limit erreicht)' : ''}`
        );
    }
}

async function main() {
    const { cmd, flags } = parseArgs(process.argv);
    try {
        if (cmd === 'help' || cmd === '-h' || cmd === '--help') {
            printHelp();
            return;
        }
        if (cmd === 'check') {
            await cmdCheck();
            return;
        }
        if (cmd === 'inventory') {
            await cmdInventory(flags);
            return;
        }
        if (cmd === 'dry-run') {
            await cmdMigrate(flags, { doUpload: false });
            return;
        }
        if (cmd === 'migrate') {
            await cmdMigrate(flags, { doUpload: true });
            return;
        }
        printHelp();
        throw new Error('Unbekannter Befehl: ' + cmd);
    } catch (e) {
        console.error('');
        console.error('Fehler:', e && e.message ? e.message : e);
        process.exitCode = 1;
    }
}

main();
