/**
 * Lädt .env.local (einfaches KEY=VALUE) und CLI-/Default-Konfiguration.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
export const SCRIPT_ROOT = path.resolve(__dirname, '..');
export const REPO_ROOT = path.resolve(SCRIPT_ROOT, '..', '..');

/**
 * @param {string} filePath
 */
export function loadEnvFile(filePath) {
    if (!fs.existsSync(filePath)) return;
    const text = fs.readFileSync(filePath, 'utf8');
    for (const rawLine of text.split(/\r?\n/)) {
        const line = rawLine.trim();
        if (!line || line.startsWith('#')) continue;
        const eq = line.indexOf('=');
        if (eq <= 0) continue;
        const key = line.slice(0, eq).trim();
        let val = line.slice(eq + 1).trim();
        if (
            (val.startsWith('"') && val.endsWith('"')) ||
            (val.startsWith("'") && val.endsWith("'"))
        ) {
            val = val.slice(1, -1);
        }
        if (!(key in process.env) || process.env[key] === '') {
            process.env[key] = val;
        }
    }
}

export function loadDefaultEnv() {
    loadEnvFile(path.join(SCRIPT_ROOT, '.env.local'));
    loadEnvFile(path.join(SCRIPT_ROOT, '.env'));
}

/**
 * Client-ID aus ms365-config.js lesen (öffentlich, kein Secret).
 */
export function readRepoClientId() {
    const cfgPath = path.join(REPO_ROOT, 'ms365-config.js');
    if (!fs.existsSync(cfgPath)) return '';
    const content = fs.readFileSync(cfgPath, 'utf8');
    const m = content.match(/clientId:\s*['"]([^'"]+)['"]/);
    return m ? String(m[1] || '').trim() : '';
}

/**
 * @returns {{
 *   nextcloudBaseUrl: string,
 *   nextcloudUser: string,
 *   nextcloudAppPassword: string,
 *   graphTenantId: string,
 *   graphClientId: string,
 *   destRootFolder: string,
 *   stateDir: string
 * }}
 */
export function getConfig() {
    loadDefaultEnv();
    const nextcloudBaseUrl = String(
        process.env.NEXTCLOUD_URL || 'https://www.dsrom-intranet.de'
    )
        .trim()
        .replace(/\/+$/, '');
    const nextcloudUser = String(process.env.NEXTCLOUD_USER || '').trim();
    const nextcloudAppPassword = String(process.env.NEXTCLOUD_APP_PASSWORD || '').trim();
    const graphTenantId = String(process.env.GRAPH_TENANT_ID || 'organizations').trim();
    const graphClientId = String(
        process.env.GRAPH_CLIENT_ID || readRepoClientId() || ''
    ).trim();
    const destRootFolder = String(
        process.env.DEST_ROOT_FOLDER || 'Import aus Nextcloud'
    ).trim();
    const stateDir = String(
        process.env.MIGRATE_STATE_DIR || path.join(SCRIPT_ROOT, '.state')
    ).trim();

    return {
        nextcloudBaseUrl,
        nextcloudUser,
        nextcloudAppPassword,
        graphTenantId,
        graphClientId,
        destRootFolder,
        stateDir
    };
}

/**
 * @param {ReturnType<typeof getConfig>} cfg
 * @param {{ requirePassword?: boolean }} [opts]
 */
export function assertNextcloudAuth(cfg, opts = {}) {
    if (!cfg.nextcloudUser) {
        throw new Error(
            'NEXTCLOUD_USER fehlt. In scripts/nextcloud-teams-migrate/.env.local setzen.'
        );
    }
    if (opts.requirePassword !== false && !cfg.nextcloudAppPassword) {
        throw new Error(
            'NEXTCLOUD_APP_PASSWORD fehlt (App-Passwort aus Nextcloud Personal → Security).'
        );
    }
}

/**
 * @param {string} mappingPath
 */
export function loadMapping(mappingPath) {
    const abs = path.isAbsolute(mappingPath)
        ? mappingPath
        : path.resolve(process.cwd(), mappingPath);
    if (!fs.existsSync(abs)) {
        throw new Error('Mapping-Datei nicht gefunden: ' + abs);
    }
    const raw = JSON.parse(fs.readFileSync(abs, 'utf8'));
    const entries = Array.isArray(raw) ? raw : Array.isArray(raw.entries) ? raw.entries : [raw];
    return entries.map((e, i) => {
        const nextcloudPath = String(e.nextcloudPath || e.sourcePath || '').trim();
        const teamName = String(e.teamName || e.destTeamName || '').trim();
        const groupId = String(e.groupId || e.destGroupId || '').trim();
        if (!nextcloudPath) throw new Error(`Mapping[${i}]: nextcloudPath fehlt`);
        if (!teamName && !groupId) {
            throw new Error(`Mapping[${i}]: teamName oder groupId erforderlich`);
        }
        return {
            id: String(e.id || `entry-${i + 1}`).trim(),
            nextcloudPath: nextcloudPath.replace(/^\/+|\/+$/g, ''),
            teamName,
            groupId,
            destSubfolder: String(e.destSubfolder || '').trim()
        };
    });
}
