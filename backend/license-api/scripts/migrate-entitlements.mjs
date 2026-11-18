/**
 * Schulen ohne Modul-JSON befüllen (legacy-full | starter).
 *
 *   node scripts/migrate-entitlements.mjs
 *   node scripts/migrate-entitlements.mjs --mode=starter
 *   node scripts/migrate-entitlements.mjs --dry-run
 */
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');
const require = createRequire(import.meta.url);

function loadLocalSettings() {
    const path = join(root, 'local.settings.json');
    const raw = JSON.parse(readFileSync(path, 'utf8'));
    const values = (raw && raw.Values) || {};
    Object.keys(values).forEach((k) => {
        if (process.env[k] == null || String(process.env[k]).trim() === '') {
            process.env[k] = String(values[k]);
        }
    });
}

loadLocalSettings();

const args = process.argv.slice(2);
const dryRun = args.includes('--dry-run');
const modeArg = args.find((a) => a.startsWith('--mode='));
const mode = modeArg ? modeArg.slice('--mode='.length) : 'legacy-full';

const { migrateSchoolEntitlements } = require('../src/lib/sharepoint-license.js');

async function main() {
    console.log('Mode:', mode, dryRun ? '(dry-run)' : '');
    const result = await migrateSchoolEntitlements({ mode, dryRun });
    console.log(JSON.stringify(result, null, 2));
}

main().catch((e) => {
    console.error('Fehler:', e && e.message ? e.message : e);
    process.exit(1);
});
