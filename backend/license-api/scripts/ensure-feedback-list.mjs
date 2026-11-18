/**
 * Einmalig: Feedback-Liste + Spalten auf der Betreiber-Site anlegen.
 * Nutzt local.settings.json (nicht committen).
 *
 *   node scripts/ensure-feedback-list.mjs
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

const { getConfig } = require('../src/lib/config.js');
const { ensureFeedbackSchema } = require('../src/lib/sharepoint-feedback.js');

async function main() {
    const cfg = getConfig();
    console.log('Site:', cfg.siteWebUrl);
    console.log('Feedback-Liste:', cfg.feedbackListDisplayName);

    const schema = await ensureFeedbackSchema();
    console.log('Liste-ID:', schema.listId);
    console.log(
        'Spalten neu:',
        (schema.columns.created || []).length ? schema.columns.created.join(', ') : '(keine)'
    );
    console.log(
        'Spalten vorhanden:',
        (schema.columns.existing || []).length ? schema.columns.existing.join(', ') : '(keine)'
    );
    console.log('Fertig.');
}

main().catch((e) => {
    console.error('Fehler:', e && e.message ? e.message : e);
    process.exit(1);
});
