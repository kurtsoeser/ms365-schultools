/**
 * Schreibt dist/ms365-config.local.js aus GitHub Secrets / Env.
 * - LICENSE_API_BASE_URL (z. B. https://func-….azurewebsites.net/api/license)
 * - LICENSE_FUNCTION_KEY (optional)
 *
 * KURSTEAMS_FUNCTION_KEY wird absichtlich ignoriert. Das Kursteams-Backend
 * verlangt ein Benutzer-Token und darf keinen Function Key in die Seite schreiben.
 */
import fs from 'node:fs';
import path from 'node:path';

const licenseBaseUrl = String(process.env.LICENSE_API_BASE_URL || '').trim();
const licenseKey = String(process.env.LICENSE_FUNCTION_KEY || '').trim();

const distRoot = path.resolve(process.cwd(), 'dist');
const outPath = path.join(distRoot, 'ms365-config.local.js');

if (!licenseBaseUrl && !licenseKey) {
    console.log(
        'write-kursteams-local-config: keine Lizenz-Secrets gesetzt – dist/ms365-config.local.js wird nicht erzeugt.'
    );
    process.exit(0);
}

if (!fs.existsSync(distRoot)) {
    console.error('write-kursteams-local-config: dist/ fehlt – zuerst npm run build.');
    process.exit(1);
}

/** @type {string[]} */
const blocks = [];

if (licenseBaseUrl || licenseKey) {
    blocks.push(
        '    MS365_LICENSE_API: {\n' +
            `        baseUrl: ${JSON.stringify(licenseBaseUrl)},\n` +
            `        functionKey: ${JSON.stringify(licenseKey)}\n` +
            '    }'
    );
}

const content =
    '/** Beim Deploy aus GitHub Secrets erzeugt – nicht ins Repo committen. */\n' +
    'window.MS365_CONFIG_LOCAL = {\n' +
    blocks.join(',\n') +
    '\n};\n';

fs.writeFileSync(outPath, content, 'utf8');
console.log('write-kursteams-local-config: dist/ms365-config.local.js erstellt.');
