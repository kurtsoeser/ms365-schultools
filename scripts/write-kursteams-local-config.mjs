/**
 * Schreibt dist/ms365-config.local.js aus GitHub Secrets / Env.
 * - KURSTEAMS_FUNCTION_KEY
 * - LICENSE_API_BASE_URL (z. B. https://func-….azurewebsites.net/api/license)
 * - LICENSE_FUNCTION_KEY (optional)
 */
import fs from 'node:fs';
import path from 'node:path';

const kursteamsKey = String(process.env.KURSTEAMS_FUNCTION_KEY || '').trim();
const licenseBaseUrl = String(process.env.LICENSE_API_BASE_URL || '').trim();
const licenseKey = String(process.env.LICENSE_FUNCTION_KEY || '').trim();

const distRoot = path.resolve(process.cwd(), 'dist');
const outPath = path.join(distRoot, 'ms365-config.local.js');

if (!kursteamsKey && !licenseBaseUrl && !licenseKey) {
    console.log(
        'write-kursteams-local-config: keine Secrets gesetzt – dist/ms365-config.local.js wird nicht erzeugt.'
    );
    process.exit(0);
}

if (!fs.existsSync(distRoot)) {
    console.error('write-kursteams-local-config: dist/ fehlt – zuerst npm run build.');
    process.exit(1);
}

/** @type {string[]} */
const blocks = [];

if (kursteamsKey) {
    blocks.push(
        '    MS365_KURSTEAMS_API: {\n' +
            `        functionKey: ${JSON.stringify(kursteamsKey)}\n` +
            '    }'
    );
}

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
