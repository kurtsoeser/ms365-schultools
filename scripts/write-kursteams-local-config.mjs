/**
 * Schreibt dist/ms365-config.local.js aus GitHub Secrets / Env (Deploy).
 *
 * Erlaubt:
 * - LICENSE_API_BASE_URL
 * - KURSTEAMS_API_BASE_URL (optional)
 * - ELECTION_API_BASE_URL (optional)
 *
 * Function Keys werden bewusst NICHT geschrieben (a9 / R10).
 * Auth läuft über Entra-Tokens (License.Access / Kursteams.Create).
 *
 * URLs mit „-dev“: Warnung. Harter Abbruch nur wenn
 * MS365_REQUIRE_PROD_URLS=true (sobald echte Prod-Function-Apps existieren).
 * Aktuell gibt es in Azure nur *-dev-Apps – Secrets dürfen deshalb vorerst
 * noch auf -dev zeigen, bis Prod-Apps angelegt sind.
 */
import fs from 'node:fs';
import path from 'node:path';

const licenseBaseUrl = String(process.env.LICENSE_API_BASE_URL || '').trim();
const kursteamsBaseUrl = String(process.env.KURSTEAMS_API_BASE_URL || '').trim();
const electionBaseUrl = String(process.env.ELECTION_API_BASE_URL || '').trim();

const requireProdUrls =
    String(process.env.MS365_REQUIRE_PROD_URLS || '').toLowerCase() === 'true' ||
    String(process.env.MS365_REQUIRE_PROD_URLS || '') === '1';

const distRoot = path.resolve(process.cwd(), 'dist');
const outPath = path.join(distRoot, 'ms365-config.local.js');

/**
 * @param {string} label
 * @param {string} url
 */
function checkDevUrl(label, url) {
    if (!url) return;
    if (!/-dev/i.test(url)) return;
    const msg =
        `write-kursteams-local-config: ${label} enthält "-dev" (${url}). ` +
        'Sobald Prod-Function-Apps ohne -dev existieren: Secret umstellen und MS365_REQUIRE_PROD_URLS=true setzen.';
    if (requireProdUrls) {
        console.error(msg);
        process.exit(1);
    }
    console.warn(msg);
}

checkDevUrl('LICENSE_API_BASE_URL', licenseBaseUrl);
checkDevUrl('KURSTEAMS_API_BASE_URL', kursteamsBaseUrl);
checkDevUrl('ELECTION_API_BASE_URL', electionBaseUrl);

if (!licenseBaseUrl && !kursteamsBaseUrl && !electionBaseUrl) {
    console.log(
        'write-kursteams-local-config: keine API-Base-URLs gesetzt – dist/ms365-config.local.js wird nicht erzeugt.'
    );
    process.exit(0);
}

if (!fs.existsSync(distRoot)) {
    console.error('write-kursteams-local-config: dist/ fehlt – zuerst npm run build.');
    process.exit(1);
}

/** @type {string[]} */
const blocks = [];

if (licenseBaseUrl) {
    blocks.push(
        '    MS365_LICENSE_API: {\n' +
            `        baseUrl: ${JSON.stringify(licenseBaseUrl)}\n` +
            '    }'
    );
}

if (kursteamsBaseUrl) {
    blocks.push(
        '    MS365_KURSTEAMS_API: {\n' +
            `        baseUrl: ${JSON.stringify(kursteamsBaseUrl)}\n` +
            '    }'
    );
}

if (electionBaseUrl) {
    blocks.push(
        '    MS365_ELECTION_API: {\n' +
            `        baseUrl: ${JSON.stringify(electionBaseUrl)}\n` +
            '    }'
    );
}

const content =
    '/** Beim Deploy aus GitHub Secrets erzeugt – nicht ins Repo committen. Kein Function Key. */\n' +
    'window.MS365_CONFIG_LOCAL = {\n' +
    blocks.join(',\n') +
    '\n};\n';

fs.writeFileSync(outPath, content, 'utf8');
console.log('write-kursteams-local-config: dist/ms365-config.local.js erstellt (ohne Function Key).');
