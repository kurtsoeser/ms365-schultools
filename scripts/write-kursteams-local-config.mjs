/**
 * Schreibt dist/ms365-config.local.js aus der Umgebungsvariable KURSTEAMS_FUNCTION_KEY.
 * Wird in GitHub Actions beim Pages-Deploy ausgenutzt (Secret, nicht im Git-Repo).
 *
 * Lokal: ms365-config.local.js manuell anlegen oder hier KURSTEAMS_FUNCTION_KEY setzen.
 */
import fs from 'node:fs';
import path from 'node:path';

const key = String(process.env.KURSTEAMS_FUNCTION_KEY || '').trim();
const distRoot = path.resolve(process.cwd(), 'dist');
const outPath = path.join(distRoot, 'ms365-config.local.js');

if (!key) {
    console.log(
        'write-kursteams-local-config: KURSTEAMS_FUNCTION_KEY nicht gesetzt – dist/ms365-config.local.js wird nicht erzeugt.'
    );
    process.exit(0);
}

if (!fs.existsSync(distRoot)) {
    console.error('write-kursteams-local-config: dist/ fehlt – zuerst npm run build.');
    process.exit(1);
}

const content =
    '/** Beim Deploy aus GitHub Secret erzeugt – nicht ins Repo committen. */\n' +
    'window.MS365_CONFIG_LOCAL = {\n' +
    '    MS365_KURSTEAMS_API: {\n' +
    `        functionKey: ${JSON.stringify(key)}\n` +
    '    }\n' +
    '};\n';

fs.writeFileSync(outPath, content, 'utf8');
console.log('write-kursteams-local-config: dist/ms365-config.local.js erstellt.');
