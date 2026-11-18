/**
 * Einmalig: SharePoint-Spalte LicenseEntitlements anlegen.
 * Nutzt local.settings.json (nicht committen).
 *
 *   node scripts/ensure-entitlements-column.mjs
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
const { getOperatorGraphToken } = require('../src/lib/msal-app-only.js');
const {
    ensureLicenseListSchema,
    resolveSiteId,
    resolveListId,
    listLicenses
} = require('../src/lib/sharepoint-license.js');

async function main() {
    const cfg = getConfig();
    console.log('Site:', cfg.siteWebUrl);
    console.log('Liste:', cfg.listDisplayName);

    const token = await getOperatorGraphToken();
    const site = await resolveSiteId(cfg.siteWebUrl, token);
    const list = await resolveListId(site.id, cfg.listDisplayName, token);
    const schema = await ensureLicenseListSchema(site.id, list.id, token);
    console.log(
        schema.entitlementsColumn.created
            ? 'Spalte LicenseEntitlements neu angelegt.'
            : 'Spalte LicenseEntitlements war bereits vorhanden.'
    );
    console.log(
        schema.itContactsColumn && schema.itContactsColumn.created
            ? 'Spalte ItContacts neu angelegt.'
            : 'Spalte ItContacts war bereits vorhanden.'
    );
    console.log(
        schema.toolSiteUrlColumn && schema.toolSiteUrlColumn.created
            ? 'Spalte ToolSiteUrl neu angelegt.'
            : 'Spalte ToolSiteUrl war bereits vorhanden.'
    );
    console.log(
        'Stammdaten-Spalten: neu',
        (schema.stammdaten.created || []).length,
        '| bereits',
        (schema.stammdaten.existing || []).length
    );
    if (schema.stammdaten.created && schema.stammdaten.created.length) {
        console.log('Neu:', schema.stammdaten.created.join(', '));
    }

    const listed = await listLicenses();
    console.log('Schulen geladen:', (listed.schools || []).length);
    const withJson = (listed.schools || []).filter((s) => s.entitlementsDoc).length;
    console.log('Mit Modul-JSON:', withJson, '| Legacy ohne JSON:', (listed.schools || []).length - withJson);
}

main().catch((e) => {
    console.error('Fehler:', e && e.message ? e.message : e);
    process.exit(1);
});
