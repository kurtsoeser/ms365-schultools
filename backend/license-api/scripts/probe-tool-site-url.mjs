/**
 * Diagnose: ToolSiteUrl-Spalte + Schreibtest auf der konfigurierten Lizenzliste.
 *
 *   node scripts/probe-tool-site-url.mjs
 *   node scripts/probe-tool-site-url.mjs --write "https://schule.sharepoint.com/sites/MS365-Schultools"
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
    loadAllListItems,
    updateLicense,
    mapListItem,
    pickToolSiteUrlFromFields,
    findToolSiteUrlColumnMeta,
    TOOL_SITE_URL_FIELD
} = require('../src/lib/sharepoint-license.js');

async function loadColumns(siteId, listId, token) {
    const GRAPH = 'https://graph.microsoft.com/v1.0';
    const res = await fetch(
        GRAPH +
            '/sites/' +
            encodeURIComponent(siteId) +
            '/lists/' +
            encodeURIComponent(listId) +
            '/columns?$top=200',
        { headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' } }
    );
    const data = await res.json();
    if (!res.ok) throw new Error(JSON.stringify(data));
    return data.value || [];
}

async function main() {
    const cfg = getConfig();
    const writeUrl = (() => {
        const i = process.argv.indexOf('--write');
        return i >= 0 ? String(process.argv[i + 1] || '').trim() : '';
    })();

    console.log('Site:', cfg.siteWebUrl);
    console.log('Liste:', cfg.listDisplayName);

    const token = await getOperatorGraphToken();
    const site = await resolveSiteId(cfg.siteWebUrl, token);
    const list = await resolveListId(site.id, cfg.listDisplayName, token);
    console.log('Site-ID:', site.id);
    console.log('Listen-ID:', list.id, '| Anzeigename:', list.displayName);

    const schema = await ensureLicenseListSchema(site.id, list.id, token);
    console.log('toolSiteUrlColumn:', JSON.stringify(schema.toolSiteUrlColumn));

    const cols = await loadColumns(site.id, list.id, token);
    const toolish = cols.filter((c) => {
        const n = String((c && c.name) || '');
        const d = String((c && c.displayName) || '');
        return /tool|sharepoint|site.?url/i.test(n + ' ' + d);
    });
    console.log(
        'Passende Spalten:',
        toolish.map((c) => ({
            name: c.name,
            displayName: c.displayName,
            hyperlink: !!c.hyperlinkOrPicture,
            text: !!c.text
        }))
    );
    console.log('findToolSiteUrlColumnMeta:', findToolSiteUrlColumnMeta(cols));

    const items = await loadAllListItems(site.id, list.id, token);
    console.log('Einträge:', items.length);
    const sample = items.slice(0, 5).map((it) => {
        const f = (it && it.fields) || {};
        return {
            id: it.id,
            title: f.Title,
            toolDirect: f[TOOL_SITE_URL_FIELD],
            toolPicked: pickToolSiteUrlFromFields(f)
        };
    });
    console.log('Sample:', sample);

    if (!writeUrl) {
        console.log('Hinweis: mit --write "https://…" einmalig auf den ersten Eintrag schreiben.');
        return;
    }
    const target = items.find((it) => /ms365\s*schule/i.test(String((it.fields || {}).Title || ''))) || items[0];
    if (!target) throw new Error('Kein Listeneintrag gefunden.');
    console.log('Schreibe auf ID', target.id, target.fields && target.fields.Title);
    const updated = await updateLicense(String(target.id), {
        toolSiteUrl: writeUrl,
        notes: String((target.fields && target.fields.Notes) || '')
    });
    console.log('updateLicense.toolSiteUrl =', updated.toolSiteUrl);
    const mapped = mapListItem(
        (
            await (
                await fetch(
                    'https://graph.microsoft.com/v1.0/sites/' +
                        encodeURIComponent(site.id) +
                        '/lists/' +
                        encodeURIComponent(list.id) +
                        '/items/' +
                        encodeURIComponent(target.id) +
                        '?$expand=fields',
                    { headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' } }
                )
            ).json()
        ),
        []
    );
    console.log('Re-Read toolSiteUrl =', mapped.toolSiteUrl);
}

main().catch((e) => {
    console.error('Fehler:', e && e.message ? e.message : e);
    process.exit(1);
});
