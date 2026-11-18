/**
 * Phase 3: Workshop-Schulen aus Excel in die Betreiber-Lizenzliste einspielen.
 *
 *   node scripts/import-workshop-schools.mjs --dry-run
 *   node scripts/import-workshop-schools.mjs --apply
 *
 * Optional:
 *   --excel "C:\\…\\MS365_Schultools_Schulen_….xlsx"
 *   --valid-until 2026-12-24
 *   --update   vorhandene Schulen (Domain/Name) Stammdaten nachziehen
 */
import { readFileSync, existsSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');
const require = createRequire(import.meta.url);

const DEFAULT_EXCEL = resolve(
    process.env.USERPROFILE || process.env.HOME || '',
    'Downloads',
    'MS365_Schultools_Schulen_Workshopbloecke_A_B_2026-11-18.xlsx'
);
const DEFAULT_VALID_UNTIL = '2026-12-24';

function loadLocalSettings() {
    const path = join(root, 'local.settings.json');
    if (!existsSync(path)) {
        throw new Error('local.settings.json fehlt – App-Only-Zugang für SharePoint nötig.');
    }
    const raw = JSON.parse(readFileSync(path, 'utf8'));
    const values = (raw && raw.Values) || {};
    Object.keys(values).forEach((k) => {
        if (process.env[k] == null || String(process.env[k]).trim() === '') {
            process.env[k] = String(values[k]);
        }
    });
}

function parseArgs(argv) {
    const out = {
        dryRun: true,
        apply: false,
        update: false,
        excel: DEFAULT_EXCEL,
        validUntil: DEFAULT_VALID_UNTIL
    };
    for (let i = 0; i < argv.length; i++) {
        const a = argv[i];
        if (a === '--apply') {
            out.apply = true;
            out.dryRun = false;
        } else if (a === '--dry-run') {
            out.dryRun = true;
            out.apply = false;
        } else if (a === '--update') {
            out.update = true;
        } else if (a === '--excel' && argv[i + 1]) {
            out.excel = resolve(argv[++i]);
        } else if (a === '--valid-until' && argv[i + 1]) {
            out.validUntil = String(argv[++i]).trim();
        }
    }
    return out;
}

function readExcelRows(excelPath) {
    if (!existsSync(excelPath)) {
        throw new Error('Excel nicht gefunden: ' + excelPath);
    }
    const py = `
import json, sys
import openpyxl
path = sys.argv[1]
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb['Schulen_Import']
headers = [c.value for c in ws[7]]
rows = []
for row in ws.iter_rows(min_row=8, max_row=40, values_only=True):
    if not row or not row[1]:
        continue
    d = {}
    for i, h in enumerate(headers):
        if not h:
            continue
        v = row[i] if i < len(row) else None
        if v is None:
            d[str(h)] = ''
        else:
            d[str(h)] = str(v).strip()
    if d.get('Schulname'):
        rows.append(d)
print(json.dumps(rows, ensure_ascii=False))
`;
    const res = spawnSync('python', ['-c', py, excelPath], {
        encoding: 'utf8',
        maxBuffer: 4 * 1024 * 1024
    });
    if (res.status !== 0) {
        throw new Error(
            'Excel-Parse fehlgeschlagen: ' + (res.stderr || res.stdout || res.error || '')
        );
    }
    return JSON.parse(res.stdout || '[]');
}

function cell(row, ...keys) {
    for (let i = 0; i < keys.length; i++) {
        const k = keys[i];
        if (row[k] != null && String(row[k]).trim() !== '') return String(row[k]).trim();
    }
    // Fuzzy: Excel-Header können Encoding-Varianten haben
    const want = keys.map((k) => String(k).toLowerCase());
    for (const [k, v] of Object.entries(row || {})) {
        if (want.includes(String(k).toLowerCase()) && String(v || '').trim()) {
            return String(v).trim();
        }
    }
    return '';
}

function isOpenText(raw) {
    return /^offen\b/i.test(String(raw || '').trim());
}

function cleanRoleText(raw) {
    const s = String(raw || '').trim();
    if (!s || isOpenText(s)) return '';
    return s;
}

function splitEmails(raw) {
    return String(raw || '')
        .split(/[\n,;]+/)
        .map((s) => s.trim().toLowerCase())
        .filter((s) => s.includes('@'));
}

function splitDomains(raw) {
    return String(raw || '')
        .split(/[\n,;]+/)
        .map((s) => s.trim().toLowerCase().replace(/^https?:\/\//, '').replace(/\/.*$/, ''))
        .filter(Boolean);
}

function normName(s) {
    return String(s || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, ' ')
        .trim();
}

function mapRowToPayload(row, validUntil) {
    const schoolName = cell(row, 'Schulname');
    const domains = splitDomains(cell(row, 'Domain(s)', 'Domains'));
    const emails = splitEmails(cell(row, 'Kontakt allgemein'));
    const itRaw = cell(row, 'IT-/M365-Administration', 'IT/M365-Administration');
    const openBits = [];
    const dataStatus = cell(row, 'Datenstatus');
    const openPoints = cell(row, 'Offene Punkte / Prüfbedarf', 'Offene Punkte');
    if (isOpenText(itRaw)) openBits.push('IT-/M365-Administration: ' + itRaw);
    if (openPoints) openBits.push(openPoints);

    const tenantRaw = cell(row, 'Tenant-ID');
    const tenantId = !tenantRaw || isOpenText(tenantRaw) ? '' : tenantRaw;

    /** @type {Record<string, string>} */
    const extra = {
        Standort: cell(row, 'Standort'),
        Strasse: cell(row, 'Straße', 'Strasse'),
        PLZ: cell(row, 'PLZ'),
        Ort: cell(row, 'Ort'),
        Direktion: cell(row, 'Direktion'),
        Schuladministration: cell(row, 'Schuladministration'),
        ItM365Admin: cleanRoleText(itRaw),
        WeitereKontakte: cell(row, 'Weitere Kontakte / Rollen'),
        Telefon: cell(row, 'Telefon'),
        Schulnummer: cell(row, 'Schulnummer(n)', 'Schulnummer'),
        DomainHinweis: cell(row, 'Domain-Beleg / Hinweis'),
        Datenstatus: dataStatus || 'Workshop-Import · Tenant-ID offen',
        WorkshopTeilnehmende: cell(row, 'Workshop-Teilnehmende'),
        WorkshopQuelle: cell(row, 'Interne Workshopquelle'),
        OeffentlicheQuellen: cell(row, 'Öffentliche Quelle(n)', 'Oeffentliche Quelle(n)'),
        OffenePunkte: openBits.join('\n')
    };

    return {
        schoolName,
        tenantId,
        status: 'trial',
        validUntil,
        primaryDomain: domains[0] || '',
        additionalDomains: domains.slice(1).join('\n'),
        contactEmail: emails[0] || '',
        itContacts: emails.slice(1).join('\n'),
        notes: 'Workshop-Import A+B · Trial bis ' + validUntil,
        extra
    };
}

function findExisting(schools, payload) {
    const wantName = normName(payload.schoolName);
    const wantDomain = String(payload.primaryDomain || '')
        .trim()
        .toLowerCase();
    const list = Array.isArray(schools) ? schools : [];
    for (let i = 0; i < list.length; i++) {
        const s = list[i];
        if (wantName && normName(s.schoolName) === wantName) return s;
        const domains = Array.isArray(s.domains)
            ? s.domains.map((d) => String(d).toLowerCase())
            : [];
        if (wantDomain && domains.includes(wantDomain)) return s;
        if (
            wantDomain &&
            String(s.primaryDomain || '')
                .trim()
                .toLowerCase() === wantDomain
        ) {
            return s;
        }
    }
    return null;
}

async function main() {
    const args = parseArgs(process.argv.slice(2));
    loadLocalSettings();

    const {
        listLicenses,
        createLicense,
        updateLicense,
        ensureLicenseListSchema,
        resolveSiteId,
        resolveListId
    } = require('../src/lib/sharepoint-license.js');
    const { getOperatorGraphToken } = require('../src/lib/msal-app-only.js');
    const { getConfig } = require('../src/lib/config.js');

    console.log('Excel:', args.excel);
    console.log('Gültig bis:', args.validUntil);
    console.log('Modus:', args.apply ? 'APPLY' : 'DRY-RUN');

    const rows = readExcelRows(args.excel);
    console.log('Zeilen in Excel:', rows.length);
    if (!rows.length) throw new Error('Keine Schulzeilen gefunden.');

    const payloads = rows.map((r) => mapRowToPayload(r, args.validUntil));

    const token = await getOperatorGraphToken();
    const cfg = getConfig();
    const site = await resolveSiteId(cfg.siteWebUrl, token);
    const list = await resolveListId(site.id, cfg.listDisplayName, token);
    const schema = await ensureLicenseListSchema(site.id, list.id, token);
    console.log(
        'Schema Stammdaten neu:',
        (schema.stammdaten.created || []).length,
        '| vorhanden:',
        (schema.stammdaten.existing || []).length
    );

    const listed = await listLicenses();
    const existing = listed.schools || [];
    console.log('Schulen aktuell in SharePoint:', existing.length);

    let created = 0;
    let updated = 0;
    let skipped = 0;

    for (const payload of payloads) {
        const hit = findExisting(existing, payload);
        if (hit) {
            if (args.update) {
                console.log(
                    (args.apply ? 'UPDATE' : 'WOULD UPDATE') + ':',
                    payload.schoolName,
                    '→',
                    hit.id,
                    '|',
                    payload.primaryDomain
                );
                if (args.apply) {
                    await updateLicense(hit.id, {
                        schoolName: payload.schoolName,
                        status: payload.status,
                        validUntil: payload.validUntil,
                        primaryDomain: payload.primaryDomain,
                        additionalDomains: payload.additionalDomains,
                        contactEmail: payload.contactEmail,
                        itContacts: payload.itContacts,
                        notes: payload.notes,
                        extra: payload.extra
                        // TenantId bewusst nicht überschreiben, falls schon echte GUID
                    });
                    updated += 1;
                } else {
                    updated += 1;
                }
            } else {
                console.log('SKIP (existiert):', payload.schoolName, '→', hit.tenantId || hit.id);
                skipped += 1;
            }
            continue;
        }

        console.log(
            (args.apply ? 'CREATE' : 'WOULD CREATE') + ':',
            payload.schoolName,
            '|',
            payload.primaryDomain || '(ohne Domain)',
            '|',
            payload.contactEmail || '(ohne Mail)',
            '| Tenant leer → PENDING'
        );
        if (args.apply) {
            const createdSchool = await createLicense(payload);
            console.log(
                '  → id',
                createdSchool.id,
                '| tenant',
                createdSchool.tenantId,
                '| status',
                createdSchool.status
            );
            created += 1;
        } else {
            created += 1;
        }
    }

    console.log('---');
    console.log(
        args.apply ? 'Fertig.' : 'Dry-Run fertig (nichts geschrieben). Mit --apply ausführen.'
    );
    console.log({ created, updated, skipped, totalExcel: payloads.length });
}

main().catch((e) => {
    console.error('Fehler:', e && e.message ? e.message : e);
    process.exit(1);
});
