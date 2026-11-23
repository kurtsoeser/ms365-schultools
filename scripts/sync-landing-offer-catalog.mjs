/**
 * Erzeugt landing/offer-catalog.js aus LICENSE_PACKAGES + Dashboard-Tool-Namen.
 * So bleiben Bausteine auf der Homepage wartbar, wenn Tools/Pakete wachsen.
 *
 * Usage: node scripts/sync-landing-offer-catalog.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { LICENSE_PACKAGES } from '../src/shared/license-feature-catalog.js';
import {
    DASHBOARD_TOOL_BLURBS,
    DASHBOARD_TOOL_LABELS
} from '../src/shared/dashboard-audience-catalog.js';
import { SALES_ADDONS, SALES_BUNDLES } from '../src/shared/sales-bundles-catalog.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const outFile = path.join(root, 'landing', 'offer-catalog.js');

/** Reihenfolge für sichtbare Chips auf der Karte (Kunden-Alltag zuerst). */
const HIGHLIGHT_PRIORITY = [
    'kursteams',
    'schularbeiten-planer',
    'schulaktivitaeten-planer',
    'projektwochen',
    'schulsprecherwahl',
    'wahlfach-zuteilung',
    'wahlfach-anmeldung',
    'freistellung-planer',
    'lehrer-freistellung-planer',
    'elternbrief',
    'einsammler',
    'krankmeldung',
    'srdp-anmeldung',
    'termine-hub',
    'aufgaben',
    'technik-meldungen',
    'raumplan',
    'ausgaben',
    'diplomarbeiten',
    'onenote-verteilung',
    'personen-verwaltung',
    'lizenzverwaltung',
    'organisations-assistent',
    'schulstruktur-sync',
    'sharepoint-intranet-hub'
];

/** Extra Suchbegriffe (Kundenworte → Tool-IDs). */
const SEARCH_ALIASES = {
    schulwahl: ['schulsprecherwahl'],
    schulwahlen: ['schulsprecherwahl'],
    wahl: ['schulsprecherwahl', 'wahlfach-zuteilung', 'wahlfach-anmeldung'],
    schulsprecher: ['schulsprecherwahl'],
    matura: ['srdp-anmeldung', 'diplomarbeiten'],
    srdp: ['srdp-anmeldung'],
    eltern: ['elternbrief', 'einsammler', 'krankmeldung', 'elternsprechtag-bookings'],
    freistellung: ['freistellung-planer', 'lehrer-freistellung-planer'],
    kursteam: ['kursteams'],
    unterrichtsteam: ['kursteams'],
    onenote: ['onenote-verteilung'],
    ausgabe: ['ausgaben'],
    schlüssel: ['ausgaben'],
    technik: ['technik-meldungen'],
    raum: ['raumplan', 'raeume-ressourcen']
};

function isSetupTool(id) {
    const s = String(id || '');
    return (
        s.startsWith('sharepoint-liste-') ||
        s.startsWith('playbook-') ||
        s.endsWith('-setup') ||
        s.startsWith('pa-')
    );
}

function toolEntry(id) {
    const label = DASHBOARD_TOOL_LABELS[id] || id;
    const blurb = DASHBOARD_TOOL_BLURBS[id] || '';
    return {
        id,
        label,
        blurb,
        setup: isSetupTool(id)
    };
}

function pickHighlights(tools) {
    const primary = tools.filter((t) => !t.setup);
    const pool = primary.length ? primary : tools;
    const ranked = pool.slice().sort((a, b) => {
        const ia = HIGHLIGHT_PRIORITY.indexOf(a.id);
        const ib = HIGHLIGHT_PRIORITY.indexOf(b.id);
        return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
    });
    return ranked.slice(0, 5).map((t) => t.label);
}

function buildModules() {
    return LICENSE_PACKAGES.map((pkg) => {
        const tools = (pkg.toolIds || []).map(toolEntry);
        const highlights = pickHighlights(tools);
        const searchBits = [
            pkg.label,
            pkg.description || '',
            ...tools.map((t) => t.label),
            ...tools.map((t) => t.id.replace(/-/g, ' '))
        ];
        for (const [alias, ids] of Object.entries(SEARCH_ALIASES)) {
            if (ids.some((id) => (pkg.toolIds || []).includes(id))) {
                searchBits.push(alias);
            }
        }
        return {
            id: pkg.id,
            label: pkg.label,
            priceEuro: typeof pkg.priceEuro === 'number' ? pkg.priceEuro : null,
            blurb: String(pkg.description || '').trim(),
            requires: Array.isArray(pkg.requires) ? pkg.requires.slice() : [],
            highlights,
            tools,
            searchText: searchBits.join(' ').toLowerCase()
        };
    });
}

const catalog = {
    generatedAt: new Date().toISOString(),
    contactMail: 'kontakt@ms365.schule',
    bundles: SALES_BUNDLES.map((b) => ({
        id: b.id,
        label: b.label,
        audience: b.audience,
        priceEuro: b.priceEuro,
        priceLabel: b.priceLabel,
        packageIds: (b.packageIds || []).slice(),
        featured: !!b.featured,
        tagline: b.tagline,
        highlights: (b.highlights || []).slice()
    })),
    modules: buildModules(),
    addons: SALES_ADDONS.map((a) => ({
        id: a.id,
        label: a.label,
        priceEuro: a.priceEuro,
        priceLabel: a.priceLabel,
        blurb: a.tagline,
        once: !!a.once,
        packageIds: (a.packageIds || []).slice()
    })),
    searchAliases: SEARCH_ALIASES
};

const body =
    '/**\n' +
    ' * Auto-generated – nicht von Hand editieren.\n' +
    ' * Quelle: license-feature-catalog + dashboard-audience-catalog\n' +
    ' * Neu erzeugen: node scripts/sync-landing-offer-catalog.mjs\n' +
    ' */\n' +
    'window.MS365_OFFER_CATALOG = ' +
    JSON.stringify(catalog, null, 2) +
    ';\n';

fs.writeFileSync(outFile, body, 'utf8');
console.log(
    'wrote',
    path.relative(root, outFile),
    '·',
    catalog.modules.length,
    'modules ·',
    catalog.modules.reduce((n, m) => n + m.tools.length, 0),
    'tools'
);
