/**
 * Welle 3 – Fluent-Chrome für alle Schulwerkzeuge (HTML-Head + Teal-Overrides).
 * Idempotent: mehrfach lauffähig.
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/** Portal-/Alltags-Schulwerkzeuge (ca. 23+) – nicht reine IT-Setup-Listen. */
const SCHOOL_TOOLS = [
    'freistellung-planer',
    'lehrer-freistellung-planer',
    'schulaktivitaeten-planer',
    'schularbeiten-planer',
    'srdp-anmeldung',
    'projektwochen',
    'termine-hub',
    'aufgaben',
    'technik-meldungen',
    'raumplan',
    'elternbrief',
    'einsammler',
    'krankmeldung',
    'kollegium-umfragen',
    'teams-antrag',
    'diplomarbeiten',
    'kursteams',
    'unterrichtsteams-katalog',
    'onenote-verteilung',
    'elternsprechtag-bookings',
    'sharepoint-intranet-hub',
    'klassenchats',
    'schulsprecherwahl',
    'wahlfach-zuteilung',
    'wahlfach-anmeldung',
    'ausgaben',
    'personen-verwaltung',
    'schueler-lehrer-gruppen',
    'lehrer-sammelgruppe',
    'schueler-sammelgruppe',
    'diplomarbeiten',
    // Empfänger / schlanke Geschwister
    'elternbrief-lesen',
    'einsammler-antwort',
    'krankmeldung-formular',
    'schulsprecherwahl-stimme',
    'kollegium-umfragen-antwort',
    'wahlfach-anmeldung'
];

const APP_CSS_V = '28';
const SHELL_V = '10';
const BLOCKS_V = '7';
const TOKENS_V = '2';

const FLUENT_LINKS = [
    `    <link rel="stylesheet" href="../src/shared/design-tokens.css?v=${TOKENS_V}">`,
    `    <link rel="stylesheet" href="../src/shared/shell-fluent.css?v=${SHELL_V}">`,
    `    <link rel="stylesheet" href="../src/shared/fluent-blocks.css?v=${BLOCKS_V}">`
].join('\n');

const TEAL_BLOCK = `
/* ===== Welle 3 UI chrome (Schulwerkzeuge) ===== */
.page-card.tm-page,
.container.page-card {
    --brand: var(--brand1);
    --primary: var(--brand1);
}
.btn-success,
.btn.btn-success,
button.btn-success {
    background: var(--brand1) !important;
    border-color: var(--brand1) !important;
    color: var(--on-accent, #04211d) !important;
}
.btn-primary:not(.btn-ghost):not(.btn-outline) {
    background: var(--brand1);
    border-color: var(--brand1);
    color: var(--on-accent, #04211d);
}
`;

function unique(list) {
    return [...new Set(list)];
}

function patchHtml(rel) {
    const file = path.join(root, 'tools', rel + '.html');
    if (!fs.existsSync(file)) return { id: rel, status: 'missing' };

    let s = fs.readFileSync(file, 'utf8');
    const before = s;

    s = s.replace(/app\.css\?v=\d+/g, `app.css?v=${APP_CSS_V}`);

    // Tokens / shell / blocks: Versions bump oder einfügen
    if (s.includes('design-tokens.css')) {
        s = s.replace(/design-tokens\.css\?v=[^"']+/g, `design-tokens.css?v=${TOKENS_V}`);
    }
    if (s.includes('shell-fluent.css')) {
        s = s.replace(/shell-fluent\.css\?v=[^"']+/g, `shell-fluent.css?v=${SHELL_V}`);
    }
    if (s.includes('fluent-blocks.css')) {
        s = s.replace(/fluent-blocks\.css\?v=[^"']+/g, `fluent-blocks.css?v=${BLOCKS_V}`);
    }

    if (!s.includes('shell-fluent.css')) {
        if (/href="\.\.\/app\.css\?v=\d+"/.test(s)) {
            s = s.replace(
                /(<link rel="stylesheet" href="\.\.\/app\.css\?v=\d+">)/,
                `$1\n${FLUENT_LINKS}`
            );
        } else if (/href="\.\.\/src\/shared\/tool-modern\.css"/.test(s)) {
            s = s.replace(
                /(<link rel="stylesheet" href="\.\.\/src\/shared\/tool-modern\.css">)/,
                `${FLUENT_LINKS}\n$1`
            );
        } else if (/<\/title>/.test(s)) {
            s = s.replace(/<\/title>/, `</title>\n${FLUENT_LINKS}`);
        }
    } else if (!s.includes('design-tokens.css')) {
        s = s.replace(
            /(<link rel="stylesheet" href="\.\.\/src\/shared\/shell-fluent\.css\?v=[^"]+">)/,
            `    <link rel="stylesheet" href="../src/shared/design-tokens.css?v=${TOKENS_V}">\n$1`
        );
    }

    if (!s.includes('Welle 3 UI chrome (Schulwerkzeuge)')) {
        if (/<\/style>/.test(s)) {
            // letztes </style> im Head/Body
            const idx = s.lastIndexOf('</style>');
            s = s.slice(0, idx) + TEAL_BLOCK + '\n' + s.slice(idx);
        } else if (/<\/head>/.test(s)) {
            s = s.replace('</head>', `    <style>${TEAL_BLOCK}\n    </style>\n</head>`);
        }
    }

    if (s === before) return { id: rel, status: 'unchanged' };
    fs.writeFileSync(file, s);
    return { id: rel, status: 'updated' };
}

const results = unique(SCHOOL_TOOLS).map(patchHtml);
const updated = results.filter((r) => r.status === 'updated').map((r) => r.id);
const missing = results.filter((r) => r.status === 'missing').map((r) => r.id);
const unchanged = results.filter((r) => r.status === 'unchanged').map((r) => r.id);

console.log(JSON.stringify({ updated, unchanged, missing, count: results.length }, null, 2));
