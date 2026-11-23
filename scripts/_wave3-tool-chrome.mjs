/**
 * Welle-3: Fluent CSS + Cache-Bust in Tool-HTML, Teal-Override ans Tool-CSS hängen.
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

const tools = [
    {
        html: 'tools/srdp-anmeldung.html',
        css: 'src/tools/srdp-anmeldung/srdp-anmeldung.css',
        cssQ: 'srdp-anmeldung.css?v=3'
    },
    {
        html: 'tools/schueler-sammelgruppe.html',
        css: null
    },
    {
        html: 'tools/lehrer-sammelgruppe.html',
        css: null
    },
    {
        html: 'tools/personen-verwaltung.html',
        css: null
    },
    {
        html: 'tools/kursteams.html',
        css: null
    },
    {
        html: 'tools/unterrichtsteams-katalog.html',
        css: 'src/tools/unterrichtsteams-katalog/unterrichtsteams-katalog.css',
        cssQ: 'unterrichtsteams-katalog.css?v=3'
    },
    {
        html: 'tools/srdp-mail-setup.html',
        css: null
    }
];

const FLUENT = `    <link rel="stylesheet" href="../src/shared/shell-fluent.css?v=9">
    <link rel="stylesheet" href="../src/shared/fluent-blocks.css?v=6">
`;

const TEAL_BLOCK = `
/* ===== Welle 3 UI chrome ===== */
.page-card.tm-page {
    --brand: var(--brand1);
    --primary: var(--brand1);
}
.btn-success,
.btn.btn-success {
    background: var(--brand1) !important;
    border-color: var(--brand1) !important;
    color: var(--on-accent, #04211d) !important;
}
.fb-greet { margin: 0 0 6px; }
.fb-sub { margin: 0 0 16px; }
.fb-actions-row { margin-top: 16px; }
`;

function ensureFluent(html) {
    let s = fs.readFileSync(path.join(root, html), 'utf8');
    s = s.replace(/app\.css\?v=\d+/g, 'app.css?v=27');
    if (!s.includes('shell-fluent.css')) {
        s = s.replace(
            /(<link rel="stylesheet" href="\.\.\/app\.css\?v=27">)/,
            `$1\n${FLUENT}`
        );
    } else {
        s = s.replace(/shell-fluent\.css\?v=\d+/g, 'shell-fluent.css?v=9');
        s = s.replace(/fluent-blocks\.css\?v=\d+/g, 'fluent-blocks.css?v=6');
    }
    return s;
}

for (const t of tools) {
    let html = ensureFluent(t.html);
    if (t.cssQ) {
        const base = t.cssQ.split('?')[0];
        html = html.replace(
            new RegExp(base.replace('.', '\\.') + '(\\?v=[^"\']*)?'),
            t.cssQ
        );
    }
    fs.writeFileSync(path.join(root, t.html), html);
    console.log('html', t.html);

    if (t.css) {
        const cssPath = path.join(root, t.css);
        let css = fs.readFileSync(cssPath, 'utf8');
        if (!css.includes('Welle 3 UI chrome')) {
            css += TEAL_BLOCK;
            fs.writeFileSync(cssPath, css);
            console.log('css', t.css);
        }
    }
}

// SA + AKT already done; append SA teal quiet status if missing
const saCss = path.join(root, 'src/tools/schularbeiten-planer/schularbeiten-planer.css');
let sa = fs.readFileSync(saCss, 'utf8');
if (!sa.includes('Welle 3 UI chrome')) {
    sa += TEAL_BLOCK + `
.sa-page .fb-greet { margin: 0 0 6px; }
.sa-quiet-status { display:flex; gap:8px; align-items:center; color:var(--muted); }
.sa-top-refresh { margin-bottom: 8px; }
`;
    fs.writeFileSync(saCss, sa);
    console.log('css schularbeiten');
}

console.log('done');
