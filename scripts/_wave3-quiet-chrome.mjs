import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

function patchPersonen() {
    const p = path.join(root, 'tools/personen-verwaltung.html');
    let s = fs.readFileSync(p, 'utf8');
    s = s.replace(/color:\s*#32325d/g, 'color: var(--heading)');
    s = s.replace(
        /background:\s*linear-gradient\(135deg,\s*color-mix\(in srgb, var\(--brand1\) 10%, transparent\) 0%,\s*color-mix\(in srgb, var\(--brand2\) 6%, transparent\) 100%\);/g,
        'background: var(--soft);'
    );
    s = s.replace(/font-weight:\s*1000/g, 'font-weight: 700');
    s = s.replace(/font-weight:\s*900/g, 'font-weight: 700');
    if (!s.includes('Welle 3 personen')) {
        s = s.replace(
            '</style>',
            `/* Welle 3 personen */
.panel{border-radius:10px;background:var(--card);}
.panel-head h2{font-weight:650!important;}
.btn-success{background:var(--brand1)!important;border-color:var(--brand1)!important;color:var(--on-accent,#04211d)!important;}
</style>`
        );
    }
    fs.writeFileSync(p, s);
    console.log('personen');
}

function patchKursteams() {
    const p = path.join(root, 'tools/kursteams.html');
    let s = fs.readFileSync(p, 'utf8');
    s = s.replace(
        /background:\s*linear-gradient\(135deg,\s*var\(--brand1\) 0%,\s*var\(--brand2\) 100%\);/g,
        'background: var(--brand1);'
    );
    s = s.replace(
        /background:\s*linear-gradient\(135deg,\s*var\(--brand1\),\s*var\(--brand2\)\);/g,
        'background: var(--brand1);'
    );
    s = s.replace(
        /box-shadow:\s*0 4px 12px color-mix\(in srgb, var\(--brand1\) 35%, transparent\);/g,
        'box-shadow: none;'
    );
    s = s.replace(
        /box-shadow:\s*0 6px 24px color-mix\(in srgb, var\(--brand1\) 18%, transparent\);/g,
        'box-shadow: none;'
    );
    if (!s.includes('Welle 3 kursteams')) {
        s = s.replace(
            '</style>',
            `/* Welle 3 kursteams */
.btn-success{background:var(--brand1)!important;border-color:var(--brand1)!important;color:var(--on-accent,#04211d)!important;}
.kt-step-badge{border-radius:10px!important;}
</style>`
        );
    }
    fs.writeFileSync(p, s);
    console.log('kursteams');
}

function patchSlg() {
    const p = path.join(root, 'tools/schueler-lehrer-gruppen.html');
    let s = fs.readFileSync(p, 'utf8');
    s = s.replace(/app\.css\?v=\d+/g, 'app.css?v=27');
    if (!s.includes('shell-fluent.css')) {
        s = s.replace(
            '<link rel="stylesheet" href="../app.css?v=27">',
            `<link rel="stylesheet" href="../app.css?v=27">
    <link rel="stylesheet" href="../src/shared/shell-fluent.css?v=9">
    <link rel="stylesheet" href="../src/shared/fluent-blocks.css?v=6">`
        );
    }
    fs.writeFileSync(p, s);
    console.log('slg');
}

function patchUtk() {
    const p = path.join(root, 'tools/unterrichtsteams-katalog.html');
    let s = fs.readFileSync(p, 'utf8');
    s = s.replace(/app\.css\?v=\d+/g, 'app.css?v=27');
    if (!s.includes('shell-fluent.css')) {
        s = s.replace(
            /(<link rel="stylesheet" href="\.\.\/app\.css\?v=27">)/,
            `$1\n    <link rel="stylesheet" href="../src/shared/shell-fluent.css?v=9">\n    <link rel="stylesheet" href="../src/shared/fluent-blocks.css?v=6">`
        );
    }
    fs.writeFileSync(p, s);
    console.log('utk');
}

patchPersonen();
patchKursteams();
patchSlg();
patchUtk();
