import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

const patches = [
    [
        'src/tools/elternbrief/elternbrief.css',
        `
/* Welle 3 quiet */
.eb-cards--quiet { gap: 10px; }
.eb-cards--quiet .eb-card {
    background: transparent;
    border: 0;
    box-shadow: none;
    padding: 4px 8px 8px 0;
    border-radius: 0;
    border-bottom: 1px solid var(--hairline, var(--border));
}
.eb-cards--quiet .eb-card__value {
    font-size: 1.35rem;
    font-weight: 650;
    letter-spacing: -0.02em;
}
.eb-head .fb-greet { margin: 0 0 4px; }
.eb-head .fb-sub { margin: 0; }
`
    ],
    [
        'src/tools/einsammler/einsammler.css',
        `
/* Welle 3 quiet */
.ein-cards--quiet { gap: 10px; }
.ein-cards--quiet .ein-card {
    background: transparent;
    border: 0;
    box-shadow: none;
    padding: 4px 8px 8px 0;
    border-radius: 0;
    border-bottom: 1px solid var(--hairline, var(--border));
}
.ein-cards--quiet .ein-card__value {
    font-size: 1.35rem;
    font-weight: 650;
    letter-spacing: -0.02em;
}
.ein-head .fb-greet { margin: 0 0 4px; }
.ein-head .fb-sub { margin: 0; }
`
    ],
    [
        'src/tools/krankmeldung/krankmeldung.css',
        `
/* Welle 3 quiet */
.km-cards--quiet { gap: 10px; }
.km-cards--quiet .km-card {
    background: transparent;
    border: 0;
    box-shadow: none;
    padding: 4px 8px 8px 0;
    border-radius: 0;
    border-bottom: 1px solid var(--hairline, var(--border));
}
.km-cards--quiet .km-card__value {
    font-size: 1.35rem;
    font-weight: 650;
    letter-spacing: -0.02em;
}
.km-head .fb-greet { margin: 0 0 4px; }
.km-head .fb-sub { margin: 0; }
`
    ],
    [
        'src/tools/projektwochen/projektwochen.css',
        `
/* Welle 3 quiet */
.pw-kpi-grid--quiet { gap: 10px; }
.pw-kpi-grid--quiet .pw-kpi {
    background: transparent;
    border: 0;
    box-shadow: none;
    padding: 6px 8px 10px 0;
    border-radius: 0;
    border-bottom: 1px solid var(--hairline, var(--border));
}
.pw-view__head .fb-greet { margin: 0 0 4px; }
.pw-view__head .fb-sub { margin: 0; }
`
    ]
];

for (const [rel, block] of patches) {
    const file = path.join(root, rel);
    let s = fs.readFileSync(file, 'utf8');
    if (s.includes('Welle 3 quiet')) {
        console.log('skip', rel);
        continue;
    }
    fs.writeFileSync(file, s.trimEnd() + '\n' + block);
    console.log('patched', rel);
}
