import { readFileSync, writeFileSync } from 'node:fs';

const files = [
    'tools/leere-gruppen-report.html',
    'tools/gaeste-verwalten.html',
    'tools/kursteams.html'
];

const map = [
    [
        /background:\s*linear-gradient\(180deg,\s*#fff5f5\s*0%,\s*#fff\s*100%\)/gi,
        'background: color-mix(in srgb, var(--danger1) 12%, var(--card))'
    ],
    [/background:\s*#fff5f5/gi, 'background: color-mix(in srgb, var(--danger1) 12%, var(--card))'],
    [/color:\s*#7b2929/gi, 'color: color-mix(in srgb, var(--danger1) 70%, var(--heading))'],
    [/color:\s*#9b2c2c/gi, 'color: color-mix(in srgb, var(--danger1) 75%, var(--heading))'],
    [/\.gv-tab:hover\s*\{\s*background:\s*rgba\(255,\s*255,\s*255,\s*0\.9\);\s*color:\s*var\(--heading\);\s*\}/gi,
        '.gv-tab:hover { background: color-mix(in srgb, var(--brand1) 10%, var(--card)); color: var(--heading); }'],
    [/\.gv-tab\s*\{\s*\n(\s*)border: 1px solid transparent;\s*\n\s*background: transparent;\s*\n\s*color: var\(--text-secondary\);/gi,
        '.gv-tab {\n$1border: 1px solid transparent;\n$1background: transparent;\n$1color: var(--text-secondary);']
];

for (const f of files) {
    let s = readFileSync(f, 'utf8');
    let n = 0;
    for (const [re, rep] of map) {
        s = s.replace(re, (...args) => {
            n += 1;
            if (typeof rep === 'string' && rep.includes('$1')) {
                return rep.replace('$1', args[1] || '            ');
            }
            return rep;
        });
    }
    // Dark-friendly selected rows
    s = s.replace(
        /tbody tr\.is-selected\s*\{\s*background:\s*color-mix\(in srgb, var\(--danger1\) 12%, var\(--card\)\);\s*\}/g,
        'tbody tr.is-selected { background: color-mix(in srgb, var(--danger1) 14%, var(--card)); }'
    );
    writeFileSync(f, s);
    console.log(f, 'warn fixes', n);
}
