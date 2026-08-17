import { readFileSync, writeFileSync } from 'node:fs';

const files = [
    'tools/leere-gruppen-report.html',
    'tools/gaeste-verwalten.html',
    'tools/kursteams.html'
];

const map = [
    [/color:\s*#32325d/gi, 'color: var(--heading)'],
    [/color:\s*#495057/gi, 'color: var(--text-secondary)'],
    [/color:\s*#6c757d/gi, 'color: var(--muted)'],
    [/color:\s*#212529/gi, 'color: var(--heading)'],
    [/color:\s*#084298/gi, 'color: var(--brand1)'],
    [/background:\s*#fff\b/gi, 'background: var(--card)'],
    [/background:\s*#ffffff/gi, 'background: var(--card)'],
    [/background:\s*#fafbff/gi, 'background: var(--soft)'],
    [/background:\s*#f8f9fa/gi, 'background: var(--soft)'],
    [/background:\s*#f8f9ff/gi, 'background: var(--soft)'],
    [/background:\s*#f1f3f5/gi, 'background: var(--surface-muted)'],
    [/background:\s*#eef1ff/gi, 'background: color-mix(in srgb, var(--brand1) 12%, var(--card))'],
    [/background:\s*#e7f0ff/gi, 'background: color-mix(in srgb, var(--brand1) 14%, var(--card))'],
    [/background:\s*#f0f7ff/gi, 'background: color-mix(in srgb, var(--brand1) 10%, var(--card))'],
    [/background:\s*#f0fff4/gi, 'background: color-mix(in srgb, var(--ok1) 12%, var(--card))'],
    [/border:\s*1px solid #ced4da/gi, 'border: 1px solid var(--input-border)'],
    [/border:\s*1px solid #dee2e6/gi, 'border: 1px solid var(--border)'],
    [/border:\s*1px solid #e9ecef/gi, 'border: 1px solid var(--border)'],
    [/border-bottom:\s*1px solid #f0f1f3/gi, 'border-bottom: 1px solid var(--border)'],
    [/border-bottom:\s*1px solid #e9ecef/gi, 'border-bottom: 1px solid var(--border)'],
    [/border-top:\s*1px solid #e9ecef/gi, 'border-top: 1px solid var(--border)'],
    [/border:\s*1px solid #b6d4fe/gi, 'border: 1px solid color-mix(in srgb, var(--brand1) 35%, var(--border))'],
    [/border:\s*1px solid #5e72e4/gi, 'border: 1px solid var(--border-strong)']
];

for (const f of files) {
    let s = readFileSync(f, 'utf8');
    let n = 0;
    for (const [re, rep] of map) {
        s = s.replace(re, () => {
            n += 1;
            return rep;
        });
    }
    writeFileSync(f, s);
    console.log(f, 'replacements', n);
}
