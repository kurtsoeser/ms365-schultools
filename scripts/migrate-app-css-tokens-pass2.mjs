/**
 * Zweiter Durchlauf: Rest-Hex in app.css → Tokens, ohne var() zu zerstören.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const APP_CSS = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'app.css');
let css = fs.readFileSync(APP_CSS, 'utf8');
const before = (css.match(/#[0-9a-fA-F]{3,8}\b/g) || []).length;

// Nested fallbacks first: var(--a, var(--b)) → var(--a)
css = css.replace(/var\((--[\w-]+)\s*,\s*var\(--[\w-]+\)\)/g, 'var($1)');

const pairs = [
    [/#b42318\b/gi, 'var(--status-danger)'],
    [/#fff0f3\b/gi, 'color-mix(in srgb, var(--danger1) 10%, var(--card))'],
    [/#cbd5e1\b/gi, 'var(--neutral-200)'],
    [/#1e293b\b/gi, 'var(--ink-soft)'],
    [/#0f172a\b/gi, 'var(--ink)'],
    [/#243044\b/gi, 'var(--surface-muted)'],
    [/#24283e\b/gi, 'var(--surface-muted)'],
    [/#146c43\b/gi, 'var(--status-ok-strong)'],
    [/#fff3cd\b/gi, 'color-mix(in srgb, var(--accent-amber) 22%, var(--card))'],
    [/#ffc107\b/gi, 'var(--accent-amber)'],
    [/#856404\b/gi, 'var(--status-warn-strong)'],
    [/#d1ecf1\b/gi, 'color-mix(in srgb, var(--status-ask) 14%, var(--card))'],
    [/#17a2b8\b/gi, 'var(--status-ask)'],
    [/#0c5460\b/gi, 'var(--status-ask)'],
    [/#d4edda\b/gi, 'color-mix(in srgb, var(--ok1) 18%, var(--card))'],
    [/#28a745\b/gi, 'var(--ok1)'],
    [/#155724\b/gi, 'var(--status-ok-strong)'],
    [/#f8f9ff\b/gi, 'var(--soft)'],
    [/#012456\b/gi, 'var(--bg2)'],
    [/#f0f0f0\b/gi, 'var(--neutral-100)'],
    [/#f5a623\b/gi, 'var(--accent-amber)'],
    [/#12b886\b/gi, 'var(--brand1)'],
    [/#0b7a56\b/gi, 'var(--status-ok)'],
    [/#64748b\b/gi, 'var(--muted)'],
    [/#475569\b/gi, 'var(--text-secondary)'],
    [/#fa5252\b/gi, 'var(--danger1)'],
    [/#c92a2a\b/gi, 'var(--status-danger)'],
    [/#f0ad4e\b/gi, 'var(--accent-amber)'],
    [/#9a6b00\b/gi, 'var(--status-warn-strong)'],
    [/#a33\b/gi, 'var(--status-danger)'],
    [/#c44\b/gi, 'var(--danger1)'],
    [/#8a5a00\b/gi, 'var(--status-warn-strong)'],
    [/#888\b/gi, 'var(--muted)'],
    [/#889\b/gi, 'var(--muted)'],
    [/#000\b/gi, 'var(--ink)'],
    [/#94a3b8\b/gi, 'var(--muted)'],
    [/#1a1d21\b/gi, 'var(--card)'],
    [/#69db7c\b/gi, 'var(--ok1)'],
    [/#91a7ff\b/gi, 'var(--brand1)'],
    [/#ffe066\b/gi, 'var(--accent-amber)'],
    [/#ffa8a8\b/gi, 'var(--danger1)'],
    [/#c0922d\b/gi, 'var(--status-warn)'],
    [/#7a5a14\b/gi, 'var(--status-warn-strong)'],
    [/#4c6ef5\b/gi, 'var(--brand1)'],
    [/#364fc7\b/gi, 'var(--brand1)'],
    [/#92400e\b/gi, 'var(--status-warn-strong)'],
    [/#1a1f36\b/gi, 'var(--ink)'],
    [/#1f2544\b/gi, 'var(--heading)'],
    [/#5c6370\b/gi, 'var(--muted)'],
    [/#d0d5e0\b/gi, 'var(--hairline)'],
    [/#4a5568\b/gi, 'var(--text-secondary)'],
    [/#f5f7fb\b/gi, 'var(--paper)'],
    [/#2b3358\b/gi, 'var(--heading)'],
    [/#e3e7ef\b/gi, 'var(--hairline-soft)'],
    [/#d9dee8\b/gi, 'var(--hairline)'],
    [/#cfd6e4\b/gi, 'var(--hairline)'],
    [/#0d8050\b/gi, 'var(--status-ok)'],
    [/#b36b00\b/gi, 'var(--status-warn)'],
    [/#eef0f5\b/gi, 'var(--neutral-100)'],
    [/#8b93a7\b/gi, 'var(--muted)'],
    [/#3a4158\b/gi, 'var(--heading)'],
    [/#b8c0d0\b/gi, 'var(--neutral-200)'],
    [/#fafbfd\b/gi, 'var(--paper)'],
    [/#eef\b/gi, 'var(--soft)'],
    [/#0b5cab\b/gi, 'var(--status-ask)'],
    [/#fafbfc\b/gi, 'var(--paper)'],
    [/#5c6578\b/gi, 'var(--text-secondary)'],
    [/#15182a\b/gi, 'var(--soft)'],
    [/#aeb4c8\b/gi, 'var(--text-secondary)'],
    [/#1c1f33\b/gi, 'var(--card)'],
    [/#f3f4ff\b/gi, 'var(--heading)'],
    [/#6ee7b7\b/gi, 'var(--ok1)'],
    [/#fbbf24\b/gi, 'var(--accent-amber)'],
    [/#fcd34d\b/gi, 'var(--accent-amber)'],
    [/#fb7185\b/gi, 'var(--danger1)'],
    [/#fda4af\b/gi, 'var(--danger1)'],
    [/#7dd3fc\b/gi, 'var(--status-ask)'],
    [/#0ea5e9\b/gi, 'var(--status-ask)'],
    [/#1a1f4a\b/gi, 'var(--bg2)'],
    [/#2a1648\b/gi, 'var(--bg2)'],
    [/#c9a227\b/gi, 'var(--status-warn)'],
    [/#e67700\b/gi, 'var(--status-warn)']
];

for (const [re, rep] of pairs) css = css.replace(re, rep);

// Any remaining var(--x, #hex) → var(--x)
css = css.replace(/var\((--[\w-]+)\s*,\s*#[0-9a-fA-F]{3,8}\)/g, 'var($1)');
// Nested again after replacements
css = css.replace(/var\((--[\w-]+)\s*,\s*var\(--[\w-]+\)\)/g, 'var($1)');

fs.writeFileSync(APP_CSS, css, 'utf8');
const after = (css.match(/#[0-9a-fA-F]{3,8}\b/g) || []).length;
const emptyVars = (css.match(/var\(\s*\)/g) || []).length;
console.log(`hex: ${before} → ${after}; empty var(): ${emptyVars}`);
if (after > 0) {
    const leftover = {};
    for (const m of css.matchAll(/#[0-9a-fA-F]{3,8}\b/g)) {
        const h = m[0].toLowerCase();
        leftover[h] = (leftover[h] || 0) + 1;
    }
    console.log(leftover);
}
