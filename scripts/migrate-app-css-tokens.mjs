/**
 * Einmal-/Wiederhol-Migration: app.css auf Design-Tokens umstellen.
 * - Token-Blöcke (:root / dark / light / alle data-brand) entfernen (leben in design-tokens.css)
 * - Rest-Lila und Marken-Fallbacks ersetzen
 * - Häufige Hex → var(--…)
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const APP_CSS = path.join(ROOT, 'app.css');

const HEADER = `/* Gemeinsames Styling für alle Tool-Seiten (tools/*.html)
 * Design-Foundation: Tokens in src/shared/design-tokens.css (einzige Hex-Quelle).
 * Marke: Teal. Farbe sonst nur für Soll/Ist-Status.
 */
@import url('src/shared/design-tokens.css');
@import url('src/shared/tools-dark.css');
@import url('src/shared/dashboard-dark.css');
@import url('src/shared/planner-apps-theme.css');
`;

function stripTokenBlocks(css) {
    // Drop old file header + imports through first :root { … }
    css = css.replace(/^\/\* Gemeinsames Styling[\s\S]*?\*\/\s*/, '');
    css = css.replace(/@import url\('src\/shared\/(?:tools-dark|dashboard-dark|planner-apps-theme)\.css'\);\s*/g, '');

    // Remove :root { … } (first occurrence — color tokens)
    css = css.replace(/:root\s*\{[\s\S]*?\n\}\n*/, '');

    // Remove html[data-theme="dark|light"] token blocks
    css = css.replace(/html\[data-theme="dark"\]\s*\{[\s\S]*?\n\}\n*/g, '');
    css = css.replace(/html\[data-theme="light"\]\s*\{[\s\S]*?\n\}\n*/g, '');

    // Remove all brand theme blocks at end (classic/wine/teal/azure/forest/amber/custom)
    css = css.replace(
        /\/\* Brand-Tokens \(am Dateiende\):[\s\S]*?(?=\/\* Admin Feedback Inbox \*\/)/,
        ''
    );
    // Fallback: any leftover html[data-brand=…] blocks
    css = css.replace(/html\[data-brand="[^"]+"\][^{]*\{[\s\S]*?\n\}\n*/g, '');
    css = css.replace(/html\[data-brand="custom"\]\s*\.header\s*\{[\s\S]*?\n\}\n*/g, '');

    return css;
}

function replaceHex(css) {
    // Fallbacks zuerst entfernen, sonst entstehen var(--a, var(--b))
    css = css.replace(/var\((--[\w-]+)\s*,\s*#[0-9a-fA-F]{3,8}\)/g, 'var($1)');
    css = css.replace(/var\((--[\w-]+)\s*,\s*var\(--[\w-]+\)\s*,\s*#[0-9a-fA-F]{3,8}\)/g, 'var($1)');

    // color: white on brand surfaces
    css = css.replace(/(^|[^-])(color\s*:\s*)#fff(?:fff)?\b/gim, '$1$2var(--on-brand)');
    css = css.replace(/(^|[^-])(color\s*:\s*)#ffffff\b/gim, '$1$2var(--on-brand)');

    const pairs = [
        // Purple / classic leftovers → brand / status
        [/#5e72e4\b/gi, 'var(--brand1)'],
        [/#825ee4\b/gi, 'var(--brand2)'],
        [/#667eea\b/gi, 'var(--brand1)'],
        [/#764ba2\b/gi, 'var(--brand2)'],
        [/#5b4bdb\b/gi, 'var(--status-new)'],
        [/#7b8cff\b/gi, 'var(--brand1)'],
        [/#a78bfa\b/gi, 'var(--brand2)'],
        [/#7c3aed\b/gi, 'var(--status-info)'],
        [/#f5f3ff\b/gi, 'color-mix(in srgb, var(--status-info) 10%, var(--card))'],

        // Brand teal hardcoded
        [/#0d9f8a\b/gi, 'var(--brand1)'],
        [/#12b59c\b/gi, 'var(--brand2)'],
        [/#14c4ab\b/gi, 'var(--brand1)'],
        [/#0f766e\b/gi, 'var(--status-ok)'],
        [/#0c6b5e\b/gi, 'var(--bg1)'],
        [/#0d6efd\b/gi, 'var(--brand1)'],
        [/#2563eb\b/gi, 'var(--brand1)'],

        // Status
        [/#b91c1c\b/gi, 'var(--status-danger)'],
        [/#b45309\b/gi, 'var(--status-warn)'],
        [/#9c6b00\b/gi, 'var(--status-warn-strong)'],
        [/#a16207\b/gi, 'var(--status-idea)'],
        [/#0369a1\b/gi, 'var(--status-ask)'],
        [/#15803d\b/gi, 'var(--status-ok-strong)'],
        [/#166534\b/gi, 'var(--status-ok-strong)'],
        [/#f5365c\b/gi, 'var(--danger1)'],
        [/#f56036\b/gi, 'var(--danger2)'],
        [/#2dce89\b/gi, 'var(--ok1)'],
        [/#11998e\b/gi, 'var(--ok2)'],
        [/#34d399\b/gi, 'var(--ok1)'],
        [/#22c55e\b/gi, 'var(--ok1)'],
        [/#fab005\b/gi, 'var(--accent-amber)'],
        [/#f59e0b\b/gi, 'var(--accent-gold)'],

        // Neutrals / text
        [/#32325d\b/gi, 'var(--heading)'],
        [/#6c757d\b/gi, 'var(--muted)'],
        [/#495057\b/gi, 'var(--neutral-700)'],
        [/#212529\b/gi, 'var(--neutral-900)'],
        [/#5b6b82\b/gi, 'var(--text-secondary)'],
        [/#0c1222\b/gi, 'var(--ink)'],
        [/#1e2440\b/gi, 'var(--ink-soft)'],
        [/#ced4da\b/gi, 'var(--input-border)'],
        [/#dee2e6\b/gi, 'var(--hairline)'],
        [/#e9ecef\b/gi, 'var(--hairline-soft)'],
        [/#f3f4f6\b/gi, 'var(--neutral-100)'],
        [/#d7dde5\b/gi, 'var(--neutral-200)'],
        [/#f8fafc\b/gi, 'var(--paper-2)'],
        [/#f4f7fb\b/gi, 'var(--paper)'],
        [/#f2f6fb\b/gi, 'var(--paper-2)'],
        [/#e8eef6\b/gi, 'var(--text)'],
        [/#1a2638\b/gi, 'var(--surface-muted)'],
        [/#121a28\b/gi, 'var(--card)'],
        [/#070b12\b/gi, 'var(--scrim)'],
        [/#07101f\b/gi, 'var(--app-backdrop-paper)'],
        [/#8b9aab\b/gi, 'var(--muted)'],
        [/#f6f7f9\b/gi, 'var(--paper)'],
        [/#f5f6fb\b/gi, 'var(--paper)'],
        [/#e8eaf4\b/gi, 'var(--soft)'],
        [/#f0f7f5\b/gi, 'var(--soft)'],
        [/#ffffff\b/gi, 'var(--card)'],
        [/#fff\b/gi, 'var(--card)'],
        [/#111\b/gi, 'var(--scrim)'],
        [/#bbb\b/gi, 'var(--scrim-muted)'],
        [/#123\b/gi, 'var(--ink)'],
        [/#234\b/gi, 'var(--ink-soft)']
    ];

    for (const [re, rep] of pairs) {
        css = css.replace(re, rep);
    }

    // rgba purple brand leftovers
    css = css.replace(
        /rgba\s*\(\s*94\s*,\s*114\s*,\s*228\s*,\s*([\d.]+)\s*\)/gi,
        (_, a) => `color-mix(in srgb, var(--brand1) ${Math.round(parseFloat(a) * 100)}%, transparent)`
    );
    css = css.replace(
        /rgba\s*\(\s*130\s*,\s*94\s*,\s*228\s*,\s*([\d.]+)\s*\)/gi,
        (_, a) => `color-mix(in srgb, var(--brand2) ${Math.round(parseFloat(a) * 100)}%, transparent)`
    );

    return css;
}

function cleanupLogoBrandRules(css) {
    // Classic logo variants no longer needed — keep teal visible
    css = css.replace(
        /\.app-brand-logo__mark--classic[\s\S]*?html:not\(\[data-brand="classic"\]\) \.app-brand-logo__mark--teal \{[\s\S]*?\n\}\n*/,
        `.app-brand-logo__mark--classic {
    display: none !important;
}
.app-brand-logo__mark--teal {
    display: block;
}
`
    );
    return css;
}

function cleanupBrandSwatches(css) {
    css = css.replace(
        /\.ms365-auth-menu__brand-swatch--teal \{[\s\S]*?\.ms365-auth-menu__brand-swatch--wine \{[\s\S]*?\n\}\n*/,
        `.ms365-auth-menu__brand-swatch--teal {
    background: linear-gradient(135deg, var(--brand1), var(--brand2));
}
`
    );
    // Remaining swatches for removed brands
    css = css.replace(/\.ms365-auth-menu__brand-swatch--(?:classic|wine|azure|forest|amber|custom)\s*\{[\s\S]*?\n\}\n*/g, '');
    return css;
}

let css = fs.readFileSync(APP_CSS, 'utf8');
const before = (css.match(/#[0-9a-fA-F]{3,8}\b/g) || []).length;

css = stripTokenBlocks(css);
css = cleanupLogoBrandRules(css);
css = cleanupBrandSwatches(css);
css = replaceHex(css);

// Ensure html backdrop still uses tokens (html { … } block kept)
css = HEADER + '\n' + css.trimStart();

// Collapse excess blank lines
css = css.replace(/\n{3,}/g, '\n\n');

fs.writeFileSync(APP_CSS, css, 'utf8');
const after = (css.match(/#[0-9a-fA-F]{3,8}\b/g) || []).length;
console.log(`app.css hex: ${before} → ${after}`);
if (after >= 30) {
    const leftover = {};
    for (const m of css.matchAll(/#[0-9a-fA-F]{3,8}\b/g)) {
        const h = m[0].toLowerCase();
        leftover[h] = (leftover[h] || 0) + 1;
    }
    console.log('leftover:', Object.entries(leftover).sort((a, b) => b[1] - a[1]));
}
