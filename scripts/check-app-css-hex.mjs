/**
 * Lint: app.css darf nur Token-Nutzung enthalten — max. MAX freie Hex-Farben.
 * Hex-Definitionen leben in src/shared/design-tokens.css.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const APP_CSS = path.join(ROOT, 'app.css');
const MAX = 30;

const css = fs.readFileSync(APP_CSS, 'utf8');
const matches = [...css.matchAll(/#[0-9a-fA-F]{3,8}\b/g)].map((m) => m[0].toLowerCase());
const count = matches.length;

if (count > MAX) {
    const freq = {};
    for (const h of matches) freq[h] = (freq[h] || 0) + 1;
    const top = Object.entries(freq)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15)
        .map(([h, n]) => `  ${n}× ${h}`)
        .join('\n');
    console.error(
        `check-app-css-hex: ${count} Hex in app.css (Limit ${MAX}).\n` +
            `Farben gehören in src/shared/design-tokens.css.\nHäufigste:\n${top}`
    );
    process.exit(1);
}

console.log(`check-app-css-hex: ok (${count}/${MAX} Hex in app.css)`);
