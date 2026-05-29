/**
 * Fügt access-config.js + pin-gate.js in alle HTML-Seiten ein (außer welcome.html).
 * Einmalig ausführen: node scripts/inject-pin-gate.mjs
 */
import fs from 'node:fs/promises';
import path from 'node:path';

const projectRoot = path.resolve(import.meta.dirname, '..');
const MARKER = 'pin-gate.js';

const depthToPrefix = {
    0: '',
    1: '../',
    2: '../../',
    3: '../../../'
};

function htmlDepth(relPath) {
    const dir = path.dirname(relPath);
    if (dir === '.') return 0;
    return dir.split(/[/\\]/).length;
}

function gateSnippet(depth) {
    const p = depthToPrefix[depth] ?? '../'.repeat(depth);
    return (
        `    <script src="${p}access-config.js"></script>\n` +
        `    <script src="${p}src/shared/pin-gate.js"></script>\n`
    );
}

async function walk(dir, files = []) {
    const entries = await fs.readdir(dir, { withFileTypes: true });
    for (const entry of entries) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
            if (entry.name === 'node_modules' || entry.name === 'dist') continue;
            await walk(full, files);
        } else if (entry.name.endsWith('.html')) {
            files.push(full);
        }
    }
    return files;
}

async function injectFile(absPath) {
    const rel = path.relative(projectRoot, absPath).replace(/\\/g, '/');
    if (rel === 'welcome.html') return 'skip-welcome';

    let html = await fs.readFile(absPath, 'utf8');
    if (html.includes(MARKER)) return 'already';

    const depth = htmlDepth(rel);
    const snippet = gateSnippet(depth);
    const charsetRe = /(<meta\s+charset="UTF-8"\s*\/?>\s*\n)/i;
    if (!charsetRe.test(html)) {
        return 'no-charset';
    }
    html = html.replace(charsetRe, `$1${snippet}`);
    await fs.writeFile(absPath, html, 'utf8');
    return 'injected';
}

const files = await walk(projectRoot);
const stats = { injected: 0, already: 0, skip: 0, failed: [] };

for (const file of files) {
    const result = await injectFile(file);
    if (result === 'injected') stats.injected += 1;
    else if (result === 'already') stats.already += 1;
    else if (result === 'skip-welcome') stats.skip += 1;
    else stats.failed.push({ file: path.relative(projectRoot, file), result });
}

console.log(
    `PIN-Gate: ${stats.injected} eingefügt, ${stats.already} bereits vorhanden, ${stats.skip} übersprungen (welcome.html).`
);
if (stats.failed.length) {
    console.warn('Nicht angepasst:', stats.failed);
    process.exitCode = 1;
}
