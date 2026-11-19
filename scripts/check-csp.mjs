/**
 * Prüft den Build-Output (dist/) gegen die Content-Security-Policy:
 * jede App-Seite trägt das CSP-Meta-Tag vor dem ersten Skript, enthält keine
 * Inline-Skripte, keine Inline-Handler und lädt weder Skripte noch Styles von
 * fremden Hosts. Bricht den Build ab, damit ein Verstoß nicht deployt wird.
 *
 *   node scripts/check-csp.mjs          (nach vite build + copy-static)
 *
 * dist/landing/ (Marketing-Seite, eigenes Deployment) ist ausgenommen.
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const DATA_SCRIPT_TYPES = /^(text\/plain|application\/json|application\/ld\+json|importmap|text\/template|text\/x-template)$/i;

/**
 * @param {string} html
 * @param {{ source?: boolean }} [opts] source: Quell-HTML vor dem Build – das Meta-Tag
 *   kommt erst beim Build dazu, und Inline-Module lagert Vite in Dateien aus.
 * @returns {string[]} Verstöße (leer = in Ordnung)
 */
export function findCspViolations(html, opts = {}) {
    const problems = [];
    const noComments = html.replace(/<!--[\s\S]*?-->/g, '');

    if (!opts.source) {
        const metaAt = noComments.search(/<meta\s+http-equiv=["']Content-Security-Policy["']/i);
        const firstScriptAt = noComments.search(/<script[\s>]/i);
        if (metaAt === -1) problems.push('CSP-Meta-Tag fehlt');
        else if (firstScriptAt !== -1 && firstScriptAt < metaAt) problems.push('CSP-Meta-Tag steht nicht vor dem ersten <script>');
    }

    for (const m of noComments.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)) {
        const attrs = m[1];
        const type = (attrs.match(/\btype\s*=\s*["']?([^"'\s>]+)/i) || [])[1] || '';
        const src = (attrs.match(/\bsrc\s*=\s*["']?([^"'\s>]+)/i) || [])[1] || '';
        if (DATA_SCRIPT_TYPES.test(type)) continue;
        if (!src) {
            if (opts.source && /^module$/i.test(type)) continue;
            if (m[2].trim()) problems.push(`Inline-Skript: ${m[2].trim().slice(0, 70).replace(/\s+/g, ' ')}…`);
            continue;
        }
        if (/^(https?:)?\/\//i.test(src)) problems.push(`Skript von fremdem Host: ${src}`);
    }

    // Markup ohne Skript-/Style-Inhalte, damit Text in Templates nicht als Attribut zählt.
    const markup = noComments.replace(/<script\b[\s\S]*?<\/script>/gi, '').replace(/<style\b[\s\S]*?<\/style>/gi, '');
    for (const m of markup.matchAll(/<[a-z][^>]*?\s(on[a-z]+)\s*=\s*["'][^>]*>/gi)) {
        problems.push(`Inline-Handler ${m[1]}: ${m[0].slice(0, 90)}…`);
    }
    for (const m of markup.matchAll(/<link\b[^>]*\brel\s*=\s*["']?stylesheet[^>]*>/gi)) {
        const href = (m[0].match(/\bhref\s*=\s*["']?([^"'\s>]+)/i) || [])[1] || '';
        if (/^(https?:)?\/\//i.test(href)) problems.push(`Stylesheet von fremdem Host: ${href}`);
    }
    if (/\b(?:href|src)\s*=\s*["']?\s*javascript:/i.test(markup)) problems.push('javascript:-URL');
    return problems;
}

async function listHtml(dir, out = []) {
    for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) await listHtml(full, out);
        else if (entry.name.endsWith('.html')) out.push(full);
    }
    return out;
}

/**
 * @param {string} distRoot
 * @returns {Promise<{ checked: number, failures: Array<{ file: string, problems: string[] }> }>}
 */
export async function checkDist(distRoot) {
    const files = await listHtml(distRoot);
    const failures = [];
    let checked = 0;
    for (const file of files) {
        const rel = path.relative(distRoot, file).replace(/\\/g, '/');
        // landing/ = Marketing-Seite; src/ und vendor/ = 1:1-Kopien ohne Seiten.
        if (/^(landing|src|vendor)\//.test(rel)) continue;
        checked++;
        const problems = findCspViolations(await fs.readFile(file, 'utf8'));
        if (problems.length) failures.push({ file: rel, problems });
    }
    return { checked, failures };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
    const distRoot = path.resolve(process.cwd(), 'dist');
    const { checked, failures } = await checkDist(distRoot);
    if (!checked) {
        console.error('check-csp: keine HTML-Seiten in dist/ gefunden – zuerst bauen.');
        process.exit(1);
    }
    if (failures.length) {
        console.error(`check-csp: ${failures.length} von ${checked} Seiten verstoßen gegen die Content-Security-Policy:`);
        for (const f of failures) {
            console.error(`  ${f.file}`);
            for (const p of f.problems) console.error(`    - ${p}`);
        }
        process.exit(1);
    }
    console.log(`check-csp: ${checked} Seiten ohne Inline-Skripte, Inline-Handler und fremde Skript-Hosts.`);
}
