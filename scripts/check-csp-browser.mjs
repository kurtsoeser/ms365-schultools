/**
 * Lädt jede App-Seite aus dist/ in Chromium (über `vite preview`) und meldet
 * Verstöße gegen die Content-Security-Policy, Skriptfehler, fehlende eigene
 * Dateien und Zugriffe auf fremde Hosts beim Seitenaufbau.
 *
 * Ersatz für eine Report-Only-Phase: Die CSP kommt als <meta>-Tag (kein
 * Header-Zugriff auf dem Hosting), dort gibt es kein Report-Only.
 *
 *   npm run build && npm run check:csp:browser
 *   node scripts/check-csp-browser.mjs tools/kursteams.html index.html   (nur einzelne Seiten)
 *
 * Geprüft wird der Seitenaufbau ohne Anmeldung. Abläufe nach Klicks
 * (Anmelden, Export, Druck) deckt das Skript nicht ab.
 */
import { spawn } from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import { chromium } from 'playwright';

const PORT = Number(process.env.CSP_CHECK_PORT || 4179);
const ORIGIN = `http://localhost:${PORT}`;
const CONCURRENCY = 6;
const SETTLE_MS = 1500;
const distRoot = path.resolve(process.cwd(), 'dist');

async function listPages(dir, out = []) {
    for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        const rel = path.relative(distRoot, full).replace(/\\/g, '/');
        if (entry.isDirectory()) {
            if (/^(landing|assets|src|vendor)$/.test(rel)) continue;
            await listPages(full, out);
        } else if (entry.name.endsWith('.html')) {
            out.push(rel);
        }
    }
    return out;
}

async function waitForServer() {
    for (let i = 0; i < 60; i++) {
        try {
            const res = await fetch(`${ORIGIN}/index.html`);
            if (res.ok) return;
        } catch {
            /* noch nicht bereit */
        }
        await new Promise((r) => setTimeout(r, 250));
    }
    throw new Error(`vite preview antwortet nicht auf ${ORIGIN}`);
}

const pages = process.argv.slice(2).length ? process.argv.slice(2) : (await listPages(distRoot)).sort();
if (!pages.length) {
    console.error('check-csp-browser: keine Seiten in dist/ – zuerst `npm run build`.');
    process.exit(1);
}

const server = spawn(process.execPath, ['node_modules/vite/bin/vite.js', 'preview', '--port', String(PORT), '--strictPort'], {
    stdio: 'ignore'
});
let browser;
const results = [];
try {
    await waitForServer();
    browser = await chromium.launch();
    const context = await browser.newContext();
    // Ohne PIN-Abfrage: sonst leiten die Seiten sofort auf welcome.html um.
    await context.addInitScript(() => {
        try {
            sessionStorage.setItem('ms365-access-granted-v1', '1');
        } catch {
            /* ignore */
        }
        document.addEventListener('securitypolicyviolation', (e) => {
            const where = e.sourceFile ? ` (${e.sourceFile}:${e.lineNumber})` : '';
            window.__cspViolation?.(`${e.effectiveDirective} blockiert ${e.blockedURI || 'inline'}${where}`);
        });
    });

    const queue = [...pages];
    async function worker() {
        for (let rel = queue.shift(); rel; rel = queue.shift()) {
            const r = { page: rel, csp: new Set(), errors: [], missing: new Set(), external: new Set(), finalUrl: '' };
            const page = await context.newPage();
            await page.exposeFunction('__cspViolation', (text) => r.csp.add(text));
            page.on('console', (msg) => {
                if (/Content Security Policy/i.test(msg.text())) r.csp.add(msg.text().slice(0, 300));
            });
            // Über das DevTools-Protokoll, weil nur dort Datei und Zeile mitkommen (auch bei Syntaxfehlern).
            const cdp = await context.newCDPSession(page);
            await cdp.send('Runtime.enable');
            cdp.on('Runtime.exceptionThrown', ({ exceptionDetails: d }) => {
                const text = d.exception?.description?.split('\n')[0] || d.text || 'Fehler';
                const where = d.url ? ` (${d.url.replace(ORIGIN, '')}:${d.lineNumber + 1})` : '';
                r.errors.push(`${text.slice(0, 240)}${where}`);
            });
            page.on('response', (res) => {
                if (res.url().startsWith(ORIGIN) && res.status() >= 400) r.missing.add(`${res.status()} ${res.url().slice(ORIGIN.length)}`);
            });
            // Fremde Hosts nicht wirklich aufrufen – nur festhalten, wer beim Laden angesprochen wird.
            await page.route(
                (url) => !url.href.startsWith(ORIGIN) && /^https?:$/.test(url.protocol),
                (route) => {
                    r.external.add(`${route.request().resourceType()} ${new URL(route.request().url()).host}`);
                    return route.abort();
                }
            );
            try {
                await page.goto(`${ORIGIN}/${rel}`, { waitUntil: 'load', timeout: 30000 });
                await page.waitForTimeout(SETTLE_MS);
                r.finalUrl = page.url().slice(ORIGIN.length);
            } catch (err) {
                r.errors.push(`Laden fehlgeschlagen: ${String(err.message || err).slice(0, 200)}`);
            }
            await page.close();
            results.push(r);
        }
    }
    await Promise.all(Array.from({ length: CONCURRENCY }, worker));
} finally {
    if (browser) await browser.close();
    server.kill();
}

results.sort((a, b) => a.page.localeCompare(b.page));
const withCsp = results.filter((r) => r.csp.size);
const withErrors = results.filter((r) => r.errors.length);
const withMissing = results.filter((r) => r.missing.size);
const externalScripts = results.filter((r) => [...r.external].some((e) => /^(script|stylesheet|font) /.test(e)));

const section = (title, list, pick) => {
    if (!list.length) return;
    console.log(`\n${title}`);
    for (const r of list) {
        console.log(`  ${r.page}`);
        for (const line of pick(r)) console.log(`    - ${line}`);
    }
};
section('CSP-Verstöße:', withCsp, (r) => [...r.csp]);
section('Skriptfehler:', withErrors, (r) => r.errors);
section('Fehlende eigene Dateien (HTTP ≥ 400):', withMissing, (r) => [...r.missing]);
section('Skripte/Styles/Schriften von fremden Hosts:', externalScripts, (r) => [...r.external]);

const hosts = new Map();
for (const r of results) for (const e of r.external) hosts.set(e, (hosts.get(e) || 0) + 1);
if (hosts.size) {
    console.log('\nFremde Hosts beim Seitenaufbau (Art Host: Seitenanzahl):');
    for (const [e, n] of [...hosts].sort()) console.log(`  ${e}: ${n}`);
}

console.log(
    `\ncheck-csp-browser: ${results.length} Seiten geladen – ${withCsp.length} mit CSP-Verstoß, ${withErrors.length} mit Skriptfehler, ` +
        `${withMissing.length} mit fehlenden Dateien, ${externalScripts.length} mit fremden Skripten/Styles.`
);
process.exit(withCsp.length || externalScripts.length ? 1 : 0);
