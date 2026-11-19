/**
 * Wächter für die Browser-Vertrauenskette (Audit R1):
 * kein Code von fremden CDNs, keine Inline-Skripte/-Handler und kein eval
 * (Voraussetzung für script-src 'self'), Deploy nur verschlüsselt,
 * GitHub Actions auf Commit-SHAs gepinnt.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

import { findCspViolations } from '../scripts/check-csp.mjs';
import { CSP_DIRECTIVES, buildCspContent, injectCspMeta } from '../scripts/csp-policy.mjs';
import { buildLftpScript, normalizeRemoteDir } from '../scripts/deploy-upload.mjs';

const projectRoot = join(dirname(fileURLToPath(import.meta.url)), '..');
const rel = (p) => relative(projectRoot, p).replace(/\\/g, '/');

function walk(dir, match, out = []) {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
        const full = join(dir, entry.name);
        if (entry.isDirectory()) walk(full, match, out);
        else if (match.test(entry.name)) out.push(full);
    }
    return out;
}

const rootFiles = (match) =>
    readdirSync(projectRoot, { withFileTypes: true })
        .filter((e) => e.isFile() && match.test(e.name))
        .map((e) => join(projectRoot, e.name));

/** App-Seiten: Root und tools/ (landing/ ist die Marketing-Seite mit eigenem Deployment). */
const appHtml = [...rootFiles(/\.html$/), ...walk(join(projectRoot, 'tools'), /\.html$/)];
const appCode = [
    ...walk(join(projectRoot, 'src'), /\.(js|mjs|css|html)$/),
    ...rootFiles(/\.(js|css)$/).filter((f) => !/\.config\.js$/.test(f))
];

describe('Browser-Vertrauenskette', () => {
    it('findet die App-Seiten und den Quellcode', () => {
        expect(appHtml.length).toBeGreaterThan(100);
        expect(appCode.length).toBeGreaterThan(300);
    });

    it('lädt nichts von fremden CDNs', () => {
        const cdn = /esm\.sh|cdnjs\.|cdn\.jsdelivr|unpkg\.com|cdn\.skypack|ga\.jspm/i;
        const hits = [...appHtml, ...appCode].filter((f) => cdn.test(readFileSync(f, 'utf8'))).map(rel);
        expect(hits, 'Bibliotheken über npm + scripts/vendor-libs.mjs einbinden').toEqual([]);
    });

    it('App-Seiten haben keine Inline-Skripte, Inline-Handler oder fremde Skript-Hosts', () => {
        const failures = appHtml
            .map((f) => ({ file: rel(f), problems: findCspViolations(readFileSync(f, 'utf8'), { source: true }) }))
            .filter((r) => r.problems.length);
        expect(failures, 'Skripte als Datei einbinden (src/pages/), Klicks per addEventListener oder data-ms365-call').toEqual([]);
    });

    it('Quellcode nutzt kein eval, new Function oder Inline-Handler in HTML-Vorlagen', () => {
        const patterns = [
            /\bnew Function\s*\(/,
            /\beval\s*\(/,
            /set(?:Timeout|Interval)\s*\(\s*['"`]/,
            // Handler-Attribut in einer HTML-Vorlage, z. B. <button onclick="…">
            /<[a-zA-Z][^<>\n]{0,300}\son[a-z]+\s*=\s*\\?["']/,
            /(?:href|src)\s*=\s*\\?["']\s*javascript:/i
        ];
        const hits = appCode
            .filter((f) => /\.(js|mjs)$/.test(f))
            .filter((f) => {
                const text = readFileSync(f, 'utf8');
                return patterns.some((re) => re.test(text));
            })
            .map(rel);
        expect(hits).toEqual([]);
    });

    it('CSP erlaubt Skripte nur von der eigenen Origin', () => {
        expect(CSP_DIRECTIVES['script-src']).toEqual(["'self'"]);
        expect(CSP_DIRECTIVES['object-src']).toEqual(["'none'"]);
        expect(CSP_DIRECTIVES['base-uri']).toEqual(["'self'"]);
        expect(CSP_DIRECTIVES['default-src']).toEqual(["'self'"]);
        const prod = buildCspContent();
        expect(prod).not.toMatch(/unsafe-eval|ws:\/\//);
        expect(prod).not.toMatch(/script-src[^;]*unsafe-inline/);
        expect(buildCspContent({ dev: true })).toContain('ws://localhost:*');
    });

    it('injectCspMeta setzt das Meta-Tag vor das erste Skript und bleibt idempotent', () => {
        const html = '<!DOCTYPE html><html><head>\n<script src="a.js"></script><title>x</title></head><body></body></html>';
        const once = injectCspMeta(html);
        expect(findCspViolations(once)).toEqual([]);
        expect(injectCspMeta(once)).toBe(once);
        expect(once.match(/Content-Security-Policy/g)).toHaveLength(1);
    });

    it('findCspViolations erkennt Verstöße', () => {
        const page = (body) => injectCspMeta(`<html><head></head><body>${body}</body></html>`);
        expect(findCspViolations(page('<script>alert(1)</script>'))).toHaveLength(1);
        expect(findCspViolations(page('<button onclick="go()">x</button>'))).toHaveLength(1);
        expect(findCspViolations(page('<script src="https://cdn.example/x.js"></script>'))).toHaveLength(1);
        expect(findCspViolations(page('<script type="application/json">{"a":1}</script>'))).toEqual([]);
        expect(findCspViolations('<html><head></head><body></body></html>')).toEqual(['CSP-Meta-Tag fehlt']);
    });
});

describe('Deploy und GitHub Actions', () => {
    const workflowDir = join(projectRoot, '.github/workflows');
    const workflows = readdirSync(workflowDir)
        .filter((n) => /\.ya?ml$/.test(n))
        .map((n) => ({ name: n, text: readFileSync(join(workflowDir, n), 'utf8') }));

    it('pinnt jede Action auf einen Commit-SHA', () => {
        const unpinned = [];
        for (const wf of workflows) {
            for (const m of wf.text.matchAll(/^\s*(?:-\s*)?uses:\s*(\S+)/gm)) {
                if (m[1].startsWith('./')) continue;
                if (!/@[0-9a-f]{40}$/.test(m[1])) unpinned.push(`${wf.name}: ${m[1]}`);
            }
        }
        expect(workflows.length).toBeGreaterThanOrEqual(5);
        expect(unpinned).toEqual([]);
    });

    it('überträgt nicht per FTP im Klartext', () => {
        for (const wf of workflows) {
            expect(wf.text, wf.name).not.toMatch(/protocol:\s*ftp\s*$/m);
            expect(wf.text, wf.name).not.toMatch(/FTP-Deploy-Action/);
            // Keine Längen/Fragmente von Secrets ins (öffentliche) Log.
            expect(wf.text, wf.name).not.toMatch(/\$\{#FTP_/);
        }
    });

    const base = { host: 'ftp.example.org', user: 'deploy', remoteDir: '/app/', localDir: 'dist' };

    it('FTPS erzwingt TLS mit Zertifikatsprüfung und löscht nichts', () => {
        const script = buildLftpScript(base);
        expect(script).toContain('set ftp:ssl-force true');
        expect(script).toContain('set ftp:ssl-protect-data true');
        expect(script).toContain('set ssl:verify-certificate true');
        expect(script).toContain('open --env-password -u "deploy" "ftp://ftp.example.org"');
        expect(script).toContain('mirror --reverse --verbose=1 --parallel=2 --no-perms "dist" "/app"');
        expect(script).not.toMatch(/--delete/);
    });

    it('lehnt unverschlüsseltes FTP und unvollständige Angaben ab', () => {
        expect(() => buildLftpScript({ ...base, protocol: 'ftp' })).toThrow(/kein unverschlüsseltes FTP/);
        expect(() => buildLftpScript({ ...base, host: 'ftp://ftp.example.org' })).toThrow(/DEPLOY_HOST/);
        expect(() => buildLftpScript({ ...base, remoteDir: '' })).toThrow(/DEPLOY_REMOTE_DIR/);
        expect(() => buildLftpScript({ ...base, protocol: 'sftp' })).toThrow(/DEPLOY_SSH_KNOWN_HOSTS/);
    });

    it('SFTP prüft den Host-Key', () => {
        const script = buildLftpScript({ ...base, protocol: 'sftp', knownHostsFile: '/tmp/kh', port: '22' });
        expect(script).toContain('StrictHostKeyChecking=yes');
        expect(script).toContain('UserKnownHostsFile=/tmp/kh');
        expect(script).toContain('open --env-password -u "deploy" -p 22 "sftp://ftp.example.org"');
    });

    it('normalisiert den Zielordner für lftp mirror', () => {
        expect(normalizeRemoteDir('/app/')).toBe('/app');
        expect(normalizeRemoteDir('/')).toBe('/');
        expect(normalizeRemoteDir('httpdocs/')).toBe('httpdocs');
        expect(normalizeRemoteDir('')).toBe('');
    });
});
