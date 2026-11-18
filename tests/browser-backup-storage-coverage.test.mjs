import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

function walk(dir, out = []) {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
        const p = join(dir, entry.name);
        if (entry.isDirectory()) {
            if (/node_modules|dist|\.git/.test(p)) continue;
            walk(p, out);
        } else if (/\.(js|mjs|ts)$/.test(entry.name)) {
            out.push(p);
        }
    }
    return out;
}

function createMemoryStorage() {
    const map = new Map();
    return {
        get length() {
            return map.size;
        },
        key(i) {
            return Array.from(map.keys())[i] ?? null;
        },
        getItem(k) {
            return map.has(k) ? map.get(k) : null;
        },
        setItem(k, v) {
            map.set(String(k), String(v));
        },
        removeItem(k) {
            map.delete(k);
        }
    };
}

function loadBackupApi() {
    const full = join(projectRoot, 'src/shared/browser-backup.js');
    const code = readFileSync(full, 'utf8');
    const sandbox = {
        console,
        document: {
            readyState: 'complete',
            getElementById() {
                return null;
            },
            querySelectorAll() {
                return [];
            },
            addEventListener() {}
        },
        localStorage: createMemoryStorage(),
        sessionStorage: createMemoryStorage()
    };
    sandbox.window = sandbox;
    createContext(sandbox);
    runInContext(code, sandbox, { filename: full });
    return sandbox.ms365BrowserBackup;
}

/**
 * Keys, die absichtlich nicht im Inventar-Katalog stehen (Session-Skip / reine Runtime-Marker).
 */
const CATALOG_OPTIONAL = new Set([
    'ms365-access-granted-v1',
    'ms365-admin-access-granted-v1',
    'ms365-post-login-url',
    'ms365-spo-compare-session-v1',
    'ms365-spo-auto-pull-done-v1',
    'ms365-admin-license-packages-draft-v1',
    'ms365-dashboard-portal-nav-v1',
    'ms365-schulstruktur-sync-match-v1',
    'ms365-schulstruktur-sync-tenant-cache-v1',
    'ms365-spo-config-bundle-v1'
]);

describe('browser-backup storage coverage', () => {
    const api = loadBackupApi();

    it('sichert jedes ms365-/webuntis-localStorage per Präfix und Restore 1:1', () => {
        const store = createMemoryStorage();
        const keys = [
            'ms365-schooltool-data-v2',
            'ms365-brand-new-tool-settings-v1',
            'webuntis-teams-creator-state-v1',
            'ms365-wahlfach-zuteilung-settings-v1'
        ];
        keys.forEach((k, i) => store.setItem(k, JSON.stringify({ i, k })));
        store.setItem('msal.cache', 'secret');
        store.setItem('other-app', 'nope');
        const payload = api.buildBackup(store);
        keys.forEach((k) => {
            expect(payload.localStorage[k]).toEqual({ i: keys.indexOf(k), k });
        });
        expect(payload.localStorage['msal.cache']).toBeUndefined();
        expect(payload.localStorage['other-app']).toBeUndefined();
        const target = createMemoryStorage({ 'ms365-stale-v1': '1' });
        api.applyBackup(payload, target);
        keys.forEach((k) => {
            expect(JSON.parse(target.getItem(k)).k).toBe(k);
        });
        expect(target.getItem('ms365-stale-v1')).toBeNull();
        expect(target.getItem('msal.cache')).toBeNull();
    });

    it('Catalog enthält alle bekannten Schul-App-Storage-Keys aus dem Quellcode', () => {
        const files = walk(join(projectRoot, 'src'));
        const found = new Set();
        const assignRe =
            /(?:STORAGE_KEY|SETTINGS_KEY|INBOX_KEY|SESSION_KEY|FEED_META_KEY|ROLE_KEY|SITE_KEY|CACHE_KEY|FLOW_DONE_KEY|PERMS_STORAGE_KEY|STEP_STORAGE_KEY|KEY)\s*=\s*['"]((?:ms365|webuntis)-[^'"]+)['"]/g;
        const lsRe = /localStorage\.(?:get|set|remove)Item\(\s*['"]((?:ms365|webuntis)-[^'"]+)['"]/g;
        for (const file of files) {
            const text = readFileSync(file, 'utf8');
            for (const re of [assignRe, lsRe]) {
                re.lastIndex = 0;
                let m;
                while ((m = re.exec(text))) found.add(m[1]);
            }
        }

        const catalog = new Set();
        api.STORAGE_CATALOG.forEach((cat) => cat.keys.forEach((k) => catalog.add(k)));

        const missing = [...found]
            .filter((k) => api.isAppStorageKey(k))
            .filter((k) => !k.includes('attachments'))
            .filter((k) => !CATALOG_OPTIONAL.has(k))
            .filter((k) => !catalog.has(k))
            .sort();

        expect(missing, 'Neue Storage-Keys bitte in STORAGE_CATALOG aufnehmen').toEqual([]);
    });

    it('kennt die vier Anhang-IndexedDBs', () => {
        expect(api.INDEXED_DB_CATALOG).toHaveLength(4);
        expect(api.VERSION).toBeGreaterThanOrEqual(7);
    });
});
