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

function loadRegisterAndBackup() {
    const regPath = join(projectRoot, 'src/shared/storage-register.js');
    const backupPath = join(projectRoot, 'src/shared/browser-backup.js');
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
    runInContext(readFileSync(regPath, 'utf8'), sandbox, { filename: regPath });
    runInContext(readFileSync(backupPath, 'utf8'), sandbox, { filename: backupPath });
    return {
        register: sandbox.ms365StorageRegister,
        api: sandbox.ms365BrowserBackup
    };
}

/**
 * Keys, die aus dem Quellcode erkannt werden, aber absichtlich nur als
 * Präfix-Sweep (nicht als Katalog-Pflicht) gelten – derzeit leer: alles
 * Bekannte steht im Storage-Register.
 */
const CATALOG_OPTIONAL = new Set([]);

describe('storage-register + browser-backup coverage', () => {
    const { register, api } = loadRegisterAndBackup();

    it('lädt das Storage-Register und hängt den Katalog am Backup-API', () => {
        expect(register).toBeTruthy();
        expect(api).toBeTruthy();
        expect(api.STORAGE_CATALOG).toBe(register.STORAGE_CATALOG);
        expect(api.INDEXED_DB_CATALOG).toBe(register.INDEXED_DB_CATALOG);
        expect(register.allRegisteredKeys().length).toBeGreaterThan(100);
        expect(register.duplicateKeys(), 'keine doppelten Keys im Register').toEqual([]);
    });

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
        const target = createMemoryStorage();
        target.setItem('ms365-stale-v1', '1');
        api.applyBackup(payload, target);
        keys.forEach((k) => {
            expect(JSON.parse(target.getItem(k)).k).toBe(k);
        });
        expect(target.getItem('ms365-stale-v1')).toBeNull();
        expect(target.getItem('msal.cache')).toBeNull();
    });

    it('Register enthält alle bekannten Schul-App-Storage-Keys aus dem Quellcode', () => {
        const files = walk(join(projectRoot, 'src'));
        const found = new Set();
        const assignRe =
            /(?:STORAGE_KEY|SETTINGS_KEY|INBOX_KEY|SESSION_KEY|FEED_META_KEY|ROLE_KEY|SITE_KEY|CACHE_KEY|FLOW_DONE_KEY|PERMS_STORAGE_KEY|STEP_STORAGE_KEY|LAST_KEY|LAST_AT_KEY|PULL_DONE_KEY|LAST_PUSH_KEY|KEY)\s*=\s*['"]((?:ms365|webuntis)-[^'"]+)['"]/g;
        const lsRe = /localStorage\.(?:get|set|remove)Item\(\s*['"]((?:ms365|webuntis)-[^'"]+)['"]/g;
        for (const file of files) {
            if (/storage-register\.js$|browser-backup\.js$/.test(file.replace(/\\/g, '/'))) continue;
            const text = readFileSync(file, 'utf8');
            for (const re of [assignRe, lsRe]) {
                re.lastIndex = 0;
                let m;
                while ((m = re.exec(text))) found.add(m[1]);
            }
        }

        const catalog = new Set(register.allRegisteredKeys());

        const missing = [...found]
            .filter((k) => api.isAppStorageKey(k) || register.isAppStorageKeyPrefix(k))
            .filter((k) => !k.includes('attachments'))
            .filter((k) => !CATALOG_OPTIONAL.has(k))
            .filter((k) => !catalog.has(k))
            .sort();

        expect(missing, 'Neue Storage-Keys bitte in src/shared/storage-register.js aufnehmen').toEqual(
            []
        );
    });

    it('kennt die vier Anhang-IndexedDBs', () => {
        expect(api.INDEXED_DB_CATALOG).toHaveLength(4);
        expect(api.VERSION).toBeGreaterThanOrEqual(7);
    });

    it('clearAllAppData entfernt alle Präfix-Keys inkl. Register-Einträge', async () => {
        const store = createMemoryStorage();
        store.setItem('ms365-schooltool-data-v2', '{}');
        store.setItem('ms365-theme-v1', '"light"');
        store.setItem('msal.cache', 'keep');
        const result = await api.clearAllAppData({ storage: store, sessionStorage: createMemoryStorage() });
        expect(result.removed).toBeGreaterThanOrEqual(2);
        expect(store.getItem('ms365-schooltool-data-v2')).toBeNull();
        expect(store.getItem('msal.cache')).toBe('keep');
    });
});
