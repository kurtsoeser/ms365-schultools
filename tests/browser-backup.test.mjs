import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

function createMemoryStorage(initial) {
    const map = new Map(Object.entries(initial || {}));
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

function loadBackup(store) {
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
        }
    };
    sandbox.window = sandbox;
    sandbox.localStorage = store;
    createContext(sandbox);
    runInContext(code, sandbox, { filename: full });
    return sandbox;
}

describe('browser-backup', () => {
    let store;
    let api;

    beforeEach(() => {
        store = createMemoryStorage({
            'ms365-schooltool-data-v2': JSON.stringify({ version: 3, core: { domain: 'schule.at' } }),
            'ms365-dashboard-favorites-v1': JSON.stringify(['kursteams']),
            'webuntis-teams-creator-state-v1': JSON.stringify({ step: 2 }),
            'ms365-dashboard-expert-open-v1': '1',
            'msal.token.keys.abc': 'secret',
            'acc-login.microsoftonline.com-idtoken-x': 'token',
            'unrelated-other-app': 'nope'
        });
        api = loadBackup(store).ms365BrowserBackup;
    });

    it('erkennt App-Keys und blendet MSAL/fremde Keys aus', () => {
        expect(api.isAppStorageKey('ms365-schooltool-data-v2')).toBe(true);
        expect(api.isAppStorageKey('webuntis-teams-creator-state-v1')).toBe(true);
        expect(api.isAppStorageKey('msal.token.keys.abc')).toBe(false);
        expect(api.isAuthKey('acc-login.microsoftonline.com-idtoken-x')).toBe(true);
        expect(api.isAppStorageKey('unrelated-other-app')).toBe(false);
    });

    it('sammelt nur App-Werte und erhält Nicht-JSON-Strings', () => {
        const collected = api.collectLocalStorage(store);
        expect(Object.keys(collected).sort()).toEqual([
            'ms365-dashboard-expert-open-v1',
            'ms365-dashboard-favorites-v1',
            'ms365-schooltool-data-v2',
            'webuntis-teams-creator-state-v1'
        ]);
        expect(collected['ms365-dashboard-expert-open-v1']).toBe('1');
        expect(collected['ms365-schooltool-data-v2']).toEqual({ version: 3, core: { domain: 'schule.at' } });
        expect(collected['msal.token.keys.abc']).toBeUndefined();
        expect(collected['unrelated-other-app']).toBeUndefined();
    });

    it('baut ein erkennbares Backup-Paket', () => {
        const now = new Date(2026, 7, 17, 15, 0, 0);
        const payload = api.buildBackup(store, now);
        expect(api.isBackupPayload(payload)).toBe(true);
        expect(payload.kind).toBe('ms365-browser-backup-v1');
        expect(payload.version).toBe(1);
        expect(payload.exportedAt).toBe(now.toISOString());
        expect(payload.keyCount).toBe(4);
        expect(api.backupFilename(now)).toBe('ms365-browser-backup-2026-08-17.json');
    });

    it('stellt ein Backup 1:1 wieder her und räumt Ziel-Reste weg', () => {
        const payload = api.buildBackup(store);
        const target = createMemoryStorage({
            'ms365-jahrgang-state-v1': '{"old":true}',
            'ms365-dashboard-tab-v1': 'sharepoint',
            'msal.cache': 'keep-me'
        });
        const result = api.applyBackup(payload, target);
        expect(target.getItem('ms365-jahrgang-state-v1')).toBeNull();
        expect(target.getItem('ms365-dashboard-tab-v1')).toBeNull();
        expect(target.getItem('msal.cache')).toBe('keep-me');
        expect(JSON.parse(target.getItem('ms365-schooltool-data-v2')).core.domain).toBe('schule.at');
        expect(target.getItem('ms365-dashboard-expert-open-v1')).toBe('1');
        expect(target.getItem('webuntis-teams-creator-state-v1')).toContain('"step":2');
        expect(result.written).toContain('ms365-schooltool-data-v2');
        expect(result.removed).toContain('ms365-jahrgang-state-v1');
    });

    it('schreibt MSAL-Keys aus einer Datei nicht zurück', () => {
        const target = createMemoryStorage();
        const result = api.applyBackup(
            {
                kind: api.KIND,
                version: 1,
                localStorage: {
                    'ms365-school-email-domain-v1': 'schule.at',
                    'msal.token.keys.abc': 'stolen'
                }
            },
            target
        );
        expect(target.getItem('ms365-school-email-domain-v1')).toBe('schule.at');
        expect(target.getItem('msal.token.keys.abc')).toBeNull();
        expect(result.skipped).toContain('msal.token.keys.abc');
    });

    it('erkennt Legacy-Schuldaten und lehnt unbekanntes JSON ab', () => {
        expect(api.isLegacyAppDataPayload({ version: 3, core: {}, structure: {}, match: {} })).toBe(true);
        expect(api.isBackupPayload({ version: 3, core: {}, structure: {}, match: {} })).toBe(false);
        expect(() => api.importPayload({ foo: 1 }, store)).toThrow(/Unbekanntes JSON-Format/);
        expect(() => api.applyBackup({ kind: 'nope', localStorage: {} }, store)).toThrow(/Kein gültiges Browser-Backup/);
    });

    it('importPayload spielt ein Backup auf den übergebenen Store', () => {
        const payload = api.buildBackup(store);
        const target = createMemoryStorage({ 'ms365-arge-state-v2': '{}' });
        const imported = api.importPayload(payload, target);
        expect(imported.mode).toBe('backup');
        expect(target.getItem('ms365-arge-state-v2')).toBeNull();
        expect(target.getItem('ms365-dashboard-favorites-v1')).toContain('kursteams');
    });
});

describe('browser-backup Seiten', () => {
    const pages = [
        { file: 'index.html', scriptPrefix: 'src/shared/browser-backup.js' },
        { file: 'tenant.html', scriptPrefix: 'src/shared/browser-backup.js' },
        { file: 'einrichtung.html', scriptPrefix: 'src/shared/browser-backup.js' }
    ];

    it.each(pages)('$file bindet Backup-Skript und Steuerelemente ein', ({ file, scriptPrefix }) => {
        const html = readFileSync(join(projectRoot, file), 'utf8');
        expect(html).toContain(scriptPrefix);
        expect(html).toContain('id="browserBackupExport"');
        expect(html).toContain('id="browserBackupImportFile"');
    });
});
