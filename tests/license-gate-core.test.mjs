import { createRequire } from 'node:module';
import { describe, expect, it, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const __dirname = dirname(fileURLToPath(import.meta.url));
const src = readFileSync(join(__dirname, '../src/shared/license-gate-core.js'), 'utf8');

function loadCore(sessionStore) {
    const storage = {
        _data: { ...(sessionStore || {}) },
        getItem(k) {
            return Object.prototype.hasOwnProperty.call(this._data, k) ? this._data[k] : null;
        },
        setItem(k, v) {
            this._data[k] = String(v);
        },
        removeItem(k) {
            delete this._data[k];
        }
    };
    const sandbox = {
        window: {},
        sessionStorage: storage,
        location: { pathname: '/index.html' }
    };
    sandbox.globalThis = sandbox;
    sandbox.window = sandbox;
    vm.runInNewContext(src, sandbox);
    return { core: sandbox.ms365LicenseGateCore, storage };
}

describe('license-gate-core', () => {
    it('erkennt Exempt-Pfade', () => {
        const { core } = loadCore();
        expect(core.isExemptPath('/admin.html')).toBe(true);
        expect(core.isExemptPath('/tools/license-backend-setup.html')).toBe(true);
        expect(core.isExemptPath('/welcome.html')).toBe(true);
        expect(core.isExemptPath('/index.html')).toBe(false);
        expect(core.isExemptPath('/tools/kursteams.html')).toBe(false);
    });

    it('cached und matched Account', () => {
        const { core, storage } = loadCore();
        core.writeCache(
            { allowed: true, reason: 'ok', message: 'ok', tenantId: 't1', license: { schoolName: 'A' } },
            'user@school.at',
            storage
        );
        const cached = core.readCache(storage);
        expect(cached.allowed).toBe(true);
        expect(core.cacheMatchesAccount(cached, 'user@school.at')).toBe(true);
        expect(core.cacheMatchesAccount(cached, 'other@school.at')).toBe(false);
        core.clearCache(storage);
        expect(core.readCache(storage)).toBe(null);
    });
});
