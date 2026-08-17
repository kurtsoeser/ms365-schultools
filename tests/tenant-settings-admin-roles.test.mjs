import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';
import { describe, expect, it, beforeEach } from 'vitest';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');

function loadTenantSettingsCore(store) {
    const full = join(projectRoot, 'src/shared/tenant-settings-core.js');
    const code = readFileSync(full, 'utf8');
    const sandbox = { console };
    sandbox.window = sandbox;
    sandbox.localStorage = {
        getItem(k) {
            return store.has(k) ? store.get(k) : null;
        },
        setItem(k, v) {
            store.set(k, String(v));
        },
        removeItem(k) {
            store.delete(k);
        }
    };
    createContext(sandbox);
    runInContext(code, sandbox, { filename: full });
    return sandbox;
}

describe('tenant-settings adminRoles', () => {
    let store;

    beforeEach(() => {
        store = new Map();
    });

    it('leitet den Rollenkatalog aus Personen ab', () => {
        const ctx = loadTenantSettingsCore(store);
        const saved = ctx.ms365TenantSettingsSave({
            domain: 'schule.at',
            admin: [{ role: 'direktion', name: '', email: '' }]
        });
        expect(saved.adminRoles.length).toBe(1);
        expect(saved.adminRoles[0].code).toBe('DIREKTION');
        expect(saved.adminRoles[0].name).toBe('direktion');
    });

    it('parseAdminRolesLines liest Kürzel;Bezeichnung', () => {
        const ctx = loadTenantSettingsCore(store);
        const rows = ctx.ms365TenantSettingsParseAdminRolesLines('SEKRETARIAT;Sekretariat\nSchularzt');
        expect(rows[0]).toEqual({ code: 'SEKRETARIAT', name: 'Sekretariat' });
        expect(rows[1].name).toBe('Schularzt');
        expect(rows[1].code).toBe('SCHULARZT');
    });

    it('mehrere Personen mit derselben Rolle bleiben erhalten', () => {
        const ctx = loadTenantSettingsCore(store);
        const saved = ctx.ms365TenantSettingsSave({
            admin: [
                { role: 'Sekretariat', name: 'Anna', email: 'a@x.at' },
                { role: 'Sekretariat', name: 'Ben', email: 'b@x.at' }
            ]
        });
        expect(saved.admin.length).toBe(2);
        expect(saved.adminRoles.some((r) => r.name === 'Sekretariat')).toBe(true);
    });

    it('renameAdminRole aktualisiert Katalog und Personen', () => {
        const ctx = loadTenantSettingsCore(store);
        const out = ctx.ms365TenantSettingsRenameAdminRole(
            [{ code: 'SEK', name: 'Sekretariat' }],
            [{ role: 'Sekretariat', name: 'Anna', email: 'a@x.at' }],
            'Sekretariat',
            'Schulsekretariat'
        );
        expect(out.roles[0].name).toBe('Schulsekretariat');
        expect(out.admin[0].role).toBe('Schulsekretariat');
    });
});

describe('Schülerlisten-Parser', () => {
    it('parseStudentsLines behält leere Klasse (Import ohne Klassenzuordnung)', () => {
        const ctx = loadTenantSettingsCore(new Map());
        const rows = ctx.ms365TenantSettingsParseStudentsLines(';Lisa Beispiel;lisa@schule.at\n1A;Tom;tom@schule.at');
        expect(rows[0]).toEqual({ klasse: '', name: 'Lisa Beispiel', email: 'lisa@schule.at' });
        expect(rows[1].klasse).toBe('1A');
    });
});

describe('Stammdaten Verwaltung UI', () => {
    it('enthält Rollenkatalog und Personenliste', () => {
        const html = readFileSync(join(projectRoot, 'tenant.html'), 'utf8');
        expect(html).toContain('id="tenantAdminRoleLines"');
        expect(html).toContain('id="tenantAdminRolesTableBody"');
        expect(html).toContain('id="tenantAdminRoleAddRow"');
        expect(html).toContain('id="tenantAdminRolesDefaults"');
        expect(html).toContain('Personen (eine Zeile pro Person)');
    });
});
