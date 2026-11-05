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
        // Leere Personenzeile (Entwurf „+ Person“) bleibt der Rolle zugeordnet.
        expect(saved.admin).toHaveLength(1);
        expect(saved.admin[0]).toMatchObject({ role: 'direktion', name: '', email: '' });
    });

    it('behält leere Personenzeilen beim Speichern (Verwaltung + Person)', () => {
        const ctx = loadTenantSettingsCore(store);
        const saved = ctx.ms365TenantSettingsSave({
            adminRoles: [{ code: 'FACILITIESORGANISATION', name: 'Facilities/Organisation' }],
            admin: [{ role: 'Facilities/Organisation', name: '', email: '' }]
        });
        expect(saved.admin).toHaveLength(1);
        expect(saved.admin[0].role).toBe('Facilities/Organisation');
        expect(saved.administration.find((g) => g.name === 'Facilities/Organisation')?.people).toHaveLength(1);
    });

    it('schreibt bei ausgeschaltetem Spiegel nicht mehr nach ms365-tenant-settings-v1', () => {
        const ctx = loadTenantSettingsCore(store);
        let wroteV2 = false;
        ctx.window.ms365TenantStorageMirror = {
            isV1MirrorWriteEnabled: function () {
                return false;
            },
            writeTenantSettingsV1Mirror: function () {
                throw new Error('v1-Spiegel sollte nicht aufgerufen werden');
            }
        };
        ctx.window.ms365AppDataV2 = {
            setCoreFromTenantSettings: function () {
                wroteV2 = true;
            },
            getContainer: function () {
                return { core: {} };
            }
        };
        const saved = ctx.ms365TenantSettingsSave({
            domain: 'demo.at',
            teachers: [{ code: 'MU', name: 'Max', email: 'max@demo.at' }]
        });
        expect(saved.teachers[0].email).toBe('max@demo.at');
        expect(wroteV2).toBe(true);
        expect(store.get('ms365-tenant-settings-v1')).toBeUndefined();
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

    it('persistiert seatMode und Organigramm in Stammdaten', () => {
        const ctx = loadTenantSettingsCore(store);
        /** @type {{ core: Record<string, unknown> }} */
        const v2 = { core: {} };
        ctx.window.ms365AppDataV2 = {
            setCoreFromTenantSettings: function (normalized) {
                v2.core = { ...normalized };
            },
            getContainer: function () {
                return v2;
            }
        };
        const saved = ctx.ms365TenantSettingsSave({
            domain: 'schule.at',
            adminRoles: [{ code: 'SEK', name: 'Sekretariat', seatMode: 'multi' }],
            admin: [{ role: 'Sekretariat', name: 'Anna', email: 'a@schule.at' }],
            verwaltungOrganigramm: {
                version: 1,
                nodes: [
                    { id: 'sec_lead', type: 'group', label: 'Schulleitung', parentId: null, order: 0, tone: 'lead' },
                    {
                        id: 'role_sek',
                        type: 'role',
                        label: 'Sekretariat',
                        parentId: 'sec_lead',
                        order: 1,
                        roleCode: 'SEK'
                    }
                ]
            }
        });
        expect(saved.adminRoles.find((r) => r.code === 'SEK')?.seatMode).toBe('multi');
        expect(saved.verwaltungOrganigramm.nodes).toHaveLength(2);
        expect(saved.verwaltungOrganigramm.nodes.find((n) => n.type === 'role')?.roleCode).toBe('SEK');

        const again = ctx.ms365TenantSettingsSave({
            domain: 'schule.at',
            teachers: [{ code: 'MU', name: 'Max', email: 'max@schule.at' }]
        });
        expect(again.verwaltungOrganigramm.nodes).toHaveLength(2);
        expect(again.verwaltungOrganigramm.nodes.find((n) => n.roleCode === 'SEK')?.label).toBe('Sekretariat');
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

describe('tenant-settings ARGE-Liste', () => {
    let store;

    beforeEach(() => {
        store = new Map();
    });

    it('parst ARGE-Leitung und leere Zeilen', () => {
        const ctx = loadTenantSettingsCore(store);
        const parse = ctx.ms365TenantSettingsParseArgesLines;
        const rows = parse('ARGE MATHE;DEMO Mathe;MATHE;Anna Leitung;anna@schule.at\n;;;;');
        expect(rows).toHaveLength(2);
        expect(rows[0].headEmail).toBe('anna@schule.at');
        expect(rows[0].subjects).toEqual(['MATHE']);
        expect(rows[1].code).toBe('');
    });
});

describe('tenant-settings Klassenliste', () => {
    let store;

    beforeEach(() => {
        store = new Map();
    });

    it('behält leere Zeilen für „+ Zeile“ im Editor', () => {
        const ctx = loadTenantSettingsCore(store);
        const parse = ctx.ms365TenantSettingsParseClassesLines;
        const rows = parse('1A;2030;Klasse 1A\n;;;;');
        expect(rows).toHaveLength(2);
        expect(rows[0].code).toBe('1A');
        expect(rows[1].code).toBe('');
        expect(rows[1].name).toBe('');
    });
});
