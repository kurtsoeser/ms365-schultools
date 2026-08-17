import { describe, it, expect } from 'vitest';
import {
    SW_ADMIN_DEFAULT_ROLES,
    getAdminDisplayRowsFromSettings,
    migrateAdminRowDefaultKey,
    normCode,
    normStr,
    randomTempPassword,
    resolveAdminSlotFromRow,
    defaultAdminRoleCatalog,
    normalizeAdminRoleCatalog,
    renameAdminRole,
    personMatchesAdminRole
} from '../src/shared/setup-wizard-admin-model.js';

describe('setup-wizard-admin-model', () => {
    it('getAdminDisplayRowsFromSettings: leere Admin-Liste → Standardrollen', () => {
        const rows = getAdminDisplayRowsFromSettings({});
        expect(rows.length).toBe(SW_ADMIN_DEFAULT_ROLES.length);
        expect(rows[0]).toEqual({ defaultKey: 'Direktion', role: 'Direktion', name: '', email: '' });
    });

    it('migrateAdminRowDefaultKey setzt defaultKey aus Rolle', () => {
        const m = migrateAdminRowDefaultKey({ role: 'bibliothek', name: 'x', email: '' });
        expect(m.defaultKey).toBe('Bibliothek');
    });

    it('resolveAdminSlotFromRow und normCode', () => {
        expect(resolveAdminSlotFromRow({ defaultKey: 'sekretariat' })).toBe('Sekretariat');
        expect(normCode('  ab12  ')).toBe('AB12');
        expect(normStr('  x  ')).toBe('x');
    });

    it('randomTempPassword: Länge und Zeichenklassen', () => {
        const pwd = randomTempPassword();
        expect(pwd.length).toBeGreaterThanOrEqual(16);
        expect(/[A-Z]/.test(pwd)).toBe(true);
        expect(/[a-z]/.test(pwd)).toBe(true);
        expect(/[0-9]/.test(pwd)).toBe(true);
    });

    it('defaultAdminRoleCatalog enthält Direktion und Sekretariat', () => {
        const cat = defaultAdminRoleCatalog();
        expect(cat.find((x) => x.name === 'Direktion').code).toBe('DIREKTION');
        expect(cat.find((x) => x.name === 'IT-Support').code).toBe('IT-SUPPORT');
        expect(cat.length).toBe(SW_ADMIN_DEFAULT_ROLES.length);
    });

    it('normalizeAdminRoleCatalog ergänzt Rollen aus Personen', () => {
        const roles = normalizeAdminRoleCatalog([], [{ role: 'Schularzt', name: 'Dr. A', email: 'a@x.at' }]);
        expect(roles.some((r) => r.name === 'Schularzt')).toBe(true);
        expect(roles[0].code).toBe('SCHULARZT');
    });

    it('renameAdminRole zieht Personen mit', () => {
        const r = renameAdminRole(
            [{ code: 'SEK', name: 'Sekretariat' }],
            [{ role: 'Sekretariat', name: 'Anna', email: 'a@x.at' }],
            'Sekretariat',
            'Schulsekretariat'
        );
        expect(r.roles[0].name).toBe('Schulsekretariat');
        expect(r.admin[0].role).toBe('Schulsekretariat');
    });

    it('personMatchesAdminRole erkennt Name und Kürzel', () => {
        const role = { code: 'DIREKTION', name: 'Direktion' };
        expect(personMatchesAdminRole({ role: 'Direktion' }, role)).toBe(true);
        expect(personMatchesAdminRole({ role: 'DIREKTION' }, role)).toBe(true);
        expect(personMatchesAdminRole({ role: 'Sekretariat' }, role)).toBe(false);
    });
});
