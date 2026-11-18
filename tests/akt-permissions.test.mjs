import { describe, it, expect } from 'vitest';
import {
    normalizeAktPermissionsConfig,
    aktEntraGroupsConfigured,
    AKT_LIST_PERM_PROFILES,
    aktPackageListTargets
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-permissions.js';
import {
    AKT_SETUP_GROUP_FIELDS,
    AKT_STAMMDATEN_GROUP_ROLES
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-permissions-ui.js';

describe('schulaktivitaeten permissions', () => {
    it('normalisiert Gruppen und Einzelpersonen', () => {
        const c = normalizeAktPermissionsConfig({
            groupAdmin: 'Direktion',
            groupAdminId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
            adminUsers: [{ mail: 'dir@schule.at', displayName: 'Direktor' }],
            lehrerUsers: [{ email: 'lehrer@schule.at' }]
        });
        expect(c.groupAdmin).toBe('Direktion');
        expect(c.adminUsers[0].mail).toBe('dir@schule.at');
        expect(c.lehrerUsers[0].mail).toBe('lehrer@schule.at');
        expect(aktEntraGroupsConfigured(c)).toBe(true);
        expect(aktEntraGroupsConfigured({})).toBe(false);
    });

    it('Lehrer schreibt Anträge und liest Stammdaten', () => {
        expect(AKT_LIST_PERM_PROFILES.aktivitaeten.lehrer).toBe('contribute');
        expect(AKT_LIST_PERM_PROFILES.klassen.lehrer).toBe('read');
        expect(AKT_LIST_PERM_PROFILES.lehrer.lehrer).toBe('read');
        expect(AKT_LIST_PERM_PROFILES.regelwerk.lehrer).toBe('read');
        expect(AKT_LIST_PERM_PROFILES.aktivitaeten.admin).toBe('fullControl');
    });

    it('Paket-Ziele enthalten Anträge und Stammdaten', () => {
        const keys = aktPackageListTargets().map((t) => t.profileKey);
        expect(keys).toEqual(['aktivitaeten', 'regelwerk', 'lehrer', 'klassen']);
    });

    it('Genehmigung und Einreicher sind beide als Gruppen wählbar', () => {
        expect(AKT_STAMMDATEN_GROUP_ROLES.size).toBe(0);
        const roles = AKT_SETUP_GROUP_FIELDS.map((f) => f.role);
        expect(roles).toContain('groupAdmin');
        expect(roles).toContain('groupLehrer');
        expect(AKT_SETUP_GROUP_FIELDS.every((f) => f.pickBtnId && f.clearBtnId)).toBe(true);
        expect(AKT_SETUP_GROUP_FIELDS.find((f) => f.role === 'groupLehrer').dialogTitle).toMatch(
            /Einreicher/i
        );
    });
});
