import { describe, it, expect } from 'vitest';
import {
    entraGroupsConfigured,
    pickRoleFromEntraGroups,
    listRolesFromEntraGroups,
    pickRoleFromStammdaten,
    resolveActivePlanerRole,
    finalizePlanerRoles,
    planerEntraGroupIds,
    hasGlobalAdministratorDirectoryRole,
    GLOBAL_ADMINISTRATOR_ROLE_TEMPLATE_ID,
    isPlanerDemoRoleUiEnabled,
    canUsePlanerRoleSwitcher,
    canShowPlanerItToolbar
} from '../src/tools/schularbeiten-planer/schularbeiten-planer-entra-role.js';

const ADMIN = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
const LEHRER = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';
const SCHUELER = 'cccccccc-cccc-cccc-cccc-cccccccccccc';

const cfg = {
    groupAdmin: 'Admin',
    groupAdminId: ADMIN,
    groupLehrer: 'Lehrer',
    groupLehrerId: LEHRER,
    groupSchueler: 'Schüler',
    groupSchuelerId: SCHUELER,
    skipPerms: false
};

describe('schularbeiten-planer-entra-role', () => {
    it('entraGroupsConfigured erkennt mindestens eine Gruppen-ID', () => {
        expect(entraGroupsConfigured({})).toBe(false);
        expect(entraGroupsConfigured({ groupLehrerId: LEHRER })).toBe(true);
        expect(
            entraGroupsConfigured({
                adminUsers: [{ mail: 'sek@schule.at', displayName: 'Sek', id: 'x' }]
            })
        ).toBe(true);
    });

    it('planerEntraGroupIds dedupliziert gültige GUIDs', () => {
        const ids = planerEntraGroupIds({
            ...cfg,
            groupAdminId: LEHRER
        });
        expect(ids).toHaveLength(2);
        expect(ids).toContain(LEHRER);
        expect(ids).toContain(SCHUELER);
    });

    it('pickRoleFromEntraGroups priorisiert admin > lehrer > schueler', () => {
        const all = new Set([ADMIN, LEHRER, SCHUELER].map((x) => x.toLowerCase()));
        expect(pickRoleFromEntraGroups(all, cfg)).toBe('admin');
        expect(pickRoleFromEntraGroups(new Set([LEHRER.toLowerCase()]), cfg)).toBe('lehrer');
        expect(pickRoleFromEntraGroups(new Set([SCHUELER.toLowerCase()]), cfg)).toBe('schueler');
        expect(pickRoleFromEntraGroups(new Set(), cfg)).toBe('');
    });

    it('listRolesFromEntraGroups liefert alle Gruppen-Rollen', () => {
        const all = new Set([ADMIN, LEHRER].map((x) => x.toLowerCase()));
        expect(listRolesFromEntraGroups(all, cfg)).toEqual(['admin', 'lehrer']);
    });

    it('resolveActivePlanerRole behält gewählte Rolle', () => {
        expect(resolveActivePlanerRole(['admin', 'lehrer'], { preferredActiveRole: 'lehrer' })).toBe('lehrer');
        expect(resolveActivePlanerRole(['admin'], { preferredActiveRole: 'lehrer' })).toBe('admin');
    });

    it('pickRoleFromStammdaten liefert erste Rolle in fester Reihenfolge', () => {
        expect(pickRoleFromStammdaten({ teacherMatch: {}, studentMatch: {} })).toBe('lehrer');
        expect(pickRoleFromStammdaten({ teacherMatch: {} })).toBe('lehrer');
        expect(pickRoleFromStammdaten({ studentMatch: {} })).toBe('schueler');
        expect(pickRoleFromStammdaten({})).toBe('');
    });

    it('finalizePlanerRoles ergänzt Schüler für Admins', () => {
        const { roles, sources } = finalizePlanerRoles(['admin', 'lehrer'], { admin: 'entra', lehrer: 'entra' });
        expect(roles).toEqual(['admin', 'lehrer', 'schueler']);
        expect(sources.schueler).toBe('admin-schueler');
        const unchanged = finalizePlanerRoles(['lehrer'], { lehrer: 'entra' });
        expect(unchanged.roles).toEqual(['lehrer']);
    });

    it('isPlanerDemoRoleUiEnabled nur mit demoRole-Override', () => {
        expect(isPlanerDemoRoleUiEnabled(false, false)).toBe(false);
        expect(isPlanerDemoRoleUiEnabled(false, true)).toBe(true);
        expect(isPlanerDemoRoleUiEnabled(true, false)).toBe(false);
    });

    it('canUsePlanerRoleSwitcher nur für Admins mit mehreren Rollen oder Demo', () => {
        expect(canUsePlanerRoleSwitcher(['lehrer'], false)).toBe(false);
        expect(canUsePlanerRoleSwitcher(['lehrer', 'schueler'], false)).toBe(false);
        expect(canUsePlanerRoleSwitcher(['admin', 'lehrer', 'schueler'], false)).toBe(true);
        expect(canUsePlanerRoleSwitcher(['lehrer'], true)).toBe(true);
    });

    it('canShowPlanerItToolbar nur in Admin-Rolle mit Verwaltungsrecht', () => {
        expect(canShowPlanerItToolbar(['lehrer'], 'lehrer')).toBe(false);
        expect(canShowPlanerItToolbar(['admin', 'lehrer'], 'lehrer')).toBe(false);
        expect(canShowPlanerItToolbar(['admin', 'lehrer'], 'admin')).toBe(true);
    });

    it('hasGlobalAdministratorDirectoryRole erkennt Global Admin', () => {
        expect(hasGlobalAdministratorDirectoryRole([])).toBe(false);
        expect(
            hasGlobalAdministratorDirectoryRole([
                { roleTemplateId: GLOBAL_ADMINISTRATOR_ROLE_TEMPLATE_ID, displayName: 'Global Administrator' }
            ])
        ).toBe(true);
    });
});
