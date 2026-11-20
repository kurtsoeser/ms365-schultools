import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import {
    canUseLfrRoleSwitcher,
    isLfrDemoRoleUiEnabled,
    sortLfrRoles,
    resolveActiveLfrRole,
    shellPreferredLfrRole,
    roleSourceLabel,
    listRolesFromEntraGroups,
    refineLfrDirektionFromEntra,
    lfrAccountIsDirektion,
    enforceLfrDirektionRole,
    applyLfrRoleFromEntra
} from '../src/tools/lehrer-freistellung-planer/lfr-entra-role.js';
import {
    accountIsLfrDirektionAccess,
    normalizeLfrPermissions
} from '../src/tools/lehrer-freistellung-planer/lfr-permissions.js';

function memoryStorage() {
    /** @type {Record<string, string>} */
    const map = {};
    return {
        getItem(k) {
            return Object.prototype.hasOwnProperty.call(map, k) ? map[k] : null;
        },
        setItem(k, v) {
            map[k] = String(v);
        },
        removeItem(k) {
            delete map[k];
        }
    };
}

describe('lfr-entra-role policy', () => {
    it('Demo nur mit Override', () => {
        expect(isLfrDemoRoleUiEnabled(true, false)).toBe(false);
        expect(isLfrDemoRoleUiEnabled(false, true)).toBe(true);
    });

    it('Switcher nur mit ?demoRole=1 (nie produktiv)', () => {
        expect(canUseLfrRoleSwitcher(['lehrer'], false)).toBe(false);
        expect(canUseLfrRoleSwitcher(['direktion'], false)).toBe(false);
        expect(canUseLfrRoleSwitcher(['direktion', 'lehrer'], false)).toBe(false);
        expect(canUseLfrRoleSwitcher(['lehrer'], true)).toBe(true);
        expect(canUseLfrRoleSwitcher(['direktion', 'lehrer'], true)).toBe(true);
    });

    it('sortiert und wählt aktive Rolle', () => {
        expect(sortLfrRoles(['lehrer', 'direktion'])).toEqual(['direktion', 'lehrer']);
        expect(resolveActiveLfrRole(['lehrer', 'direktion'], { preferredActiveRole: 'lehrer' })).toBe(
            'lehrer'
        );
    });

    it('roleSourceLabel für Stammdaten-Direktion', () => {
        expect(roleSourceLabel('stammdaten-direktion')).toBe('Schulleitung (Stammdaten)');
        expect(roleSourceLabel('setup-direktion')).toBe('Genehmiger (Setup)');
    });

    it('listRolesFromEntraGroups: nur LFR-Direktion-Gruppe (nicht Schüler-FR)', () => {
        const dirId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
        const member = new Set([dirId]);
        expect(
            listRolesFromEntraGroups(member, {
                groupDirektionId: dirId,
                groupLehrerId: 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
            })
        ).toEqual(['direktion']);
        // Schüler-FR-Gruppe wird nicht mehr als 3. Argument ausgewertet
        expect(
            listRolesFromEntraGroups(member, { groupDirektionId: '', groupLehrerId: '' })
        ).toEqual([]);
    });

    it('listRolesFromEntraGroups: gleiche GUID Direktion=Lehrkräfte → nur Lehrkraft', () => {
        const same = 'cccccccc-cccc-cccc-cccc-cccccccccccc';
        const member = new Set([same]);
        expect(
            listRolesFromEntraGroups(member, {
                groupDirektionId: same,
                groupLehrerId: same
            })
        ).toEqual(['lehrer']);
    });

    it('refineLfrDirektionFromEntra: transitive Entra-Direktion ohne Direktmitgliedschaft streichen', () => {
        const refined = refineLfrDirektionFromEntra({
            roles: ['direktion', 'lehrer'],
            sources: { direktion: 'entra', lehrer: 'entra' },
            hasDirectDirektion: false,
            hasStrongDirektionSignal: false
        });
        expect(refined.roles).toEqual(['lehrer']);
        expect(refined.sources.direktion).toBeUndefined();
    });

    it('refineLfrDirektionFromEntra: direkte Mitgliedschaft oder starkes Signal behalten', () => {
        expect(
            refineLfrDirektionFromEntra({
                roles: ['direktion', 'lehrer'],
                sources: { direktion: 'entra', lehrer: 'entra' },
                hasDirectDirektion: true,
                hasStrongDirektionSignal: false
            }).roles
        ).toEqual(['direktion', 'lehrer']);
        expect(
            refineLfrDirektionFromEntra({
                roles: ['direktion', 'lehrer'],
                sources: { direktion: 'entra', lehrer: 'stammdaten' },
                hasDirectDirektion: false,
                hasStrongDirektionSignal: true
            }).roles
        ).toEqual(['direktion', 'lehrer']);
    });

    it('Direktor bleibt Direktion, auch wenn die Portal-Rolle Lehrkraft wäre', () => {
        expect(lfrAccountIsDirektion('Direktor@ms365.schule', {})).toBe(true);
        expect(lfrAccountIsDirektion('lehrer01.demo@ms365.schule', {})).toBe(false);
        expect(
            enforceLfrDirektionRole(
                'lehrer',
                ['direktion', 'lehrer'],
                'direktor@ms365.schule',
                {}
            )
        ).toBe('direktion');
        expect(
            enforceLfrDirektionRole('lehrer', ['lehrer'], 'lehrer01.demo@ms365.schule', {})
        ).toBe('lehrer');
    });

    it('ohne gelesenes Konto wird keine Lehrkraft-Rolle gespeichert', async () => {
        const prev = globalThis.localStorage;
        globalThis.localStorage = memoryStorage();
        const state = { accountEmail: '', role: 'direktion', stammdaten: { teachers: [] } };
        await applyLfrRoleFromEntra(state, { preferShellDefault: true });
        expect(state.planerRoleSources.lehrer).toBe('konto-offen');
        expect(state.roleHint).toContain('Konto wird gelesen');
        expect(localStorage.getItem('ms365-lfr-planer-role-v1')).toBe(null);
        if (prev === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prev;
    });

    it('Direktor-Konto wird Direktion, auch ohne Gruppe und trotz gespeicherter Lehrkraft', async () => {
        const prev = globalThis.localStorage;
        globalThis.localStorage = memoryStorage();
        localStorage.setItem('ms365-lfr-planer-role-v1', 'lehrer');
        const state = {
            accountEmail: 'direktor@ms365.schule',
            role: 'lehrer',
            stammdaten: { teachers: [] }
        };
        await applyLfrRoleFromEntra(state, {
            preferShellDefault: true,
            preferredActiveRole: null
        });
        expect(state.role).toBe('direktion');
        expect(state.planerRoleSources.direktion).toBe('demo-direktor');
        expect(localStorage.getItem('ms365-lfr-planer-role-v1')).toBe('direktion');
        if (prev === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prev;
    });

    it('accountIsLfrDirektionAccess ignoriert Schüler-FR-Genehmiger-Fallback', () => {
        const prev = globalThis.localStorage;
        globalThis.localStorage = memoryStorage();
        localStorage.setItem(
            'ms365-freistellung-setup-v1',
            JSON.stringify({ emailDirektion: 'lehrer01.demo@ms365.schule' })
        );
        const perms = normalizeLfrPermissions({
            emailDirektion: '',
            direktionUsers: []
        });
        expect(accountIsLfrDirektionAccess('lehrer01.demo@ms365.schule', perms)).toBe(false);
        if (prev === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prev;
    });
});

describe('shellPreferredLfrRole', () => {
    const key = 'ms365-dashboard-it-preview-v1';
    const prev = globalThis.localStorage;
    const prevWin = globalThis.window;

    beforeEach(() => {
        globalThis.localStorage = memoryStorage();
    });

    afterEach(() => {
        if (prev === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prev;
        if (prevWin === undefined) delete globalThis.window;
        else globalThis.window = prevWin;
    });

    it('Portal → Lehrer nur für Schul-IT (Vorschau)', () => {
        localStorage.setItem(key, 'portal');
        globalThis.window = {
            __ms365DashboardPersonasLast: { loggedIn: true, isIt: true }
        };
        expect(shellPreferredLfrRole(['direktion', 'lehrer'])).toBe('lehrer');
    });

    it('Portal → Direktion für Genehmiger (nicht Schul-IT)', () => {
        localStorage.setItem(key, 'portal');
        globalThis.window = {
            __ms365DashboardPersonasLast: { loggedIn: true, isIt: false, isLehrer: true }
        };
        expect(shellPreferredLfrRole(['direktion', 'lehrer'])).toBe('direktion');
    });

    it('IT-Backend → Direktion bei Mehrfachrollen', () => {
        localStorage.setItem(key, 'full');
        globalThis.window = {
            __ms365DashboardPersonasLast: { loggedIn: true, isIt: true }
        };
        expect(shellPreferredLfrRole(['direktion', 'lehrer'])).toBe('direktion');
    });
});

describe('resolveActiveLfrRole shell default', () => {
    const key = 'ms365-dashboard-it-preview-v1';
    const roleKey = 'ms365-lfr-planer-role-v1';
    const prev = globalThis.localStorage;
    const prevWin = globalThis.window;

    beforeEach(() => {
        globalThis.localStorage = memoryStorage();
        localStorage.setItem(roleKey, 'lehrer');
        localStorage.setItem(key, 'portal');
        globalThis.window = {
            __ms365DashboardPersonasLast: { loggedIn: true, isIt: false }
        };
    });

    afterEach(() => {
        if (prev === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prev;
        if (prevWin === undefined) delete globalThis.window;
        else globalThis.window = prevWin;
    });

    it('preferShellDefault: sticky Lehrer verdrängt Direktion nicht', () => {
        expect(
            resolveActiveLfrRole(['direktion', 'lehrer'], {
                preferredActiveRole: null,
                preferShellDefault: true
            })
        ).toBe('direktion');
    });
});
