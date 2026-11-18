import { describe, it, expect } from 'vitest';
import {
    isDemoRoleUiEnabled,
    canUseAppRoleSwitcher,
    canShowPrivilegedAppToolbar,
    primaryPersonaLabel,
    primaryPersonaHint,
    unknownPersonaHint,
    normalizeAppRoles
} from '../src/shared/app-role-switcher-policy.js';

describe('app-role-switcher-policy', () => {
    it('Demo-UI nur bei Override', () => {
        expect(isDemoRoleUiEnabled(false)).toBe(false);
        expect(isDemoRoleUiEnabled(true)).toBe(true);
    });

    it('normalisiert Rollen', () => {
        expect(normalizeAppRoles(['Admin', 'lehrer', 'ADMIN', ''])).toEqual(['admin', 'lehrer']);
    });

    it('SA-Muster: Switcher nur Admin+Mehrfach oder Demo', () => {
        expect(canUseAppRoleSwitcher(['lehrer'], { privilegedRoles: ['admin'] })).toBe(false);
        expect(canUseAppRoleSwitcher(['admin'], { privilegedRoles: ['admin'] })).toBe(false);
        expect(
            canUseAppRoleSwitcher(['admin', 'lehrer'], { privilegedRoles: ['admin'] })
        ).toBe(true);
        expect(
            canUseAppRoleSwitcher(['lehrer'], { demoRoleUi: true, privilegedRoles: ['admin'] })
        ).toBe(true);
    });

    it('FR-Muster: Direktion+Mehrfach oder Demo', () => {
        expect(canUseAppRoleSwitcher(['kv'], { privilegedRoles: ['direktion'] })).toBe(false);
        expect(
            canUseAppRoleSwitcher(['kv', 'schueler'], { privilegedRoles: ['direktion'] })
        ).toBe(false);
        expect(
            canUseAppRoleSwitcher(['direktion', 'schueler'], { privilegedRoles: ['direktion'] })
        ).toBe(true);
        expect(
            canUseAppRoleSwitcher(['schueler'], {
                demoRoleUi: true,
                privilegedRoles: ['direktion']
            })
        ).toBe(true);
    });

    it('privilegierte Toolbar nur in aktiver Admin-Rolle', () => {
        expect(canShowPrivilegedAppToolbar(['admin', 'lehrer'], 'admin', 'admin')).toBe(true);
        expect(canShowPrivilegedAppToolbar(['admin', 'lehrer'], 'lehrer', 'admin')).toBe(false);
        expect(canShowPrivilegedAppToolbar(['lehrer'], 'lehrer', 'admin')).toBe(false);
    });

    it('Persona-Label und Hinweis', () => {
        expect(primaryPersonaLabel({ loggedIn: true, isIt: true })).toBe('Schul-IT');
        expect(primaryPersonaLabel({ loggedIn: true, isLehrer: true })).toBe('Lehrkraft');
        expect(primaryPersonaLabel({ loggedIn: true, isSchueler: true })).toBe('Schüler/in');
        expect(
            primaryPersonaLabel({ loggedIn: true, isSchueler: true, unknownUser: true })
        ).toBe('Noch nicht zugeordnet');
        expect(primaryPersonaLabel({ loggedIn: false })).toBe('');
        expect(primaryPersonaHint({ loggedIn: true, isLehrer: true })).toMatch(/kein Umschalter/i);
        expect(primaryPersonaHint({ loggedIn: true, isIt: true })).toMatch(/IT-Backend/i);
    });

    it('unknownPersonaHint: Domain ≠ Lehrkraft, konkrete Checklist', () => {
        expect(
            unknownPersonaHint({
                unknownUser: true,
                personaDiagnostics: { hasDomain: false, teacherCount: 0 }
            })
        ).toMatch(/Schul-Domain fehlt/i);

        expect(
            unknownPersonaHint({
                unknownUser: true,
                personaDiagnostics: { hasDomain: true, teacherCount: 0 }
            })
        ).toMatch(/Lehrerliste ist leer/i);

        expect(
            unknownPersonaHint({
                unknownUser: true,
                personaSource: 'dashboard-entra-none',
                personaDiagnostics: {
                    hasDomain: true,
                    teacherCount: 12,
                    groupsConfigured: true,
                    inTeacherList: false
                }
            })
        ).toMatch(/Lehrkräfte-Gruppe/i);

        expect(
            unknownPersonaHint({
                unknownUser: true,
                personaDiagnostics: {
                    hasDomain: true,
                    teacherCount: 5,
                    groupsConfigured: false,
                    inTeacherList: false
                }
            })
        ).toMatch(/nicht in der Lehrerliste/i);

        expect(
            primaryPersonaHint({
                loggedIn: true,
                unknownUser: true,
                personaDiagnostics: { hasDomain: true, teacherCount: 0 }
            })
        ).toMatch(/Lehrerliste ist leer/i);
    });
});
