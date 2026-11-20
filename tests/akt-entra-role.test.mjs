import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import {
    canUseAktRoleSwitcher,
    isAktDemoRoleUiEnabled,
    resolveActiveAktRole,
    shellPreferredAktRole,
    shouldAutoApplyAktDemoTeacherPersona,
    sortAktRoles,
    roleSourceLabel
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-entra-role.js';

function memoryStorage() {
    /** @type {Record<string, string>} */
    const data = {};
    return {
        getItem(k) {
            return Object.prototype.hasOwnProperty.call(data, k) ? data[k] : null;
        },
        setItem(k, v) {
            data[k] = String(v);
        },
        removeItem(k) {
            delete data[k];
        },
        clear() {
            Object.keys(data).forEach((k) => delete data[k]);
        }
    };
}

describe('schulaktivitaeten-entra-role policy', () => {
    it('Demo nur mit Override', () => {
        expect(isAktDemoRoleUiEnabled(true, false)).toBe(false);
        expect(isAktDemoRoleUiEnabled(false, true)).toBe(true);
    });

    it('Kein Admin/KV/Lehrer-Umschalter – nur explizite Demo (?demoRole=1)', () => {
        expect(canUseAktRoleSwitcher(['lehrer', 'kv'], false)).toBe(false);
        expect(canUseAktRoleSwitcher(['admin'], false)).toBe(false);
        expect(canUseAktRoleSwitcher(['admin', 'lehrer'], false)).toBe(false);
        expect(canUseAktRoleSwitcher(['admin', 'kv', 'lehrer'], false)).toBe(false);
        expect(canUseAktRoleSwitcher(['kv'], true)).toBe(true);
        expect(canUseAktRoleSwitcher(['admin', 'lehrer'], true)).toBe(true);
    });

    it('sortiert Rollen', () => {
        expect(sortAktRoles(['lehrer', 'admin', 'kv'])).toEqual(['admin', 'kv', 'lehrer']);
    });

    it('roleSourceLabel kennt designated-it und Schulleitung', () => {
        expect(roleSourceLabel('designated-it')).toBe('Schul-IT');
        expect(roleSourceLabel('stammdaten-direktion')).toBe('Schulleitung (Stammdaten)');
    });
});

describe('shouldAutoApplyAktDemoTeacherPersona', () => {
    const prevWindow = globalThis.window;

    afterEach(() => {
        globalThis.window = prevWindow;
    });

    it('nur Schul-IT (designated-it / global-admin)', () => {
        expect(
            shouldAutoApplyAktDemoTeacherPersona({
                planerRoles: ['admin', 'lehrer'],
                planerRoleSources: { admin: 'designated-it' }
            })
        ).toBe(true);
        expect(
            shouldAutoApplyAktDemoTeacherPersona({
                planerRoles: ['admin'],
                planerRoleSources: { admin: 'global-admin' }
            })
        ).toBe(true);
    });

    it('Direktion / Genehmiger nie in Dave-Persona zwingen', () => {
        globalThis.window = {
            __ms365DashboardPersonasLast: { loggedIn: true, isIt: false }
        };
        for (const src of ['stammdaten-direktion', 'admin-user', 'direktion-user', 'entra']) {
            expect(
                shouldAutoApplyAktDemoTeacherPersona({
                    planerRoles: ['admin', 'kv', 'lehrer'],
                    planerRoleSources: { admin: src }
                })
            ).toBe(false);
        }
    });

    it('ohne Admin-Rolle false', () => {
        expect(
            shouldAutoApplyAktDemoTeacherPersona({
                planerRoles: ['lehrer', 'kv'],
                planerRoleSources: {}
            })
        ).toBe(false);
    });
});

describe('shellPreferredAktRole', () => {
    const key = 'ms365-dashboard-it-preview-v1';
    const prev = globalThis.localStorage;

    beforeEach(() => {
        globalThis.localStorage = memoryStorage();
    });

    afterEach(() => {
        if (prev === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prev;
    });

    it('Portal → Lehrer nur für Schul-IT (Vorschau)', () => {
        localStorage.setItem(key, 'portal');
        globalThis.window = {
            __ms365DashboardPersonasLast: { loggedIn: true, isIt: true }
        };
        expect(shellPreferredAktRole(['admin', 'kv', 'lehrer'])).toBe('lehrer');
    });

    it('Portal → Admin für Direktion (nicht Schul-IT)', () => {
        localStorage.setItem(key, 'portal');
        globalThis.window = {
            __ms365DashboardPersonasLast: { loggedIn: true, isIt: false, isLehrer: false }
        };
        expect(shellPreferredAktRole(['admin', 'kv', 'lehrer'])).toBe('admin');
    });

    it('IT-Backend → Admin bei Mehrfachrollen', () => {
        localStorage.setItem(key, 'full');
        globalThis.window = {
            __ms365DashboardPersonasLast: { loggedIn: true, isIt: true }
        };
        expect(shellPreferredAktRole(['admin', 'kv', 'lehrer'])).toBe('admin');
    });

    it('ohne passende Rolle null-sicher', () => {
        expect(shellPreferredAktRole([])).toBe(null);
    });
});

describe('resolveActiveAktRole', () => {
    const key = 'ms365-dashboard-it-preview-v1';
    const roleKey = 'ms365-akt-planer-role-v1';
    const prev = globalThis.localStorage;

    beforeEach(() => {
        globalThis.localStorage = memoryStorage();
    });

    afterEach(() => {
        if (prev === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prev;
    });

    it('preferredActiveRole gewinnt', () => {
        localStorage.setItem(key, 'full');
        expect(
            resolveActiveAktRole(['admin', 'lehrer'], { preferredActiveRole: 'lehrer' })
        ).toBe('lehrer');
    });

    it('preferShellDefault: Portal → Lehrer nur als Schul-IT', () => {
        localStorage.setItem(key, 'portal');
        localStorage.setItem(roleKey, 'admin');
        globalThis.window = {
            __ms365DashboardPersonasLast: { loggedIn: true, isIt: true }
        };
        expect(
            resolveActiveAktRole(['admin', 'lehrer'], { preferShellDefault: true })
        ).toBe('lehrer');
    });

    it('preferShellDefault: Direktion bleibt Admin trotz Portal-Shell', () => {
        localStorage.setItem(key, 'portal');
        localStorage.setItem(roleKey, 'lehrer');
        globalThis.window = {
            __ms365DashboardPersonasLast: { loggedIn: true, isIt: false }
        };
        expect(
            resolveActiveAktRole(['admin', 'lehrer'], { preferShellDefault: true })
        ).toBe('admin');
    });

    it('preferShellDefault: IT → Admin', () => {
        localStorage.setItem(key, 'full');
        localStorage.setItem(roleKey, 'lehrer');
        globalThis.window = {
            __ms365DashboardPersonasLast: { loggedIn: true, isIt: true }
        };
        expect(
            resolveActiveAktRole(['admin', 'lehrer'], { preferShellDefault: true })
        ).toBe('admin');
    });
});
