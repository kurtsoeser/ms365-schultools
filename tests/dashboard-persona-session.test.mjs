import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import {
    isDashboardItResolutionInconclusive,
    refreshDashboardPersonaSession,
    clearDashboardPersonaCache
} from '../src/shared/dashboard-persona-session.js';

describe('isDashboardItResolutionInconclusive', () => {
    const originalWindow = globalThis.window;

    beforeEach(() => {
        globalThis.window = {
            ms365LicenseGateCore: null,
            ms365OperatorAccess: null
        };
    });

    afterEach(() => {
        globalThis.window = originalWindow;
        vi.restoreAllMocks();
    });

    it('false wenn ausgeloggt oder bereits IT', () => {
        expect(isDashboardItResolutionInconclusive(null)).toBe(false);
        expect(isDashboardItResolutionInconclusive({ loggedIn: false })).toBe(false);
        expect(isDashboardItResolutionInconclusive({ loggedIn: true, isIt: true })).toBe(false);
    });

    it('true wenn License-API konfiguriert, aber noch kein Cache', () => {
        globalThis.window.ms365LicenseGateCore = {
            isLicenseApiConfigured: () => true,
            readCache: () => null
        };
        expect(
            isDashboardItResolutionInconclusive({ loggedIn: true, isIt: false, isLehrer: true })
        ).toBe(true);
    });

    it('true wenn Admin-Menü pending (Operator noch unklar)', () => {
        globalThis.window.ms365LicenseGateCore = {
            isLicenseApiConfigured: () => false
        };
        globalThis.window.ms365OperatorAccess = {
            shouldShowAdminMenuLink: () => true,
            isCurrentUserOperator: () => false
        };
        expect(
            isDashboardItResolutionInconclusive({ loggedIn: true, isIt: false, isLehrer: true })
        ).toBe(true);
    });

    it('false wenn License-Cache da und Operator klar nein', () => {
        globalThis.window.ms365LicenseGateCore = {
            isLicenseApiConfigured: () => true,
            readCache: () => ({ allowed: true, isDesignatedSchoolIt: false })
        };
        globalThis.window.ms365OperatorAccess = {
            shouldShowAdminMenuLink: () => false,
            isCurrentUserOperator: () => false
        };
        expect(
            isDashboardItResolutionInconclusive({ loggedIn: true, isIt: false, isLehrer: true })
        ).toBe(false);
    });
});

describe('refreshDashboardPersonaSession ohne Konto', () => {
    const originalWindow = globalThis.window;
    const originalSession = globalThis.sessionStorage;

    beforeEach(() => {
        const store = new Map();
        globalThis.sessionStorage = {
            getItem: (k) => (store.has(k) ? store.get(k) : null),
            setItem: (k, v) => store.set(k, String(v)),
            removeItem: (k) => store.delete(k)
        };
        globalThis.window = {
            dispatchEvent: () => true,
            addEventListener: () => {},
            ms365AuthGetAccountInfo: () => null
        };
        clearDashboardPersonaCache();
    });

    afterEach(() => {
        globalThis.window = originalWindow;
        globalThis.sessionStorage = originalSession;
        delete globalThis.location;
        delete globalThis.localStorage;
    });

    it('startet nach persona-ready keine zweite Auflösung', async () => {
        let events = 0;
        globalThis.window.dispatchEvent = (ev) => {
            if (ev && ev.type === 'ms365-dashboard-persona-ready') {
                events += 1;
                setTimeout(() => {
                    refreshDashboardPersonaSession({ force: false });
                }, 0);
            }
            return true;
        };

        const personas = await refreshDashboardPersonaSession({ force: true });
        await new Promise((r) => setTimeout(r, 30));

        expect(personas.loggedIn).toBe(false);
        expect(events).toBe(1);
    });

    it('meldet die Screenshot-Persona nur einmal, auch wenn der Listener neu auflöst', async () => {
        globalThis.location = { search: '?landingShot=1&persona=lehrer' };
        globalThis.localStorage = {
            setItem() {},
            getItem() {
                return null;
            }
        };
        let events = 0;
        globalThis.window.dispatchEvent = (ev) => {
            if (ev && ev.type === 'ms365-dashboard-persona-ready') {
                events += 1;
                refreshDashboardPersonaSession({ force: true });
            }
            return true;
        };

        const personas = await refreshDashboardPersonaSession({ force: true });

        expect(personas.isLehrer).toBe(true);
        expect(personas.personaSource).toBe('landing-shot');
        expect(events).toBe(1);
    });
});
