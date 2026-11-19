import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { isDashboardItResolutionInconclusive } from '../src/shared/dashboard-persona-session.js';

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
