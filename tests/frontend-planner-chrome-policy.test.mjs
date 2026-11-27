import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import {
    isFrontendPlannerPage,
    showItAppChrome,
    shouldShowItAppChromeSync
} from '../src/shared/frontend-planner-chrome-policy.js';

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
        }
    };
}

describe('frontend-planner-chrome-policy', () => {
    it('isFrontendPlannerPage erkennt die vier Planer', () => {
        expect(isFrontendPlannerPage('/tools/freistellung-planer.html')).toBe(true);
        expect(isFrontendPlannerPage('/x/tools/schularbeiten-planer.html')).toBe(true);
        expect(isFrontendPlannerPage('/tools/projektwochen.html')).toBe(true);
        expect(isFrontendPlannerPage('/tools/lehrer-freistellung-planer.html')).toBe(true);
        expect(isFrontendPlannerPage('/tools/schulaktivitaeten-planer.html')).toBe(true);
        expect(isFrontendPlannerPage('/tools/freistellung-setup.html')).toBe(false);
    });

    describe('showItAppChrome', () => {
        const key = 'ms365-dashboard-it-preview-v1';
        const prev = globalThis.localStorage;

        beforeEach(() => {
            globalThis.localStorage = memoryStorage();
        });

        afterEach(() => {
            if (prev === undefined) delete globalThis.localStorage;
            else globalThis.localStorage = prev;
        });

        it('nur für IT/Global/Designated und Shell full', () => {
            localStorage.setItem(key, 'full');
            expect(showItAppChrome(null)).toBe(false);
            expect(showItAppChrome({ loggedIn: true, isIt: false, globalAdmin: false })).toBe(false);
            expect(showItAppChrome({ loggedIn: true, isIt: true })).toBe(true);
            expect(showItAppChrome({ loggedIn: true, globalAdmin: true })).toBe(true);
            expect(showItAppChrome({ loggedIn: true, designatedSchoolIt: true })).toBe(true);
            expect(showItAppChrome({ loggedIn: true, isSchueler: true, isIt: false })).toBe(false);
        });

        it('Portal-Shell blendet IT-Chrome aus', () => {
            localStorage.setItem(key, 'portal');
            expect(showItAppChrome({ loggedIn: true, isIt: true })).toBe(false);
        });
    });

    describe('shouldShowItAppChromeSync', () => {
        const key = 'ms365-dashboard-it-preview-v1';
        const prevLocal = globalThis.localStorage;
        const prevSession = globalThis.sessionStorage;

        beforeEach(() => {
            globalThis.localStorage = memoryStorage();
            globalThis.sessionStorage = memoryStorage();
            vi.stubGlobal('document', {
                body: { classList: { contains: () => false } },
                documentElement: {}
            });
            vi.stubGlobal('window', {
                location: { pathname: '/tools/verwaltung.html' },
                __ms365DashboardPersonasLast: null
            });
            vi.stubGlobal('location', { pathname: '/tools/verwaltung.html' });
        });

        afterEach(() => {
            if (prevLocal === undefined) delete globalThis.localStorage;
            else globalThis.localStorage = prevLocal;
            if (prevSession === undefined) delete globalThis.sessionStorage;
            else globalThis.sessionStorage = prevSession;
            vi.unstubAllGlobals();
        });

        it('hält IT-Chrome während Workspace-Seitenwechsel (ohne Persona noch)', () => {
            localStorage.setItem(key, 'full');
            sessionStorage.setItem('ms365-it-workspace-v1', '1');
            expect(shouldShowItAppChromeSync()).toBe(true);
        });

        it('bleibt portal-first ohne Workspace-Flag und ohne Persona', () => {
            localStorage.setItem(key, 'full');
            expect(shouldShowItAppChromeSync()).toBe(false);
        });
    });
});
