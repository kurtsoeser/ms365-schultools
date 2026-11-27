import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { readItShellChoice, resolveLayoutModeFromPersonas, writeItPreviewChoice } from '../src/shared/dashboard-layout-mode.js';

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
        key(i) {
            return Object.keys(data)[i] || null;
        },
        get length() {
            return Object.keys(data).length;
        }
    };
}

describe('dashboard-layout-mode IT-Shell Choice', () => {
    const prevLocal = globalThis.localStorage;

    beforeEach(() => {
        globalThis.localStorage = memoryStorage();
        vi.stubGlobal('window', {
            ms365AuthGetAccountInfo: () => null
        });
    });

    afterEach(() => {
        if (prevLocal === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prevLocal;
        vi.unstubAllGlobals();
    });

    it('liest full auch aus Tenant-Key, wenn Unscoped leer ist', () => {
        localStorage.setItem('ms365-dashboard-it-preview-v1:tenant-a', 'full');
        expect(readItShellChoice()).toBe('full');
    });

    it('schreibt full parallel in Unscoped und Tenant-Key', () => {
        window.ms365AuthGetAccountInfo = () => ({ tenantId: 'tenant-b' });
        writeItPreviewChoice('full');
        expect(localStorage.getItem('ms365-dashboard-it-preview-v1')).toBe('full');
        expect(localStorage.getItem('ms365-dashboard-it-preview-v1:tenant-b')).toBe('full');
        expect(readItShellChoice()).toBe('full');
    });

    it('respektiert explizites portal auf dem aktuellen Key', () => {
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'portal');
        localStorage.setItem('ms365-dashboard-it-preview-v1:other', 'full');
        expect(readItShellChoice()).toBe('portal');
    });

    it('bei aktiver IT-Arbeitsfläche gewinnt Tenant-full gegen Unscoped-portal', () => {
        globalThis.sessionStorage = memoryStorage();
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'portal');
        localStorage.setItem('ms365-dashboard-it-preview-v1:tenant-x', 'full');
        sessionStorage.setItem('ms365-it-workspace-v1', '1');
        expect(readItShellChoice()).toBe('full');
    });

    it('bleibt im IT-Bereich, wenn die Shell auf full steht und die Persona Lehrkraft ist', () => {
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'full');
        expect(
            resolveLayoutModeFromPersonas({
                loggedIn: true,
                isIt: false,
                isLehrer: true,
                unknownUser: false
            })
        ).toBe('full');
        expect(resolveLayoutModeFromPersonas(null)).toBe('full');
    });
});
