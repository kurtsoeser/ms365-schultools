import { describe, it, expect, beforeEach } from 'vitest';
import {
    applyPortalSectionOrder,
    emptyPortalDefaultsConfig,
    loadPortalDefaultsConfig,
    normalizePortalDefaultsConfig,
    normalizeToolIdList,
    resolvePortalHomeToolIds,
    savePortalDefaultsConfig,
    setPortalHomePinned,
    setPortalSectionOrder,
    DASHBOARD_PORTAL_DEFAULTS_STORAGE_KEY
} from '../src/shared/dashboard-portal-defaults-store.js';

function memoryStorage(seed = {}) {
    /** @type {Record<string, string>} */
    const map = { ...seed };
    return {
        getItem(k) {
            return Object.prototype.hasOwnProperty.call(map, k) ? map[k] : null;
        },
        setItem(k, v) {
            map[k] = String(v);
        },
        removeItem(k) {
            delete map[k];
        },
        raw: map
    };
}

describe('dashboard-portal-defaults-store', () => {
    beforeEach(() => {
        // notify/dispatch may touch window; vitest node has limited window
    });

    it('normalisiert nur portal-fähige Tool-IDs', () => {
        expect(normalizeToolIdList(['srdp-anmeldung', 'jahrgang', 'srdp-anmeldung', ''])).toEqual([
            'srdp-anmeldung'
        ]);
    });

    it('resolvePortalHomeToolIds: Schul-Pins vor Favoriten vor Fallback', () => {
        const ids = resolvePortalHomeToolIds({
            visibleIds: ['a', 'b', 'c', 'd', 'e', 'f', 'g'],
            schoolPinned: ['c', 'a'],
            personalFavorites: ['e', 'a', 'g'],
            fallbackPriority: ['b', 'd'],
            max: 5
        });
        expect(ids).toEqual(['c', 'a', 'e', 'g', 'b']);
    });

    it('applyPortalSectionOrder sortiert bekannte Sektionen zuerst', () => {
        const sections = [
            { id: 'eltern', label: 'E' },
            { id: 'unterricht', label: 'U' },
            { id: 'schulalltag', label: 'S' }
        ];
        expect(applyPortalSectionOrder(sections, ['schulalltag', 'unterricht']).map((s) => s.id)).toEqual([
            'schulalltag',
            'unterricht',
            'eltern'
        ]);
    });

    it('speichert und lädt portalDefaults pro Persona', () => {
        const storage = memoryStorage();
        const saved = setPortalHomePinned(
            'lehrer',
            ['termine-hub', 'srdp-anmeldung', 'jahrgang'],
            emptyPortalDefaultsConfig(),
            storage
        );
        expect(saved.lehrer.homePinned).toEqual(['termine-hub', 'srdp-anmeldung']);
        expect(storage.getItem(DASHBOARD_PORTAL_DEFAULTS_STORAGE_KEY)).toContain('termine-hub');

        setPortalSectionOrder('schueler', ['meine', 'nope'], saved, storage);
        const loaded = loadPortalDefaultsConfig(storage);
        expect(loaded.schueler.sectionOrder).toEqual(['meine']);
        expect(loaded.lehrer.homePinned).toEqual(['termine-hub', 'srdp-anmeldung']);
    });

    it('normalizePortalDefaultsConfig verwirft Müll', () => {
        const n = normalizePortalDefaultsConfig({
            version: 9,
            lehrer: { homePinned: ['kursteams', 12, null], sectionOrder: ['unterricht', 'x'] },
            schueler: null
        });
        expect(n.version).toBe(1);
        expect(n.lehrer.homePinned).toEqual(['kursteams']);
        expect(n.lehrer.sectionOrder).toEqual(['unterricht']);
        expect(n.schueler.homePinned).toEqual([]);
    });

    it('savePortalDefaultsConfig schreibt JSON', () => {
        const storage = memoryStorage();
        savePortalDefaultsConfig(
            {
                version: 1,
                lehrer: { homePinned: ['elternbrief'], sectionOrder: [] },
                schueler: { homePinned: [], sectionOrder: [] }
            },
            storage
        );
        const again = loadPortalDefaultsConfig(storage);
        expect(again.lehrer.homePinned).toEqual(['elternbrief']);
    });
});
