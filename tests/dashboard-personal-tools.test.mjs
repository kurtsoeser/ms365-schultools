import { describe, expect, it, beforeEach } from 'vitest';
import {
    DASHBOARD_RECENT_TOOLS_KEY,
    expandLegacyCatalogToolIds,
    loadRecentDashboardToolIds,
    normalizeDashboardToolHref,
    recordDashboardToolVisit,
    toolIdFromToolPageUrl
} from '../src/shared/dashboard-personal-tools.js';

describe('dashboard-personal-tools', () => {
    /** @type {Storage} */
    let storage;

    beforeEach(() => {
        const bag = new Map();
        storage = {
            getItem: (k) => (bag.has(k) ? bag.get(k) : null),
            setItem: (k, v) => bag.set(k, String(v))
        };
    });

    it('normalisiert Werkzeug-Hrefs', () => {
        expect(normalizeDashboardToolHref('tools/freistellung-planer.html')).toBe(
            'tools/freistellung-planer.html'
        );
        expect(normalizeDashboardToolHref('./tools/sharepoint-liste-lehrer.html?q=1')).toBe(
            'tools/sharepoint-liste-lehrer.html'
        );
    });

    it('speichert zuletzt verwendet ohne Duplikate', () => {
        recordDashboardToolVisit('a', { storage, max: 5 });
        recordDashboardToolVisit('b', { storage, max: 5 });
        recordDashboardToolVisit('a', { storage, max: 5 });
        expect(loadRecentDashboardToolIds(storage)).toEqual(['a', 'b']);
    });

    it('begrenzt die Recent-Liste', () => {
        for (let i = 0; i < 10; i += 1) {
            recordDashboardToolVisit('t' + i, { storage, max: 3 });
        }
        expect(loadRecentDashboardToolIds(storage)).toHaveLength(3);
    });

    it('expandLegacyCatalogToolIds splittet slg', () => {
        expect(expandLegacyCatalogToolIds(['slg', 'kursteams'])).toEqual([
            'slg-schueler',
            'slg-lehrer',
            'kursteams'
        ]);
    });

    it('toolIdFromToolPageUrl nutzt href-Map', () => {
        const map = new Map([['tools/jahrgangsgruppen.html', 'jahrgang']]);
        expect(
            toolIdFromToolPageUrl('https://schule.example/tools/jahrgangsgruppen.html', map)
        ).toBe('jahrgang');
    });

    it('schreibt in den erwarteten Storage-Key', () => {
        recordDashboardToolVisit('freistellung-planer', { storage });
        expect(storage.getItem(DASHBOARD_RECENT_TOOLS_KEY)).toContain('freistellung-planer');
    });
});
