import { describe, expect, it } from 'vitest';
import {
    DASHBOARD_CATALOG_TAB_IDS,
    normalizeDashboardCatalogTab
} from '../src/shared/dashboard-catalog-layout.js';
import { DASHBOARD_CLUSTER_ORDER, DASHBOARD_TOOL_CLUSTER } from '../src/shared/dashboard-audience-catalog.js';

describe('dashboard-catalog-layout', () => {
    it('hat acht Sidebar-Tabs ohne Übersicht', () => {
        expect(DASHBOARD_CATALOG_TAB_IDS).toHaveLength(8);
        expect(DASHBOARD_CATALOG_TAB_IDS).not.toContain('uebersicht');
    });

    it('migriert gespeicherte Legacy-Tabs', () => {
        expect(normalizeDashboardCatalogTab('uebersicht')).toBe('kommunikation');
        expect(normalizeDashboardCatalogTab('regeln')).toBe('kommunikation');
        expect(normalizeDashboardCatalogTab('website')).toBe('intranet');
    });

    it('ordnet Personen-Werkzeuge dem Cluster Personen & Gäste zu', () => {
        expect(DASHBOARD_TOOL_CLUSTER['personen-verwaltung']).toBe('personen');
        expect(DASHBOARD_TOOL_CLUSTER['schulstruktur-sync']).toBe('hygiene');
    });

    it('listet Cluster aufgabenorientiert', () => {
        expect(DASHBOARD_CLUSTER_ORDER).toEqual([
            'gruppen',
            'unterricht',
            'personen',
            'planung',
            'intranet',
            'schulapps',
            'hygiene',
            'kommunikation',
            'automationen'
        ]);
    });
});
