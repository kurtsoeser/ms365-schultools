import { describe, expect, it } from 'vitest';
import {
    catalogPanelFromHash,
    CATALOG_PANEL_HASH,
    hashForCatalogPanel
} from '../src/shared/dashboard-catalog-hash.js';

describe('dashboard-catalog-hash', () => {
    it('mappt Spec-Hashes auf Katalog-Panels', () => {
        expect(catalogPanelFromHash('#mitgliedschaften')).toBe('gruppen');
        expect(catalogPanelFromHash('#aufraeumen')).toBe('kommunikation');
        expect(catalogPanelFromHash('#automationen')).toBe('automationen');
    });

    it('liefert Hash für Panel', () => {
        expect(hashForCatalogPanel('intranet')).toBe(CATALOG_PANEL_HASH.intranet);
        expect(hashForCatalogPanel('schuljahr')).toBe('schuljahresstart');
    });
});
