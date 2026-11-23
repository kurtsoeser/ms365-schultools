import { describe, expect, it } from 'vitest';
import {
    SALES_BUNDLES,
    SALES_ADDONS,
    entitlementsForSalesBundle,
    matchSalesBundle,
    salesBundleById
} from '../src/shared/sales-bundles-catalog.js';
import { sumPackageListPrices } from '../src/shared/license-feature-catalog.js';

describe('sales-bundles-catalog', () => {
    it('drei Verkaufsbündel mit festen Preisen', () => {
        expect(SALES_BUNDLES.length).toBe(3);
        expect(salesBundleById('starter-it')?.priceEuro).toBe(390);
        expect(salesBundleById('schulalltag')?.priceEuro).toBe(890);
        expect(salesBundleById('schule-plus')?.priceEuro).toBe(1690);
        expect(salesBundleById('schulalltag')?.featured).toBe(true);
    });

    it('Schulalltag enthält Starter-Pakete plus Portal-Apps', () => {
        const starter = new Set(salesBundleById('starter-it').packageIds);
        const alltag = salesBundleById('schulalltag').packageIds;
        for (const id of starter) expect(alltag).toContain(id);
        expect(alltag).toContain('intranet');
        expect(alltag).toContain('schulapps');
        expect(alltag).toContain('freistellung');
        expect(alltag).not.toContain('ablauf');
    });

    it('Plus enthält Ablauf und Automationen', () => {
        const plus = salesBundleById('schule-plus').packageIds;
        expect(plus).toContain('ablauf');
        expect(plus).toContain('ausgaben');
        expect(plus).toContain('betrieb');
        expect(plus).toContain('automationen');
    });

    it('Bundle günstiger als Summe der Listenpreise', () => {
        for (const b of SALES_BUNDLES) {
            const sum = sumPackageListPrices(b.packageIds);
            expect(sum).toBeGreaterThan(b.priceEuro);
        }
    });

    it('entitlementsForSalesBundle setzt Pakete und leert Overrides', () => {
        const e = entitlementsForSalesBundle('schulalltag');
        expect(e.packages).toContain('freistellung');
        expect(e.toolsInclude).toEqual([]);
        expect(e.toolsExclude).toEqual([]);
    });

    it('matchSalesBundle erkennt Schulalltag', () => {
        const pkgs = salesBundleById('schulalltag').packageIds;
        const m = matchSalesBundle(pkgs);
        expect(m.bundle?.id).toBe('schulalltag');
        expect(m.exact).toBe(true);
    });

    it('matchSalesBundle: Bundle mit Extra ist Treffer aber nicht exact', () => {
        const pkgs = [...salesBundleById('starter-it').packageIds, 'kommunikation'];
        const m = matchSalesBundle(pkgs);
        expect(m.bundle?.id).toBe('starter-it');
        expect(m.exact).toBe(false);
    });

    it('nur Setup als Extra (keine Doppelung zu Bausteinen)', () => {
        expect(SALES_ADDONS).toHaveLength(1);
        expect(SALES_ADDONS[0].id).toBe('addon-setup');
        expect(SALES_ADDONS[0].once).toBe(true);
    });
});
