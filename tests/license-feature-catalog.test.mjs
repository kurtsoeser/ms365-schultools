import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';
import {
    LICENSE_PACKAGES,
    DEFAULT_NEW_SCHOOL_PACKAGES,
    resolveFeatureEntitlements,
    allPackageIds,
    applyToolEntitlementToggle,
    pruneToolOverrides,
    toolCoveredByPackages
} from '../src/shared/license-feature-catalog.js';

const require = createRequire(import.meta.url);
const backend = require('../backend/license-api/src/lib/feature-catalog.js');

describe('license-feature-catalog Spiegel', () => {
    it('Frontend und Backend haben dieselben Paket-IDs und Tool-Listen', () => {
        expect(DEFAULT_NEW_SCHOOL_PACKAGES).toEqual(backend.DEFAULT_NEW_SCHOOL_PACKAGES);
        expect(allPackageIds()).toEqual(backend.allPackageIds());
        const frontById = Object.fromEntries(LICENSE_PACKAGES.map((p) => [p.id, p]));
        for (const pkg of backend.LICENSE_PACKAGES) {
            expect(frontById[pkg.id]).toBeTruthy();
            expect(frontById[pkg.id].toolIds).toEqual(pkg.toolIds);
            expect(frontById[pkg.id].requires || []).toEqual(pkg.requires || []);
            expect(frontById[pkg.id].priceHint || '').toEqual(pkg.priceHint || '');
            expect(frontById[pkg.id].priceEuro || null).toEqual(pkg.priceEuro || null);
        }
    });

    it('feste Modul-Listenpreise und Bundle-Vorteil', () => {
        const byId = Object.fromEntries(LICENSE_PACKAGES.map((p) => [p.id, p]));
        expect(byId.core.priceEuro).toBe(250);
        expect(byId.schulapps.priceEuro).toBe(300);
        expect(byId.ablauf.priceEuro).toBe(290);
        const starter =
            byId.core.priceEuro +
            byId.personen.priceEuro +
            byId.schuljahr.priceEuro +
            byId.unterricht.priceEuro;
        expect(starter).toBeGreaterThan(390);
        const alltag =
            starter + byId.intranet.priceEuro + byId.schulapps.priceEuro + byId.freistellung.priceEuro;
        expect(alltag).toBeGreaterThan(890);
    });

    it('Starter-Set löst sinnvolle Features auf', () => {
        const r = resolveFeatureEntitlements(
            { packages: ['core', 'personen', 'schuljahr'], toolsInclude: [], toolsExclude: [] },
            { legacyFullAccess: false }
        );
        expect(r.features).toContain('personen-verwaltung');
        expect(r.features).toContain('webuntis-stammdaten-import');
        expect(r.features).not.toContain('kursteams');
        expect(r.features).not.toContain('schueler-stammdaten-portal');
    });

    it('toolsInclude/toolsExclude steuern einzelne Werkzeuge', () => {
        const r = resolveFeatureEntitlements(
            {
                packages: ['core'],
                toolsInclude: ['kursteams'],
                toolsExclude: ['datenhygiene']
            },
            { legacyFullAccess: false }
        );
        expect(r.features).toContain('kursteams');
        expect(r.features).not.toContain('datenhygiene');
    });

    it('applyToolEntitlementToggle schaltet Include/Exclude und Abhängigkeiten', () => {
        const doc = {
            packages: ['core'],
            toolsInclude: [],
            toolsExclude: []
        };
        applyToolEntitlementToggle(doc, 'datenhygiene', false);
        expect(doc.toolsExclude).toContain('datenhygiene');
        expect(toolCoveredByPackages(doc, 'datenhygiene')).toBe(true);

        applyToolEntitlementToggle(doc, 'datenhygiene', true);
        expect(doc.toolsExclude).not.toContain('datenhygiene');

        const result = applyToolEntitlementToggle(doc, 'schueler-stammdaten-portal', true);
        expect(doc.toolsInclude).toContain('schueler-stammdaten-portal');
        expect(doc.packages).toContain('intranet');
        expect(result.packagesEnabled).toContain('intranet');

        doc.packages.push('stammdaten-portal');
        pruneToolOverrides(doc);
        expect(doc.toolsInclude).not.toContain('schueler-stammdaten-portal');
    });
});
