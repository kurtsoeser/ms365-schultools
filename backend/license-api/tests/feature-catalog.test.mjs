import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const {
    LICENSE_PACKAGES,
    DEFAULT_NEW_SCHOOL_PACKAGES,
    parseEntitlementsRaw,
    resolveFeatureEntitlements,
    defaultNewSchoolEntitlements,
    serializeEntitlements,
    allKnownToolIds
} = require('../src/lib/feature-catalog.js');
const { evaluateLicense } = require('../src/lib/evaluate-license.js');

describe('feature-catalog', () => {
    it('hat Core/Personen/Schuljahr als Default für neue Schulen', () => {
        expect(DEFAULT_NEW_SCHOOL_PACKAGES).toEqual(['core', 'personen', 'schuljahr']);
        const doc = defaultNewSchoolEntitlements();
        const resolved = resolveFeatureEntitlements(doc, { legacyFullAccess: false });
        expect(resolved.packages).toEqual(['core', 'personen', 'schuljahr']);
        expect(resolved.featureCount).toBeGreaterThan(10);
    });

    it('Legacy ohne JSON = alle Pakete', () => {
        const resolved = resolveFeatureEntitlements(null, { legacyFullAccess: true });
        expect(resolved.packages.length).toBe(LICENSE_PACKAGES.length);
        expect(resolved.featureCount).toBe(allKnownToolIds().size);
        expect(resolved.legacyFullAccess).toBe(true);
    });

    it('schließt Tools aus und nimmt Includes auf', () => {
        const resolved = resolveFeatureEntitlements(
            {
                packages: ['core'],
                toolsInclude: ['kursteams'],
                toolsExclude: ['datenhygiene']
            },
            { legacyFullAccess: false }
        );
        expect(resolved.features).toContain('kursteams');
        expect(resolved.features).not.toContain('datenhygiene');
        expect(resolved.packages).toEqual(['core']);
    });

    it('serialisiert und parst Entitlements', () => {
        const raw = serializeEntitlements({
            packages: ['core', 'personen'],
            toolsInclude: [],
            toolsExclude: ['teams-archiv']
        });
        const parsed = parseEntitlementsRaw(raw);
        expect(parsed.packages).toEqual(['core', 'personen']);
        expect(parsed.toolsExclude).toEqual(['teams-archiv']);
    });

    it('evaluateLicense liefert Entitlements mit', () => {
        const r = evaluateLicense({
            tenantId: 'aaa',
            fields: {
                Title: 'Test',
                Status: 'active',
                LicenseEntitlements: serializeEntitlements({
                    packages: ['core', 'personen'],
                    toolsInclude: [],
                    toolsExclude: []
                })
            },
            allowedStatuses: ['active', 'trial'],
            now: new Date('2026-06-15T12:00:00Z')
        });
        expect(r.allowed).toBe(true);
        expect(r.entitlements.packages).toEqual(['core', 'personen']);
        expect(r.entitlements.features).toContain('personen-verwaltung');
        expect(r.entitlements.features).not.toContain('kursteams');
    });
});
