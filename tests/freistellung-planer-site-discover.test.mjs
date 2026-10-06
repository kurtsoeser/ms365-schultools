import { describe, it, expect } from 'vitest';
import { freistellungSiteDiscoveryHints } from '../src/tools/freistellung-planer/freistellung-planer-site-discover.js';

describe('freistellung-planer-site-discover', () => {
    it('freistellungSiteDiscoveryHints enthält MS365-Schultools', () => {
        globalThis.window = {
            MS365_FREISTELLUNG_PLANER: { sitePaths: ['/sites/MS365-Schultools'] },
            location: { search: '' }
        };
        const hints = freistellungSiteDiscoveryHints();
        expect(hints.some((h) => /MS365-Schultools/i.test(h))).toBe(true);
    });
});
