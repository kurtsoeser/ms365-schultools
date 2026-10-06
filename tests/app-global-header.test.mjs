import { describe, expect, it } from 'vitest';
import { buildAppChromeHeaderHtml, resolveAppRootHref } from '../src/shared/app-header-chrome.js';
import { shouldMountAppGlobalHeader } from '../src/shared/app-global-header.js';

describe('app-global-header', () => {
    it('resolveAppRootHref liefert index aus tools/', () => {
        expect(resolveAppRootHref('index.html')).toBe('index.html');
    });

    it('shouldMountAppGlobalHeader ist ohne document false', () => {
        expect(shouldMountAppGlobalHeader()).toBe(false);
    });

    it('baut Tool-Header mit adminAppTopActions und 3 Zonen', () => {
        const html = buildAppChromeHeaderHtml('tool');
        expect(html).toContain('id="dashCompactHeader"');
        expect(html).toContain('id="adminAppTopActions"');
        expect(html).toContain('dash-compact-header__nav--spacer');
        expect(html).toContain('dash-compact-header__search-dash-link');
    });
});
