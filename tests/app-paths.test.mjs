import { describe, expect, it } from 'vitest';
import { resolveToolsHref } from '../src/shared/app-paths.js';

describe('app-paths', () => {
    it('resolveToolsHref vom Root', () => {
        globalThis.window = { location: { pathname: '/index.html' } };
        expect(resolveToolsHref('datenhygiene.html')).toBe('/tools/datenhygiene.html');
        expect(resolveToolsHref('tools/datenhygiene.html')).toBe('/tools/datenhygiene.html');
        delete globalThis.window;
    });

    it('resolveToolsHref aus /tools/ heraus verdoppelt nicht', () => {
        globalThis.window = {
            location: { pathname: '/tools/datenhygiene.html' }
        };
        expect(resolveToolsHref('datenhygiene.html')).toBe('/tools/datenhygiene.html');
        expect(resolveToolsHref('tools/datenhygiene.html')).toBe('/tools/datenhygiene.html');
        delete globalThis.window;
    });

    it('resolveToolsHref mit Repo-Unterpfad', () => {
        globalThis.window = {
            location: { pathname: '/ms365-schultools/tools/jahrgangsgruppen.html' }
        };
        expect(resolveToolsHref('datenhygiene.html')).toBe('/ms365-schultools/tools/datenhygiene.html');
        delete globalThis.window;
    });
});
