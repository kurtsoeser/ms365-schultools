import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const projectRoot = join(dirname(fileURLToPath(import.meta.url)), '..');

function read(rel) {
    return readFileSync(join(projectRoot, rel), 'utf8');
}

describe('Eine globale Menüleiste', () => {
    it('legt Schuljahr/Domain nicht mehr in eine Header-Toolbar', () => {
        const src = read('src/shared/context-bar.js');
        expect(src).not.toContain('data-ms365-auto-toolbar');
        expect(src).not.toContain('ms365ContextBar');
        expect(src).not.toContain('ms365-context-bar');
        expect(src).toContain('placeDashboardNav');
        expect(src).toContain('app-shell-chrome');
        expect(src).toContain('ms365NormalizeToolPageChrome');
        expect(src).toContain('hideEmptyToolbar');
        expect(src).toContain('ms365HeaderNav');
        expect(src).toContain('schoolYearSelect');
        expect(src).toContain('ms365AuthCtxDomain');
        expect(src).toContain('fillAccountContext');
    });

    it('setzt Dashboard unter der App-Leiste (app-shell), Konto rechts oben', () => {
        const auth = read('src/shared/msal-auth-ui.js');
        expect(auth).toContain('placeAuthWidgetInMenuHeader');
        expect(auth).toContain('ensureSchulregisterHeaderLink');
        expect(auth).toContain('ms365HeaderRightCluster');
        expect(auth).toContain('ms365-auth-menu');
        expect(auth).toContain('ms365AuthDropdown');
        expect(auth).toContain('schoolYearSelect');
        expect(auth).toContain('schoolYearAddBtn');
        expect(auth).toContain('ms365AuthCtxDomain');
        expect(auth).toContain('Schuljahr');
        expect(auth).toContain('Domain');
        expect(auth).toContain("badgeText.textContent = a ? name : 'Konto'");
        expect(auth).toContain('if (menu) menu.hidden = !a');
        expect(auth).toContain('Konto wechseln');
        expect(auth).toContain('Abmelden');
        expect(auth).toContain('aria-haspopup');
        expect(auth).not.toContain('ms365-auth-btn--icon');

        const css = read('app.css');
        expect(css).toContain('.ms365-header-nav');
        expect(css).toContain('.header > #ms365AuthWidget');
        expect(css).toContain('.ms365-auth-menu__panel');
        expect(css).toContain('.ms365-auth-menu__ctx');
        expect(css).toContain('.header .toolbar:not(:has(a[href], button, input, select, textarea))');
        expect(css).not.toContain('.ms365-context-bar');
    });
});
