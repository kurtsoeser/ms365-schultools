import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

describe('brand theme assets', () => {
    it('theme-toggle.js unterstützt wine', () => {
        const js = readFileSync(join(root, 'src/shared/theme-toggle.js'), 'utf8');
        expect(js).toContain("wine: 'wine'");
        expect(js).toContain('syncBrandUi');
    });

    it('msal-auth-ui setzt data-brand vor setBrand', () => {
        const js = readFileSync(join(root, 'src/shared/msal-auth-ui.js'), 'utf8');
        expect(js).toContain('applyMenuBrandChoice');
        expect(js).toMatch(/setAttribute\('data-brand', next\)/);
    });

    it('app.css enthält wine-Variablen am Dateiende', () => {
        const css = readFileSync(join(root, 'app.css'), 'utf8');
        expect(css).toMatch(/html\[data-brand="wine"\]\[data-theme="light"\][\s\S]*--bg1:\s*#7f1d1d/);
    });
});
