import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';
import { describe, expect, it } from 'vitest';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

function read(rel) {
    return readFileSync(join(root, rel), 'utf8');
}

function quotedList(source, name) {
    const match = source.match(new RegExp(name + ' = \\[([^\\]]+)\\]'));
    expect(match, name).toBeTruthy();
    return match[1].match(/'([^']+)'/g).map((entry) => entry.slice(1, -1));
}

function bootContext(stored) {
    const style = {
        colorScheme: '',
        props: {},
        setProperty(name, value) {
            this.props[name] = value;
        },
        removeProperty(name) {
            delete this.props[name];
        }
    };
    const rootEl = {
        attrs: {},
        style,
        setAttribute(name, value) {
            this.attrs[name] = value;
        },
        getAttribute(name) {
            return this.attrs[name];
        }
    };
    const store = Object.assign({}, stored);
    const sandbox = {
        document: { documentElement: rootEl },
        localStorage: {
            getItem(key) {
                return Object.prototype.hasOwnProperty.call(store, key) ? store[key] : null;
            },
            setItem(key, value) {
                store[key] = String(value);
            }
        }
    };
    sandbox.window = sandbox;
    vm.createContext(sandbox);
    vm.runInContext(read('src/shared/theme-boot.js'), sandbox);
    sandbox.store = store;
    return sandbox;
}

describe('brand theme assets', () => {
    const ids = ['teal', 'classic', 'wine', 'azure', 'forest', 'amber', 'custom'];

    it('dieselbe Markenliste in Boot, Toggle und Konto-Menü', () => {
        expect(quotedList(read('src/shared/theme-boot.js'), 'var BRAND_IDS')).toEqual(ids);
        expect(quotedList(read('src/shared/theme-toggle.js'), 'const VALID_BRAND_LIST')).toEqual(ids);
        expect(quotedList(read('src/shared/msal-auth-ui.js'), 'const MENU_BRANDS')).toEqual(ids);
    });

    it('jede Kennung hat einen CSS-Block, Klassisch und Wein nur einmal', () => {
        const css = read('app.css');
        for (const id of ids) {
            expect(css).toContain(`html[data-brand="${id}"]`);
        }
        expect(css.match(/html\[data-brand="classic"\]\[data-theme="light"\]/g)).toHaveLength(1);
        expect(css.match(/html\[data-brand="wine"\]\[data-theme="light"\]/g)).toHaveLength(1);
        expect(css).toMatch(/html\[data-brand="wine"\]\[data-theme="light"\][\s\S]*--bg1:\s*#7f1d1d/);
        expect(css).toMatch(/html\[data-brand="azure"\]\[data-theme="light"\][\s\S]*--brand1:\s*#1d4ed8/);
        expect(css).toMatch(/html\[data-brand="forest"\]\[data-theme="light"\][\s\S]*--brand1:\s*#15803d/);
        expect(css).toMatch(/html\[data-brand="amber"\]\[data-theme="light"\][\s\S]*--brand1:\s*#c2410c/);
    });

    it('Konto-Menü zeigt sechs Marken und eigene Farbe', () => {
        const js = read('src/shared/msal-auth-ui.js');
        expect(js).toContain('applyMenuBrandChoice');
        expect(js).toMatch(/setAttribute\('data-brand', next\)/);
        expect(js).toContain('brandGridHtml');
        expect(js).toContain("wine: 'Wein'");
        expect(js).toContain("azure: 'Blau'");
        expect(js).toContain("forest: 'Grün'");
        expect(js).toContain("amber: 'Bernstein'");
        expect(js).not.toContain('brand-row--pair');
        expect(js).toContain('Eigene Farbe');
        expect(js).toContain("data-ms365-auth-menu-v', '3'");
        expect(js).toContain("menuV !== '3'");
    });

    it('Dashboard und Admin laden theme-boot zuerst, ohne eigene Markenliste', () => {
        const indexHead = read('index.html').slice(0, 400);
        const adminHead = read('admin.html').slice(0, 400);
        expect(indexHead.indexOf('src/shared/theme-boot.js')).toBeGreaterThan(-1);
        expect(indexHead.indexOf('src/shared/theme-boot.js')).toBeLessThan(indexHead.indexOf('src/pages/index-boot.js'));
        expect(adminHead.indexOf('src/shared/theme-boot.js')).toBeLessThan(adminHead.indexOf('src/pages/admin-boot.js'));
        expect(read('src/pages/index-boot.js')).not.toContain('ms365-brand-v1');
        expect(read('src/pages/admin-boot.js')).not.toContain('ms365-brand-v1');
        expect(read('src/pages/index-boot.js')).toContain('data-dash-main-section');
    });

    it('helles Eigenfarben-Akzent wird für den Header abgedunkelt, ungültiger Speicher fällt auf teal', () => {
        const boot = bootContext().window.ms365BrandBoot;
        expect(boot.headerLumMax).toBe(0.55);
        expect(boot.darkenForHeader('#0d9f8a')).toBe('#0d9f8a');
        expect(boot.luminance(boot.darkenForHeader('#ffffff'))).toBeLessThanOrEqual(0.55);
        expect(boot.darkenForHeader('#ffffff')).not.toBe('#ffffff');
        expect(boot.parseHex('#abc')).toBe('#aabbcc');
        expect(boot.parseHex('rot')).toBe('');

        const invalid = bootContext({
            'ms365-brand-v1': 'custom',
            'ms365-brand-custom-v1': '{'
        });
        expect(invalid.document.documentElement.attrs['data-brand']).toBe('teal');
        expect(invalid.store['ms365-brand-v1']).toBe('teal');
        expect(invalid.document.documentElement.style.props['--brand1']).toBeUndefined();

        const custom = bootContext({
            'ms365-brand-v1': 'custom',
            'ms365-brand-custom-v1': JSON.stringify({ brand1: '#ffff00', brand2: '#1d4ed8' })
        });
        const props = custom.document.documentElement.style.props;
        expect(custom.document.documentElement.attrs['data-brand']).toBe('custom');
        expect(props['--brand1']).toBe('#ffff00');
        expect(props['--brand2']).toBe('#1d4ed8');
        expect(boot.luminance(props['--brand-header-1'])).toBeLessThanOrEqual(0.55);
        expect(props['--brand-header-2']).toBe('#1d4ed8');

        custom.window.ms365BrandBoot.syncCustom(custom.document.documentElement, 'azure');
        expect(custom.document.documentElement.style.props['--brand1']).toBeUndefined();
    });
});
