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
    const ids = ['teal'];

    it('dieselbe Markenliste in Boot, Toggle und Konto-Menü (nur Teal)', () => {
        expect(quotedList(read('src/shared/theme-boot.js'), 'var BRAND_IDS')).toEqual(ids);
        expect(quotedList(read('src/shared/theme-toggle.js'), 'const VALID_BRAND_LIST')).toEqual(ids);
        expect(quotedList(read('src/shared/msal-auth-ui.js'), 'const MENU_BRANDS')).toEqual(ids);
    });

    it('Design-Tokens sind die Hex-Quelle; app.css bleibt unter 30 Hex', () => {
        const tokens = read('src/shared/design-tokens.css');
        const app = read('app.css');
        expect(tokens).toMatch(/--brand1:\s*#0d9f8a/);
        expect(tokens).toContain('--status-ok');
        expect(tokens).toContain('--status-warn');
        expect(tokens).toContain('--status-danger');
        expect(app).toContain("url('src/shared/design-tokens.css')");
        expect(app).not.toMatch(/html\[data-brand="classic"\]/);
        expect(app).not.toMatch(/html\[data-brand="wine"\]/);
        expect(app).not.toMatch(/#5e72e4/i);
        const hexCount = (app.match(/#[0-9a-fA-F]{3,8}\b/g) || []).length;
        expect(hexCount).toBeLessThan(30);
    });

    it('Konto-Menü zeigt Teal als feste Marke, ohne Variantenwahl', () => {
        const js = read('src/shared/msal-auth-ui.js');
        expect(js).toContain('applyMenuBrandChoice');
        expect(js).toContain('brandGridHtml');
        expect(js).not.toContain('Markenfarbe der App');
        expect(js).not.toContain("wine: 'Wein'");
        expect(js).not.toContain("classic: 'Klassisch'");
        expect(js).not.toContain('Eigene Farbe');
        expect(js).toContain("data-ms365-auth-menu-v', '12'");
        expect(js).toContain("menuV !== '12'");
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

    it('abgekündigte Marken und Custom fallen auf teal', () => {
        const boot = bootContext().window.ms365BrandBoot;
        expect(boot.headerLumMax).toBe(0.55);
        expect(boot.darkenForHeader('#0d9f8a')).toBe('#0d9f8a');
        expect(boot.luminance(boot.darkenForHeader('#ffffff'))).toBeLessThanOrEqual(0.55);
        expect(boot.parseHex('#abc')).toBe('#aabbcc');
        expect(boot.parseHex('rot')).toBe('');

        const classic = bootContext({ 'ms365-brand-v1': 'classic' });
        expect(classic.document.documentElement.attrs['data-brand']).toBe('teal');
        expect(classic.store['ms365-brand-v1']).toBe('teal');

        const wine = bootContext({ 'ms365-brand-v1': 'wine' });
        expect(wine.document.documentElement.attrs['data-brand']).toBe('teal');

        const custom = bootContext({
            'ms365-brand-v1': 'custom',
            'ms365-brand-custom-v1': JSON.stringify({ brand1: '#ffff00', brand2: '#1d4ed8' })
        });
        expect(custom.document.documentElement.attrs['data-brand']).toBe('teal');
        expect(custom.document.documentElement.style.props['--brand1']).toBeUndefined();
    });
});
