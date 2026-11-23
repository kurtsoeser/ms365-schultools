import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
    applyAuthMenuItAudience,
    ensureHeaderShellSwitch,
    syncHeaderShellSwitch
} from '../src/shared/dashboard-auth-menu-policy.js';

const projectRoot = join(dirname(fileURLToPath(import.meta.url)), '..');

function el(id, extras) {
    const state = Object.assign(
        {
            id: id || '',
            hidden: false,
            className: '',
            dataset: {},
            title: '',
            innerHTML: '',
            parentElement: null,
            nextElementSibling: null,
            children: [],
            attrs: /** @type {Record<string, string>} */ ({}),
            classList: {
                _set: new Set(),
                contains(c) {
                    return this._set.has(c);
                },
                add(c) {
                    this._set.add(c);
                },
                remove(c) {
                    this._set.delete(c);
                },
                toggle(c, force) {
                    if (force === true) this._set.add(c);
                    else if (force === false) this._set.delete(c);
                    else if (this._set.has(c)) this._set.delete(c);
                    else this._set.add(c);
                }
            },
            setAttribute(k, v) {
                this.attrs[k] = String(v);
            },
            getAttribute(k) {
                return Object.prototype.hasOwnProperty.call(this.attrs, k) ? this.attrs[k] : null;
            },
            removeAttribute(k) {
                delete this.attrs[k];
            },
            querySelectorAll(sel) {
                const s = String(sel || '');
                if (s === '[data-dash-layout]') {
                    return this.children.filter((c) => c.getAttribute('data-dash-layout'));
                }
                const m = s.match(/^\[data-dash-layout="([^"]+)"\]$/);
                if (m) {
                    return this.children.filter((c) => c.getAttribute('data-dash-layout') === m[1]);
                }
                return [];
            },
            querySelector(sel) {
                return this.querySelectorAll(sel)[0] || null;
            },
            appendChild(child) {
                child.parentElement = this;
                this.children.push(child);
                return child;
            },
            insertBefore(child, before) {
                child.parentElement = this;
                const idx = this.children.indexOf(before);
                if (idx >= 0) this.children.splice(idx, 0, child);
                else this.children.push(child);
                return child;
            },
            addEventListener() {}
        },
        extras || {}
    );
    return state;
}

describe('Kontomenü v11 / Header Portal-IT-Chip', () => {
    /** @type {Map<string, any>} */
    let byId;

    beforeEach(() => {
        byId = new Map();
        const actions = el('adminAppTopActions');
        const auth = el('ms365AuthWidget');
        actions.appendChild(auth);
        byId.set('adminAppTopActions', actions);
        byId.set('ms365AuthWidget', auth);

        const tenantRow = el('legacyTenant', {
            className: 'ms365-auth-menu__ctx-row ms365-auth-menu__ctx-row--tenant',
            hidden: false
        });
        const actionLog = el('ms365AuthActionLogLink', { hidden: false });
        const itTour = el('ms365AuthItTourBtn', { hidden: false });
        byId.set('ms365AuthActionLogLink', actionLog);
        byId.set('ms365AuthItTourBtn', itTour);

        vi.stubGlobal('document', {
            getElementById: (id) => byId.get(id) || null,
            querySelector: (sel) => {
                if (sel === '.ms365-auth-menu__ctx-row--tenant') return tenantRow;
                return null;
            },
            querySelectorAll: () => [],
            createElement: (tag) => {
                const node = el('', { tagName: String(tag || '').toUpperCase() });
                Object.defineProperty(node, 'id', {
                    configurable: true,
                    get() {
                        return this._id || '';
                    },
                    set(v) {
                        this._id = String(v || '');
                        if (this._id) byId.set(this._id, node);
                    }
                });
                // emulate button children created via innerHTML assignment for our switcher
                Object.defineProperty(node, 'innerHTML', {
                    configurable: true,
                    get() {
                        return this._html || '';
                    },
                    set(v) {
                        this._html = String(v || '');
                        this.children = [];
                        const re = /data-dash-layout="([^"]+)"/g;
                        let m;
                        while ((m = re.exec(this._html))) {
                            const btn = el('', {
                                tagName: 'BUTTON',
                                className: 'ms365-header-shell-switch__btn'
                            });
                            btn.setAttribute('data-dash-layout', m[1]);
                            btn.setAttribute('aria-pressed', 'false');
                            node.appendChild(btn);
                        }
                    }
                });
                return node;
            }
        });
        vi.stubGlobal('window', {
            ms365AuthIsLoggedIn: () => true,
            addEventListener: () => {},
            location: { pathname: '/index.html' }
        });
    });

    afterEach(() => {
        vi.unstubAllGlobals();
    });

    it('blendet Tenant-ID dauerhaft aus und Aktionsprotokoll nur für IT', () => {
        applyAuthMenuItAudience(false);
        expect(document.querySelector('.ms365-auth-menu__ctx-row--tenant').hidden).toBe(true);
        expect(document.getElementById('ms365AuthActionLogLink').hidden).toBe(true);
        expect(document.getElementById('ms365AuthItTourBtn').hidden).toBe(true);

        applyAuthMenuItAudience(true);
        expect(document.querySelector('.ms365-auth-menu__ctx-row--tenant').hidden).toBe(true);
        expect(document.getElementById('ms365AuthActionLogLink').hidden).toBe(false);
        expect(document.getElementById('ms365AuthItTourBtn').hidden).toBe(false);
    });

    it('baut den Portal/IT-Chip im Header und zeigt ihn nur für Schul-IT', () => {
        const switcher = ensureHeaderShellSwitch();
        expect(switcher).toBeTruthy();
        expect(switcher.id).toBe('ms365HeaderShellSwitch');
        byId.set('ms365HeaderShellSwitch', switcher);
        expect(switcher.querySelectorAll('[data-dash-layout]').length).toBe(2);

        syncHeaderShellSwitch({ loggedIn: true, isIt: false }, 'preview-lehrer');
        expect(switcher.hidden).toBe(true);

        syncHeaderShellSwitch({ loggedIn: true, isIt: true }, 'full');
        expect(switcher.hidden).toBe(false);
        const itBtn = switcher.querySelector('[data-dash-layout="full"]');
        const portalBtn = switcher.querySelector('[data-dash-layout="preview-lehrer"]');
        expect(itBtn.classList.contains('is-active')).toBe(true);
        expect(portalBtn.classList.contains('is-active')).toBe(false);

        syncHeaderShellSwitch({ loggedIn: true, isIt: true }, 'preview-schueler');
        expect(portalBtn.classList.contains('is-active')).toBe(true);
        expect(itBtn.classList.contains('is-active')).toBe(false);
    });

    it('Auth-Menü-Markup v11 enthält Dunkles Design und keine Tenant-ID mehr', () => {
        const auth = readFileSync(join(projectRoot, 'src/shared/msal-auth-ui.js'), 'utf8');
        expect(auth).toContain("data-ms365-auth-menu-v', '12'");
        expect(auth).toContain('Dunkles Design');
        expect(auth).not.toContain('ms365AuthCtxTenant');
        expect(auth).not.toContain('schoolYearAddBtn');
        expect(auth).toContain('ensureHeaderShellSwitch');

        const chrome = readFileSync(join(projectRoot, 'src/shared/app-header-chrome.js'), 'utf8');
        expect(chrome).toContain('ms365HeaderShellSwitch');
    });
});
