import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';
import { describe, expect, it } from 'vitest';
import { dlgToast } from '../src/shared/utils/dialog.js';
import { showToast } from '../src/shared/utils/dom.js';

const projectRoot = join(dirname(fileURLToPath(import.meta.url)), '..');

function loadAppDialog() {
    const code = readFileSync(join(projectRoot, 'src/shared/app-dialog.js'), 'utf8');
    const appended = [];
    const body = {
        appendChild(el) {
            appended.push(el);
            return el;
        },
        contains() {
            return true;
        }
    };
    const sandbox = {
        console,
        document: {
            body,
            createElement(tag) {
                const el = {
                    tagName: String(tag).toUpperCase(),
                    id: '',
                    className: '',
                    style: {},
                    children: [],
                    _attrs: {},
                    _html: '',
                    _text: '',
                    parentNode: null,
                    setAttribute(k, v) {
                        this._attrs[k] = v;
                    },
                    getAttribute(k) {
                        return this._attrs[k];
                    },
                    addEventListener() {},
                    removeEventListener() {},
                    querySelector(sel) {
                        if (sel.includes('title')) return { textContent: '' };
                        if (sel.includes('message')) return { textContent: '' };
                        if (sel.includes('icon')) {
                            return { className: '', innerHTML: '', textContent: '' };
                        }
                        if (sel.includes('prompt')) return { style: { display: 'none' } };
                        if (sel.includes('input-label-text')) return { textContent: '' };
                        if (sel.includes('input') && !sel.includes('label')) {
                            return { value: '', focus() {}, select() {} };
                        }
                        if (sel.includes('ok')) return { textContent: '', className: '', focus() {} };
                        if (sel.includes('cancel')) return { textContent: '', style: { display: '' } };
                        if (sel.includes('toast-title') || sel.includes('toast-msg')) {
                            return { textContent: '' };
                        }
                        if (sel.includes('toast-close')) return { addEventListener() {} };
                        return null;
                    },
                    classList: {
                        _c: new Set(),
                        add(c) {
                            this._c.add(c);
                        },
                        remove(c) {
                            this._c.delete(c);
                        },
                        contains(c) {
                            return this._c.has(c);
                        }
                    },
                    set innerHTML(v) {
                        this._html = v;
                    },
                    get innerHTML() {
                        return this._html;
                    },
                    set textContent(v) {
                        this._text = v;
                    },
                    get textContent() {
                        return this._text;
                    },
                    appendChild(child) {
                        child.parentNode = this;
                        this.children.push(child);
                        return child;
                    },
                    removeChild(child) {
                        this.children = this.children.filter((c) => c !== child);
                        child.parentNode = null;
                    }
                };
                return el;
            },
            addEventListener() {},
            removeEventListener() {}
        },
        requestAnimationFrame(fn) {
            fn();
        },
        setTimeout,
        clearTimeout
    };
    sandbox.window = sandbox;
    sandbox.document.body.contains = () => appended.length > 0;
    const ctx = createContext(sandbox);
    runInContext(code, ctx);
    return { api: sandbox, appended };
}

describe('Zentrale Toast/Dialog-API', () => {
    it('showToast und dlgToast nutzen ms365ShowToast', () => {
        const calls = [];
        globalThis.window = globalThis;
        window.ms365ShowToast = (msg, opts) => {
            calls.push({ msg, opts });
        };
        showToast('Listen-Check OK', { kind: 'success' });
        dlgToast('Listen-Check: bitte Protokoll prüfen', { kind: 'warning' });
        expect(calls).toHaveLength(2);
        expect(calls[0].opts.kind).toBe('success');
        expect(calls[1].opts.kind).toBe('warning');
    });

    it('app-dialog.js stellt Toast-API bereit und erkennt Warnungen', () => {
        const { api } = loadAppDialog();
        expect(typeof api.ms365ShowToast).toBe('function');
        expect(typeof api.ms365ToastOrAlert).toBe('function');
        expect(typeof api.ms365AppDialogAlert).toBe('function');

        const result = api.ms365ShowToast('Listen-Check: bitte Protokoll prüfen');
        expect(result).toHaveProperty('dismiss');
        expect(typeof result.dismiss).toBe('function');
        result.dismiss();
    });

    it('HTML-Seiten laden app-dialog.js', () => {
        const html = readFileSync(join(projectRoot, 'tools/sharepoint-liste-schularbeiten.html'), 'utf8');
        expect(html).toMatch(/app-dialog\.js/);
    });
});
