import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const __dirname = dirname(fileURLToPath(import.meta.url));
const src = readFileSync(join(__dirname, '../src/shared/school-domain.js'), 'utf8');

function loadSchoolDomain(initial, pathname) {
    const storage = {
        _data: { ...(initial || {}) },
        getItem(k) {
            return Object.prototype.hasOwnProperty.call(this._data, k) ? this._data[k] : null;
        },
        setItem(k, v) {
            this._data[k] = String(v);
        },
        removeItem(k) {
            delete this._data[k];
        }
    };
    const created = [];
    const listeners = {};
    const document = {
        body: {
            getAttribute() {
                return null;
            },
            appendChild() {}
        },
        getElementById() {
            return null;
        },
        createElement(tag) {
            created.push(tag);
            return {
                classList: { add() {}, remove() {} },
                setAttribute() {},
                addEventListener() {},
                querySelector() {
                    return { addEventListener() {} };
                }
            };
        },
        addEventListener(type, fn) {
            listeners[type] = fn;
        }
    };
    const sandbox = {
        localStorage: storage,
        document,
        location: { pathname: pathname || '/tools/jahrgangsgruppen.html' },
        setTimeout(fn) {
            fn();
        }
    };
    sandbox.window = sandbox;
    sandbox.globalThis = sandbox;
    vm.runInNewContext(src, sandbox);
    if (typeof listeners.DOMContentLoaded === 'function') listeners.DOMContentLoaded();
    return { sandbox, storage, created };
}

describe('school-domain Stammdaten-Hinweis', () => {
    it('erkennt die Domain aus den kanonischen Stammdaten (v2), ohne v1-Spiegel', () => {
        const { sandbox, created } = loadSchoolDomain({
            'ms365-schooltool-data-v2': JSON.stringify({
                core: { domain: 'bg-musterstadt.at', schoolName: 'BG Musterstadt' }
            }),
            'ms365-school-email-domain-v1': 'ms365.schule'
        });
        expect(sandbox.ms365IsTenantSchoolDomainConfigured()).toBe(true);
        expect(sandbox.ms365GetSchoolDomainNoAt()).toBe('bg-musterstadt.at');
        expect(created).toEqual([]);
    });

    it('liest die Domain auch über ms365AppDataV2', () => {
        const { sandbox } = loadSchoolDomain({});
        sandbox.ms365AppDataV2 = {
            getContainer() {
                return { core: { domain: '@htl-beispiel.ac.at' } };
            }
        };
        expect(sandbox.ms365IsTenantSchoolDomainConfigured()).toBe(true);
        expect(sandbox.ms365GetSchoolDomainNoAt()).toBe('htl-beispiel.ac.at');
    });

    it('zeigt den Hinweis nur, wenn in den Stammdaten keine Domain steht', () => {
        const { sandbox, created } = loadSchoolDomain({
            'ms365-school-email-domain-v1': 'ms365.schule'
        });
        expect(sandbox.ms365IsTenantSchoolDomainConfigured()).toBe(false);
        expect(created.length).toBeGreaterThan(0);
    });

    it('akzeptiert weiterhin den alten v1-Spiegel', () => {
        const { sandbox, created } = loadSchoolDomain({
            'ms365-tenant-settings-v1': JSON.stringify({ domain: 'schule.example' })
        });
        expect(sandbox.ms365IsTenantSchoolDomainConfigured()).toBe(true);
        expect(sandbox.ms365GetSchoolDomainNoAt()).toBe('schule.example');
        expect(created).toEqual([]);
    });
});
