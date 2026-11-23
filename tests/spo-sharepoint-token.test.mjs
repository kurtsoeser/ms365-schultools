import { readFileSync } from 'node:fs';
import { createContext, runInContext } from 'node:vm';
import { describe, expect, it } from 'vitest';

function loadApi(hooks) {
    const sandbox = {
        ms365AuthAcquireTokenSilent: hooks.silent,
        ms365AuthAcquireTokenFresh: hooks.fresh,
        ms365AuthAcquireTokenPopup: hooks.popup,
        ms365AuthAcquireToken: hooks.redirect
    };
    sandbox.window = sandbox;
    sandbox.globalThis = sandbox;
    const code = readFileSync(
        new URL('../src/tools/sharepoint/spo-graph-shared.js', import.meta.url),
        'utf8'
    );
    runInContext(code, createContext(sandbox));
    return sandbox.ms365SpoGraph;
}

function resourceOf(scope) {
    return new URL(scope).origin;
}

describe('acquireSharePointToken', () => {
    it('nimmt still AllSites.Write, ohne Graph-Scope', async () => {
        const seen = [];
        const api = loadApi({
            silent: async (scopes) => {
                seen.push(scopes);
                if (String(scopes[0]).endsWith('/AllSites.Write')) return 'write-silent';
                throw new Error('nicht erteilt');
            },
            fresh: async () => {
                throw new Error('soll nicht interaktiv werden');
            },
            popup: async () => {
                throw new Error('soll nicht interaktiv werden');
            }
        });
        const tok = await api.acquireSharePointToken('kurtrocks.sharepoint.com');
        expect(tok).toBe('write-silent');
        expect(seen).toHaveLength(1);
        expect(seen[0]).toEqual(['https://kurtrocks.sharepoint.com/AllSites.Write']);
        expect(resourceOf(seen[0][0])).toBe('https://kurtrocks.sharepoint.com');
    });

    it('fällt still auf AllSites.Read zurück, immer nur eine Ressource', async () => {
        const seen = [];
        const api = loadApi({
            silent: async (scopes) => {
                seen.push(scopes.slice());
                if (String(scopes[0]).endsWith('/AllSites.Read')) return 'read-silent';
                throw new Error('interaction_required');
            }
        });
        const tok = await api.acquireSharePointToken('https://kurtrocks.sharepoint.com/sites/schule');
        expect(tok).toBe('read-silent');
        expect(seen.map((s) => s[0])).toEqual([
            'https://kurtrocks.sharepoint.com/AllSites.Write',
            'https://kurtrocks.sharepoint.com/AllSites.Read'
        ]);
        for (const scopes of seen) {
            expect(scopes).toHaveLength(1);
            expect(scopes.join(' ')).not.toContain('graph.microsoft.com');
        }
    });

    it('öffnet nach fehlendem Silent nur SharePoint, nicht Graph und SharePoint zusammen', async () => {
        const seen = [];
        const api = loadApi({
            silent: async () => {
                throw new Error('interaction_required');
            },
            fresh: async (scopes) => {
                seen.push(scopes.slice());
                throw new Error('interaction_required');
            },
            popup: async (scopes) => {
                seen.push(scopes.slice());
                return 'popup-' + scopes[0];
            }
        });
        const tok = await api.acquireSharePointToken('kurtrocks.sharepoint.com');
        expect(tok).toBe('popup-https://kurtrocks.sharepoint.com/AllSites.Write');
        expect(seen.length).toBeGreaterThan(0);
        for (const scopes of seen) {
            expect(scopes).toHaveLength(1);
            expect(scopes.join(' ')).not.toContain('graph.microsoft.com');
            expect(resourceOf(scopes[0])).toBe('https://kurtrocks.sharepoint.com');
        }
    });

    it('bricht nach Abbruch ab und fragt den zweiten Scope nicht mehr', async () => {
        let popups = 0;
        const api = loadApi({
            silent: async () => {
                throw new Error('interaction_required');
            },
            fresh: async () => {
                throw new Error('interaction_required');
            },
            popup: async () => {
                popups += 1;
                const err = new Error('Anmeldung abgebrochen');
                err.errorCode = 'user_cancelled';
                throw err;
            }
        });
        await expect(api.acquireSharePointToken('kurtrocks.sharepoint.com')).rejects.toThrow(
            /SharePoint-Token fehlt/
        );
        expect(popups).toBe(1);
    });
});

describe('spo-graph-shared → graph-client', () => {
    it('leitet graphJson an window.ms365GraphClient weiter (bestehende Signatur)', async () => {
        const seen = [];
        const sandbox = {
            ms365AuthAcquireTokenSilent: async () => null,
            ms365GraphClient: {
                graphJson: async (method, path, token, body, extra, opts) => {
                    seen.push({ method, path, token, body, extra, opts });
                    return { ok: true, path };
                },
                graphRequest: async () => new Response('{}', { status: 200 }),
                sleep: (ms) => new Promise((r) => setTimeout(r, ms))
            }
        };
        sandbox.window = sandbox;
        sandbox.globalThis = sandbox;
        const code = readFileSync(
            new URL('../src/tools/sharepoint/spo-graph-shared.js', import.meta.url),
            'utf8'
        );
        runInContext(code, createContext(sandbox));
        const api = sandbox.ms365SpoGraph;
        const data = await api.graphJson('GET', '/sites/root', 'tok', undefined, 'beta');
        expect(data).toEqual({ ok: true, path: '/sites/root' });
        expect(seen).toHaveLength(1);
        expect(seen[0].method).toBe('GET');
        expect(seen[0].path).toBe('/sites/root');
        expect(seen[0].token).toBe('tok');
        expect(seen[0].opts).toEqual({ version: 'beta' });
    });
});
