import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { fetchAllPages, getGraphApi, odataEscape } from '../src/shared/graph-client.js';

describe('graph-client', () => {
    const prev = globalThis.window;

    beforeEach(() => {
        globalThis.window = {
            ms365GraphUnifiedGroups: {
                GRAPH_SCOPES: ['https://graph.microsoft.com/User.Read'],
                getGraphToken: vi.fn(async () => 'tok'),
                graphRequest: vi.fn(),
                graphJson: vi.fn(),
                sleep: (ms) => new Promise((r) => setTimeout(r, ms)),
                odataEscape: (s) => String(s).replace(/'/g, "''"),
                fetchGroupMembers: vi.fn()
            }
        };
    });

    afterEach(() => {
        globalThis.window = prev;
    });

    it('getGraphApi wirft ohne Shared-Modul', () => {
        globalThis.window = {};
        expect(() => getGraphApi()).toThrow(/nicht geladen/);
    });

    it('odataEscape escaped Quotes', () => {
        expect(odataEscape("O'Brien")).toBe("O''Brien");
    });

    it('fetchAllPages setzt truncated bei maxItems', async () => {
        let calls = 0;
        window.ms365GraphUnifiedGroups.graphJson = vi.fn(async () => {
            calls++;
            if (calls === 1) {
                return {
                    value: [{ id: '1' }, { id: '2' }],
                    '@odata.nextLink': '/page2'
                };
            }
            return { value: [{ id: '3' }] };
        });
        const r = await fetchAllPages('tok', '/groups', { maxItems: 2, maxPages: 10 });
        expect(r.items).toHaveLength(2);
        expect(r.truncated).toBe(true);
    });

    it('fetchAllPages truncated=false wenn fertig', async () => {
        window.ms365GraphUnifiedGroups.graphJson = vi.fn(async () => ({
            value: [{ id: 'a' }]
        }));
        const r = await fetchAllPages('tok', '/users', { maxItems: 100 });
        expect(r.items).toHaveLength(1);
        expect(r.truncated).toBe(false);
        expect(r.pages).toBe(1);
    });
});
