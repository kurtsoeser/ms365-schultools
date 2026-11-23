import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { clientKindForScopes } from '../src/shared/auth-login-scopes.js';
import {
    fetchAllPages,
    getDefaultGraphScopes,
    getGraphApi,
    graphJson,
    graphRequest,
    noteGraphAction,
    odataEscape,
    resolveGraphUrl,
    resolveRetryWaitMs,
    setFetchImpl,
    setJournalHook,
    shouldRetryGraphStatus
} from '../src/shared/graph-client.js';
import { IT_CLIENT_MISSING_ERROR, resolveItClientId, resolveItMsalConfig } from '../src/shared/it-app-auth.js';

function jwtWithScp(scp) {
    const payload = Buffer.from(JSON.stringify({ scp }), 'utf8').toString('base64url');
    return 'e30.' + payload + '.sig';
}

function jsonResponse(status, body, headers = {}) {
    const h = new Headers(headers);
    return new Response(typeof body === 'string' ? body : JSON.stringify(body), { status, headers: h });
}

describe('graph-client', () => {
    const prevWindow = globalThis.window;

    beforeEach(() => {
        setFetchImpl(null);
        setJournalHook(null);
        globalThis.window = {
            ms365GraphUnifiedGroups: {
                GRAPH_SCOPES: ['https://graph.microsoft.com/User.Read'],
                getGraphToken: vi.fn(async () => 'tok'),
                sleep: (ms) => new Promise((r) => setTimeout(r, ms)),
                odataEscape: (s) => String(s).replace(/'/g, "''"),
                fetchGroupMembers: vi.fn()
            }
        };
    });

    afterEach(() => {
        setFetchImpl(null);
        setJournalHook(null);
        globalThis.window = prevWindow;
        vi.useRealTimers();
    });

    it('getGraphApi wirft ohne Shared-Modul', () => {
        globalThis.window = {};
        expect(() => getGraphApi()).toThrow(/nicht geladen/);
    });

    it('odataEscape escaped Quotes', () => {
        expect(odataEscape("O'Brien")).toBe("O''Brien");
    });

    it('resolveGraphUrl baut v1.0 und beta', () => {
        expect(resolveGraphUrl('/me')).toBe('https://graph.microsoft.com/v1.0/me');
        expect(resolveGraphUrl('users')).toBe('https://graph.microsoft.com/v1.0/users');
        expect(resolveGraphUrl('/sites/root', 'beta')).toBe('https://graph.microsoft.com/beta/sites/root');
        expect(resolveGraphUrl('https://graph.microsoft.com/v1.0/groups')).toContain('/groups');
    });

    it('shouldRetryGraphStatus erkennt 429 und 5xx', () => {
        expect(shouldRetryGraphStatus(429)).toBe(true);
        expect(shouldRetryGraphStatus(503)).toBe(true);
        expect(shouldRetryGraphStatus(500)).toBe(true);
        expect(shouldRetryGraphStatus(404)).toBe(false);
        expect(shouldRetryGraphStatus(400)).toBe(false);
    });

    it('resolveRetryWaitMs bevorzugt Retry-After', () => {
        const res = jsonResponse(429, {}, { 'Retry-After': '7' });
        expect(resolveRetryWaitMs(res, 0)).toBe(7000);
    });

    it('resolveRetryWaitMs nutzt Exponential Backoff ohne Header', () => {
        const res = jsonResponse(503, {});
        expect(resolveRetryWaitMs(res, 0, { backoffBaseMs: 400, backoffCapMs: 10000 })).toBe(400);
        expect(resolveRetryWaitMs(res, 1, { backoffBaseMs: 400, backoffCapMs: 10000 })).toBe(800);
        expect(resolveRetryWaitMs(res, 10, { backoffBaseMs: 400, backoffCapMs: 10000 })).toBe(10000);
    });

    it('resolveRetryWaitMs capped OneNote auf 8s', () => {
        const res = jsonResponse(429, {}, { 'Retry-After': '60' });
        expect(
            resolveRetryWaitMs(res, 0, { url: 'https://graph.microsoft.com/v1.0/users/x/onenote/pages' })
        ).toBe(8000);
    });

    it('graphRequest retried 429 mit Retry-After', async () => {
        vi.useFakeTimers();
        let calls = 0;
        setFetchImpl(async () => {
            calls++;
            if (calls === 1) return jsonResponse(429, { error: 'throttle' }, { 'Retry-After': '1' });
            return jsonResponse(200, { ok: true });
        });
        const p = graphRequest('GET', '/me', 'tok');
        await vi.advanceTimersByTimeAsync(1000);
        const res = await p;
        expect(res.status).toBe(200);
        expect(calls).toBe(2);
    });

    it('graphRequest retried 503 mit Backoff', async () => {
        vi.useFakeTimers();
        let calls = 0;
        setFetchImpl(async () => {
            calls++;
            if (calls < 3) return jsonResponse(503, { error: 'busy' });
            return jsonResponse(200, { value: [] });
        });
        const p = graphRequest('GET', '/groups', 'tok', undefined, undefined, {
            backoffBaseMs: 100,
            backoffCapMs: 1000
        });
        await vi.advanceTimersByTimeAsync(100);
        await vi.advanceTimersByTimeAsync(200);
        const res = await p;
        expect(res.status).toBe(200);
        expect(calls).toBe(3);
    });

    it('graphJson wirft mit status bei Fehler', async () => {
        setFetchImpl(async () => jsonResponse(404, { error: { message: 'missing' } }));
        await expect(graphJson('GET', '/users/x', 'tok')).rejects.toMatchObject({
            status: 404,
            message: expect.stringContaining('missing')
        });
    });

    it('graphJson journal-Hook bei Erfolg und Fehler', async () => {
        const entries = [];
        setJournalHook((e) => entries.push(e));
        setFetchImpl(async () => jsonResponse(200, { id: '1' }));
        await graphJson('PATCH', '/groups/1', 'tok', { displayName: 'A' }, undefined, {
            journal: true,
            tool: 'test',
            summary: 'rename'
        });
        expect(entries).toHaveLength(1);
        expect(entries[0].result).toBe('ok');
        expect(entries[0].tool).toBe('test');

        setFetchImpl(async () => jsonResponse(400, { error: { message: 'bad' } }));
        await expect(
            graphJson('POST', '/groups', 'tok', {}, undefined, { journal: { tool: 'test', action: 'create' } })
        ).rejects.toBeTruthy();
        expect(entries).toHaveLength(2);
        expect(entries[1].result).toBe('error');
    });

    it('noteGraphAction fällt auf ms365ActionLog zurück', () => {
        const append = vi.fn();
        globalThis.window.ms365ActionLog = { append };
        noteGraphAction({ tool: 't', action: 'a', summary: 's' });
        expect(append).toHaveBeenCalledOnce();
    });

    it('fetchAllPages setzt truncated bei maxItems', async () => {
        let calls = 0;
        setFetchImpl(async () => {
            calls++;
            if (calls === 1) {
                return jsonResponse(200, {
                    value: [{ id: '1' }, { id: '2' }],
                    '@odata.nextLink': 'https://graph.microsoft.com/v1.0/page2'
                });
            }
            return jsonResponse(200, { value: [{ id: '3' }] });
        });
        const r = await fetchAllPages('tok', '/groups', { maxItems: 2, maxPages: 10 });
        expect(r.items).toHaveLength(2);
        expect(r.truncated).toBe(true);
    });

    it('getDefaultGraphScopes bleibt bei User.Read', () => {
        expect(getDefaultGraphScopes()).toEqual(['https://graph.microsoft.com/User.Read']);
        globalThis.window = {};
        expect(getDefaultGraphScopes()).toEqual(['https://graph.microsoft.com/User.Read']);
        expect(getDefaultGraphScopes().join(' ')).not.toContain('ReadWrite.All');
    });

    it('POST /groups holt auch bei vorhandenem Write.All ein neues Token (IT-App)', async () => {
        const writeTok = jwtWithScp('User.Read Group.ReadWrite.All');
        globalThis.window.ms365GraphUnifiedGroups.getGraphToken.mockClear();
        setFetchImpl(async () => jsonResponse(201, { id: 'g1' }));
        await graphJson('POST', '/groups', writeTok, { displayName: 'Klasse' });
        expect(globalThis.window.ms365GraphUnifiedGroups.getGraphToken).toHaveBeenCalled();
        const arg = globalThis.window.ms365GraphUnifiedGroups.getGraphToken.mock.calls.at(-1)[0];
        expect(clientKindForScopes(arg)).toBe('it');
    });

    it('GET /groups bleibt auf der Portal-App', () => {
        expect(clientKindForScopes(['https://graph.microsoft.com/Group.Read.All'])).toBe('portal');
        expect(
            clientKindForScopes(['https://graph.microsoft.com/User.Read', 'https://graph.microsoft.com/RoleManagement.Read.Directory'])
        ).toBe('portal');
    });

    it('POST /groups fordert Group.ReadWrite.All nach, nicht User.ReadWrite.All', async () => {
        const readTok = jwtWithScp('User.Read Group.Read.All');
        setFetchImpl(async () => jsonResponse(201, { id: 'g1' }));
        await graphJson('POST', '/groups', readTok, { displayName: 'Klasse' });
        const calls = globalThis.window.ms365GraphUnifiedGroups.getGraphToken.mock.calls;
        const arg = calls[calls.length - 1][0];
        expect(arg).toContain('https://graph.microsoft.com/Group.ReadWrite.All');
        expect(arg).not.toContain('https://graph.microsoft.com/User.ReadWrite.All');
    });

    it('GET /groups mit Lesetoken fordert kein Write.All nach', async () => {
        const tok = jwtWithScp('User.Read Group.Read.All');
        globalThis.window.ms365GraphUnifiedGroups.getGraphToken.mockClear();
        setFetchImpl(async () => jsonResponse(200, { value: [] }));
        await graphJson('GET', '/groups', tok);
        expect(globalThis.window.ms365GraphUnifiedGroups.getGraphToken).not.toHaveBeenCalled();
    });

    it('POST /users fordert User.ReadWrite.All ohne Group.ReadWrite.All', async () => {
        const tok = jwtWithScp('User.Read');
        setFetchImpl(async () => jsonResponse(201, { id: 'u1' }));
        await graphJson('POST', '/users', tok, { displayName: 'Neu' });
        const calls = globalThis.window.ms365GraphUnifiedGroups.getGraphToken.mock.calls;
        const arg = calls[calls.length - 1][0];
        expect(arg).toContain('https://graph.microsoft.com/User.ReadWrite.All');
        expect(arg).not.toContain('https://graph.microsoft.com/Group.ReadWrite.All');
    });

    it('fetchAllPages truncated=false wenn fertig', async () => {
        setFetchImpl(async () => jsonResponse(200, { value: [{ id: 'a' }] }));
        const r = await fetchAllPages('tok', '/users', { maxItems: 100 });
        expect(r.items).toHaveLength(1);
        expect(r.truncated).toBe(false);
        expect(r.pages).toBe(1);
    });
});

describe('it-app client id', () => {
    it('liest itClientId aus der Config und fällt nicht auf die Portal-ID zurück', () => {
        const env = {
            MS365_MSAL_CONFIG: {
                clientId: 'e1d877c3-004c-4040-8c3b-81a59e0c7050',
                itClientId: 'a1f27f94-78c2-4503-9603-192c86924842',
                authority: 'https://login.microsoftonline.com/organizations',
                redirectUri: 'https://app.ms365.schule/ms365-schooltool.html'
            },
            document: { querySelector: () => null }
        };
        expect(resolveItClientId(env)).toBe('a1f27f94-78c2-4503-9603-192c86924842');
        expect(resolveItMsalConfig(env).authority).toBe('https://login.microsoftonline.com/organizations');
        expect(resolveItMsalConfig(env).redirectUri).toBe('https://app.ms365.schule/ms365-schooltool.html');
        expect(resolveItMsalConfig(env).clientId).not.toBe(env.MS365_MSAL_CONFIG.clientId);
    });

    it('nutzt das Meta-Tag, wenn itClientId leer ist', () => {
        const env = {
            MS365_MSAL_CONFIG: { clientId: 'e1d877c3-004c-4040-8c3b-81a59e0c7050', itClientId: '' },
            document: {
                querySelector: (sel) =>
                    sel === 'meta[name="ms365-graph-it-client-id"]'
                        ? { getAttribute: () => 'a1f27f94-78c2-4503-9603-192c86924842' }
                        : null
            }
        };
        expect(resolveItClientId(env)).toBe('a1f27f94-78c2-4503-9603-192c86924842');
    });

    it('wirft deutsch, wenn die IT-Client-ID fehlt, und nennt nicht die Portal-App als Ersatz', () => {
        const env = {
            MS365_MSAL_CONFIG: { clientId: 'e1d877c3-004c-4040-8c3b-81a59e0c7050' },
            document: { querySelector: () => null }
        };
        expect(resolveItClientId(env)).toBe('');
        expect(() => resolveItMsalConfig(env)).toThrow(IT_CLIENT_MISSING_ERROR);
        try {
            resolveItMsalConfig(env);
        } catch (e) {
            expect(e.code).toBe('ms365_it_client_missing');
            expect(e.message).toMatch(/nicht über die Portal-App/);
        }
    });
});
