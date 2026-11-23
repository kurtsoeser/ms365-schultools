import { describe, it, expect, afterEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { setFetchImpl } from '../src/shared/graph-client.js';
import {
    alignPrivateChannelOwnership,
    buildPrivateChannelMembersPayload,
    formatKlassenchatsGraphError,
    isBenignMemberError,
    isTeamMembershipPendingError,
    resolveChannelOwnerIds
} from '../src/tools/klassenchats/klassenchats-graph.js';
import { kvEmailFromPlan } from '../src/tools/klassenchats/klassenchats-logic.js';

const graphSrc = readFileSync(
    join(dirname(fileURLToPath(import.meta.url)), '../src/tools/klassenchats/klassenchats-graph.js'),
    'utf8'
);

describe('klassenchats-graph errors', () => {
    it('listChannelMembers nutzt kein $select in der URL (userId nur auf Subtyp)', () => {
        expect(graphSrc).toMatch(
            /encodeURIComponent\(cid\) \+ '\/members'/
        );
        expect(graphSrc).not.toMatch(
            /channels\/' \+\s*encodeURIComponent\(cid\) \+ '\s*\/members\?\$select/
        );
    });

    const parentRosterMsg =
        'POST /teams/x/channels: {"code":"Forbidden","message":"Failed to execute Templates backend request CreateChannelRequest. ... A private channel can only contain members of parent team, users 8:orgid:aaa42207-60f3-4537-8d99-bb23dd34d587 are not part of the parent roster","innerError":{"code":"AccessDenied"}}';

    it('erkennt parent-roster als Team-Mitgliedschaft noch nicht bereit', () => {
        expect(isTeamMembershipPendingError(new Error(parentRosterMsg))).toBe(true);
        expect(
            isTeamMembershipPendingError(
                new Error('User must be a member of the team before adding to private channel')
            )
        ).toBe(true);
        expect(isTeamMembershipPendingError(new Error('Channel.Create denied'))).toBe(false);
    });

    it('formatiert parent-roster verständlich auf Deutsch', () => {
        const text = formatKlassenchatsGraphError(new Error(parentRosterMsg));
        expect(text).toMatch(/nicht Mitglied des Lehrerteams/i);
        expect(text).toContain('aaa42207-60f3-4537-8d99-bb23dd34d587');
        expect(text.length).toBeLessThan(parentRosterMsg.length);
    });

    it('behandelt bereits-Mitglied als harmlos, nicht jedes 400', () => {
        expect(
            isBenignMemberError(new Error('400 One or more added object references already exist'))
        ).toBe(true);
        expect(isBenignMemberError(new Error('409 Conflict'))).toBe(true);
        expect(isBenignMemberError(new Error('400 Bad Request: invalid user'))).toBe(false);
        expect(isBenignMemberError(new Error('404 User does not exist'))).toBe(false);
    });

    it('setzt KV als Owner und andere als Mitglieder im Create-Payload', () => {
        const members = buildPrivateChannelMembersPayload(
            ['admin-id', 'kv-id'],
            ['t1', 't2', 'kv-id']
        );
        const byBind = (id) =>
            members.find((m) => m['user@odata.bind'].indexOf(id) !== -1);
        expect(byBind('admin-id').roles).toEqual(['owner']);
        expect(byBind('kv-id').roles).toEqual(['owner']);
        expect(byBind('t1').roles).toEqual([]);
        expect(byBind('t2').roles).toEqual([]);
        expect(members.filter((m) => m['user@odata.bind'].indexOf('kv-id') !== -1)).toHaveLength(1);
    });

    it('resolveChannelOwnerIds: Create inkl. Admin, Final nur KV', () => {
        const idByEmail = new Map([
            ['kv@school.at', 'kv-id'],
            ['a@school.at', 'a-id']
        ]);
        const plan = {
            members: [
                { email: 'a@school.at', role: 'lehrer' },
                { email: 'kv@school.at', role: 'kv' }
            ]
        };
        expect(kvEmailFromPlan(plan)).toBe('kv@school.at');
        expect(resolveChannelOwnerIds(idByEmail, plan, 'admin-id', { forCreate: true })).toEqual([
            'admin-id',
            'kv-id'
        ]);
        expect(resolveChannelOwnerIds(idByEmail, plan, 'admin-id', { forCreate: false })).toEqual([
            'kv-id'
        ]);
    });

    afterEach(() => {
        setFetchImpl(null);
    });

    it('alignPrivateChannelOwnership stuft KV hoch und andere herab', async () => {
        /** @type {Array<{ id: string, userId: string, roles: string[] }>} */
        let members = [
            { id: 'm-admin', userId: 'admin-id', roles: ['owner'] },
            { id: 'm-kv', userId: 'kv-id', roles: [] },
            { id: 'm-t1', userId: 't1-id', roles: ['owner'] }
        ];
        const patches = [];
        const prevWindow = globalThis.window;
        globalThis.window = {
            ms365GraphUnifiedGroups: {
                GRAPH_SCOPES: ['https://graph.microsoft.com/User.Read'],
                getGraphToken: async () => 'tok'
            }
        };
        setFetchImpl(async (url, init) => {
            const method = String((init && init.method) || 'GET').toUpperCase();
            const path = String(url || '').replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/, '');
            let body = null;
            if (init && init.body) {
                try {
                    body = JSON.parse(String(init.body));
                } catch {
                    body = null;
                }
            }
            if (method === 'GET' && path.indexOf('/members') !== -1) {
                return new Response(
                    JSON.stringify({ value: members.map((m) => ({ ...m, roles: m.roles.slice() })) }),
                    { status: 200, headers: { 'Content-Type': 'application/json' } }
                );
            }
            if (method === 'PATCH' && path.indexOf('/members/') !== -1) {
                const mid = decodeURIComponent(path.split('/members/')[1]);
                const m = members.find((x) => x.id === mid);
                const roles = body && Array.isArray(body.roles) ? body.roles.slice() : [];
                patches.push({ mid, roles });
                if (m) m.roles = roles;
                return new Response(JSON.stringify({ ...m }), {
                    status: 200,
                    headers: { 'Content-Type': 'application/json' }
                });
            }
            if (method === 'DELETE' && path.indexOf('/members/') !== -1) {
                const mid = decodeURIComponent(path.split('/members/')[1]);
                members = members.filter((x) => x.id !== mid);
                return new Response(null, { status: 204 });
            }
            return new Response(JSON.stringify({ error: { message: 'unexpected ' + method + ' ' + path } }), {
                status: 500,
                headers: { 'Content-Type': 'application/json' }
            });
        });
        try {
            const res = await alignPrivateChannelOwnership(
                {},
                'tok',
                'team',
                'chan',
                ['kv-id'],
                ['kv-id', 't1-id'],
                'admin-id'
            );
            expect(res.verifiedOwnerIds).toEqual(['kv-id']);
            expect(members.find((m) => m.userId === 'kv-id').roles).toEqual(['owner']);
            expect(members.find((m) => m.userId === 't1-id').roles).toEqual([]);
            expect(members.find((m) => m.userId === 'admin-id')).toBeFalsy();
            expect(patches.some((p) => p.mid === 'm-kv' && p.roles[0] === 'owner')).toBe(true);
        } finally {
            globalThis.window = prevWindow;
        }
    });
});
