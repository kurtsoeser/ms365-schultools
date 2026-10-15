import { describe, expect, it } from 'vitest';
import {
    seedSchulOrganigrammFromStammdaten,
    setSchulOrganigrammAreaM365,
    NODE
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';
import {
    areaNeedsM365Provision,
    listOpenSchulOrgProvisions,
    markSchulOrgAreasForProvision,
    collectSchulOrgStructureDiff,
    healSchulOrgStructureFromLists,
    runSchulOrgProvisionBatch,
    provisionJobFromArea
} from '../src/tools/schul-organigramm/schul-organigramm-provision-diff.js';

describe('schul-organigramm-provision-diff', () => {
    it('erkennt offene Create-Jobs und markiert unbound Areas', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [],
            teachers: [],
            students: []
        });
        layout.nodes.push({
            id: 'sec_sek',
            type: 'area',
            label: 'Sekretariat',
            parentId: 'sec_verwaltung',
            order: 99,
            refKind: 'none',
            refKey: '',
            m365: { groupMode: 'create', groupKind: 'unified', groupId: null }
        });
        layout = setSchulOrganigrammAreaM365(layout, 'sec_sek', {
            groupMode: 'create',
            groupKind: 'mailSecurity',
            groupId: null,
            mailboxMode: 'none'
        });
        const node = layout.nodes.find((n) => n.id === 'sec_sek');
        expect(areaNeedsM365Provision(node)).toBe(true);
        const jobs = listOpenSchulOrgProvisions(layout);
        expect(jobs.some((j) => j.nodeId === 'sec_sek')).toBe(true);
        expect(jobs.find((j) => j.nodeId === 'sec_sek')?.groupKind).toBe('mailSecurity');

        let unbound = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: []
        });
        unbound.nodes.push({
            id: 'sec_it',
            type: 'area',
            label: 'IT',
            parentId: null,
            order: 50,
            refKind: 'none',
            refKey: ''
        });
        unbound = markSchulOrgAreasForProvision(unbound, ['sec_it'], 'security');
        const it = unbound.nodes.find((n) => n.id === 'sec_it');
        expect(it?.m365?.groupMode).toBe('create');
        expect(it?.m365?.groupKind).toBe('security');
        expect(listOpenSchulOrgProvisions(unbound).map((j) => j.nodeId)).toContain('sec_it');
    });

    it('Diff: fehlende Klassen auf der Landkarte / Orphans', () => {
        const settings = {
            adminRoles: [{ code: 'SEK', name: 'Sekretariat', tier: 'verwaltung' }],
            admin: [],
            arges: [],
            classes: [
                { code: '1A', name: '1A' },
                { code: '2B', name: '2B' }
            ],
            teachers: [],
            students: [],
            schulOrganigramm: {
                version: 2,
                nodes: [
                    {
                        id: 'sec_verwaltung',
                        type: 'area',
                        label: 'Verwaltung',
                        parentId: null,
                        order: 0,
                        refKind: 'none',
                        refKey: ''
                    },
                    {
                        id: 'role_GHOST',
                        type: NODE.ROLE,
                        label: 'Geist',
                        parentId: 'sec_verwaltung',
                        order: 1,
                        refKind: 'adminRole',
                        refKey: 'GHOST'
                    }
                ]
            }
        };
        const diff = collectSchulOrgStructureDiff(settings);
        expect(diff.counts.missingInLayout).toBeGreaterThanOrEqual(2);
        expect(diff.missingInLayout.some((x) => x.kind === 'class' && x.key === '1A')).toBe(true);
        expect(diff.orphanInLayout.some((x) => x.key === 'GHOST')).toBe(true);

        const healed = healSchulOrgStructureFromLists(settings);
        // Klassen liegen als Bereichskachel (type area + refKind class) bzw. Legacy type class
        expect(
            healed.layout.nodes.some(
                (n) =>
                    n.refKey === '1A' &&
                    (n.type === 'class' || n.refKind === 'class')
            )
        ).toBe(true);
        expect(healed.added).toBeGreaterThanOrEqual(1);
    });

    it('Batch legt Gruppen an und schreibt groupId', async () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: []
        });
        layout = markSchulOrgAreasForProvision(
            {
                ...layout,
                nodes: layout.nodes.concat([
                    {
                        id: 'sec_batch',
                        type: 'area',
                        label: 'Batch Test',
                        parentId: null,
                        order: 80,
                        refKind: 'none',
                        refKey: ''
                    }
                ])
            },
            ['sec_batch'],
            'unified'
        );
        const jobs = listOpenSchulOrgProvisions(layout).filter((j) => j.nodeId === 'sec_batch');
        expect(jobs).toHaveLength(1);
        const job = Object.assign(
            {},
            provisionJobFromArea(layout.nodes.find((n) => n.id === 'sec_batch')),
            {
                memberEmails: ['a@schule.at', 'b@schule.at'],
                ownerEmail: 'a@schule.at'
            }
        );
        /** @type {string[][]} */
        const synced = [];
        /** @type {string[][]} */
        const owners = [];
        const created = await runSchulOrgProvisionBatch(layout, [job], {
            token: 'tok',
            createGroupByKind: async (_t, kind, name, nick) => ({
                id: 'gid-batch-1',
                displayName: name,
                mail: nick + '@schule.at',
                groupTypes: kind === 'unified' ? ['Unified'] : []
            }),
            ensureOwners: async (_t, gid, list) => {
                owners.push([gid, ...list]);
            },
            syncEmailsToGroup: async (_t, gid, emails) => {
                synced.push([gid, ...emails]);
                return { ok: emails.length, skip: 0, fail: 0 };
            },
            sleep: async () => {}
        });
        expect(created.okCount).toBe(1);
        expect(created.failCount).toBe(0);
        const node = created.layout.nodes.find((n) => n.id === 'sec_batch');
        expect(node?.m365?.groupMode).toBe('match');
        expect(node?.m365?.groupId).toBe('gid-batch-1');
        expect(owners).toEqual([['gid-batch-1', 'a@schule.at']]);
        expect(synced).toEqual([['gid-batch-1', 'a@schule.at', 'b@schule.at']]);
    });
});
