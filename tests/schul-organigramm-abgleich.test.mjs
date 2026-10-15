import { describe, expect, it } from 'vitest';
import {
    seedSchulOrganigrammFromStammdaten,
    setSchulOrganigrammAreaM365
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';
import {
    buildSchulOrgM365AbgleichDraft,
    diffGroupMembers,
    enrichSchulOrgM365AbgleichPlan,
    executeSchulOrgGraphAbgleich,
    buildAbgleichMailboxLauncher,
    emailsFromAreaPeople,
    resolveArgeLeiterDesiredEmails,
    resolveKlassenvorstandDesiredEmails,
    summarizeAbgleichCompare
} from '../src/tools/schul-organigramm/schul-organigramm-abgleich.js';
import {
    SCHUL_ORG_SEC_ARGE_LEITER,
    SCHUL_ORG_SEC_KLASSENVORSTAND
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';

describe('schul-organigramm-abgleich', () => {
    it('diffGroupMembers ergänzt nur fehlende', () => {
        const diff = diffGroupMembers(
            ['a@schule.at', 'b@schule.at', 'a@schule.at'],
            [{ mail: 'a@schule.at' }, { userPrincipalName: 'c@schule.at' }]
        );
        expect(diff.toAdd).toEqual(['b@schule.at']);
        expect(diff.already).toEqual(['a@schule.at']);
        expect(diff.desired).toEqual(['a@schule.at', 'b@schule.at']);
    });

    it('Draft: Gruppe anlegen + Postfach + Mitglieder-Sync', () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [{ code: 'SEK', name: 'Sekretariat', seatMode: 'multi' }],
            admin: [
                { role: 'Sekretariat', name: 'Anna', email: 'a@schule.at' },
                { role: 'Sekretariat', name: 'Ben', email: 'b@schule.at' }
            ],
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
            m365: {
                groupMode: 'create',
                groupKind: 'unified',
                groupId: null,
                mailboxMode: 'create',
                mailboxUpn: 'sekretariat@schule.at',
                mailboxLabel: 'Sekretariat'
            }
        });
        // Rolle unter dem neuen Bereich
        const sekRole = layout.nodes.find((n) => n.type === 'role' && n.refKey === 'SEK');
        if (sekRole) sekRole.parentId = 'sec_sek';

        layout.nodes.push({
            id: 'sec_bound',
            type: 'area',
            label: 'IT',
            parentId: null,
            order: 40,
            refKind: 'none',
            refKey: '',
            m365: {
                groupMode: 'match',
                groupKind: 'unified',
                groupId: 'gid-1',
                groupLabel: 'IT',
                mailboxMode: 'none'
            }
        });
        layout.nodes.push({
            id: 'person_it',
            type: 'person',
            label: 'Chris',
            parentId: 'sec_bound',
            order: 1,
            refKind: 'adminPerson',
            refKey: 'chris@schule.at'
        });
        layout = setSchulOrganigrammAreaM365(layout, 'sec_bound', {
            groupMode: 'match',
            groupKind: 'unified',
            groupId: 'gid-1',
            groupLabel: 'IT',
            mailboxMode: 'none'
        });

        const settings = {
            adminRoles: [{ code: 'SEK', name: 'Sekretariat', seatMode: 'multi' }],
            admin: [
                { role: 'Sekretariat', name: 'Anna', email: 'a@schule.at' },
                { role: 'Sekretariat', name: 'Ben', email: 'b@schule.at' }
            ],
            schulOrganigramm: layout
        };
        const draft = buildSchulOrgM365AbgleichDraft(settings);
        const types = draft.actions.map((a) => a.type).sort();
        expect(types).toContain('createGroup');
        expect(types).toContain('createMailbox');
        expect(types).toContain('syncMembers');

        const create = draft.actions.find((a) => a.id === 'createGroup:sec_sek');
        expect(create?.memberEmails?.sort()).toEqual(['a@schule.at', 'b@schule.at']);
        const mb = draft.actions.find((a) => a.id === 'createMailbox:sec_sek');
        expect(mb?.mailbox?.primarySmtp).toBe('sekretariat@schule.at');
        const sync = draft.actions.find((a) => a.id === 'syncMembers:sec_bound');
        expect(sync?.memberEmails).toEqual(['chris@schule.at']);
        expect(sync?.needsGraph).toBe(true);
    });

    it('enrich: noop wenn Mitglieder schon da', async () => {
        const draft = {
            actions: [
                {
                    id: 'syncMembers:x',
                    type: 'syncMembers',
                    channel: 'graph',
                    groupId: 'g1',
                    groupLabel: 'IT',
                    label: 'IT',
                    memberEmails: ['a@schule.at'],
                    selected: true,
                    needsGraph: true,
                    status: 'needsGraph'
                }
            ]
        };
        const progress = [];
        const enriched = await enrichSchulOrgM365AbgleichPlan(
            draft,
            {
                token: 't',
                fetchGroupMembers: async () => ({
                    items: [{ mail: 'a@schule.at' }]
                })
            },
            (ev) => progress.push(ev)
        );
        expect(enriched.actions[0].status).toBe('noop');
        expect(enriched.actions[0].selected).toBe(false);
        expect(enriched.actions[0].toAdd).toEqual([]);
        expect(enriched.actions[0].already).toEqual(['a@schule.at']);
        expect(progress).toHaveLength(1);
        expect(progress[0].done).toBe(1);
        expect(progress[0].total).toBe(1);
        const cmp = summarizeAbgleichCompare(enriched.actions);
        expect(cmp.counts.ok).toBe(1);
        expect(cmp.counts.missing).toBe(0);
    });

    it('summarizeAbgleichCompare trennt passt / fehlt / neu', () => {
        const cmp = summarizeAbgleichCompare([
            { type: 'syncMembers', status: 'noop', already: ['a@x.at'], toAdd: [] },
            { type: 'syncMembers', status: 'pending', already: [], toAdd: ['b@x.at'] },
            { type: 'createGroup', status: 'pending' },
            { type: 'createMailbox', status: 'pending' },
            { type: 'syncMembers', status: 'error' }
        ]);
        expect(cmp.counts.ok).toBe(1);
        expect(cmp.counts.missing).toBe(1);
        expect(cmp.counts.create).toBe(1);
        expect(cmp.counts.mailbox).toBe(1);
        expect(cmp.counts.errors).toBe(1);
    });

    it('execute createGroup + syncMembers', async () => {
        let layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: []
        });
        layout.nodes.push({
            id: 'sec_new',
            type: 'area',
            label: 'Neu',
            parentId: null,
            order: 10,
            refKind: 'none',
            refKey: '',
            m365: {
                groupMode: 'create',
                groupKind: 'security',
                groupId: null,
                mailboxMode: 'create',
                mailboxUpn: 'neu@schule.at',
                mailboxLabel: 'Neu'
            }
        });
        const settings = { schulOrganigramm: layout };
        const synced = [];
        const result = await executeSchulOrgGraphAbgleich(
            settings,
            [
                {
                    id: 'createGroup:sec_new',
                    type: 'createGroup',
                    channel: 'graph',
                    nodeId: 'sec_new',
                    selected: true,
                    groupKind: 'security',
                    displayName: 'Neu',
                    mailNickname: 'neu',
                    memberEmails: ['a@schule.at'],
                    ownerEmail: 'a@schule.at',
                    status: 'pending'
                }
            ],
            {
                token: 'tok',
                createGroupByKind: async () => ({
                    id: 'gid-new',
                    displayName: 'Neu',
                    mail: 'neu@schule.at'
                }),
                ensureOwners: async () => {},
                syncEmailsToGroup: async (_t, gid, emails) => {
                    synced.push({ gid, emails });
                    return { ok: emails.length, skip: 0, fail: 0 };
                },
                sleep: async () => {}
            }
        );
        expect(result.okCount).toBe(1);
        const node = result.layout.nodes.find((n) => n.id === 'sec_new');
        expect(node?.m365?.groupId).toBe('gid-new');
        expect(node?.m365?.mailboxUpn).toBe('neu@schule.at');
        expect(synced[0]?.emails).toEqual(['a@schule.at']);
    });

    it('Mailbox-Launcher erzeugt .cmd', () => {
        const launcher = buildAbgleichMailboxLauncher(
            [
                {
                    type: 'createMailbox',
                    label: 'Sekretariat',
                    mailbox: {
                        displayName: 'Sekretariat',
                        alias: 'sekretariat',
                        primarySmtp: 'sekretariat@schule.at',
                        delegates: ['a@schule.at']
                    }
                }
            ],
            'admin@schule.at'
        );
        expect(launcher.count).toBe(1);
        expect(launcher.cmd).toContain('@echo off');
        expect(launcher.cmd).toContain('New-Mailbox');
        expect(launcher.script).toContain('Admin-Konto wählen');
    });

    it('emailsFromAreaPeople dedupliziert', () => {
        expect(
            emailsFromAreaPeople([
                { email: 'A@Schule.at' },
                { email: 'a@schule.at' },
                { email: 'b@schule.at' }
            ])
        ).toEqual(['a@schule.at', 'b@schule.at']);
    });

    it('ARGE-Leiter-Draft nutzt Stammdaten-Köpfe', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [
                {
                    code: 'MAT',
                    name: 'Mathematik',
                    headName: 'Anna',
                    headEmail: 'anna@schule.at'
                },
                {
                    code: 'MUS',
                    name: 'Musik',
                    headName: 'Bert',
                    headEmail: 'bert@schule.at'
                },
                {
                    code: 'SPO',
                    name: 'Sport',
                    headName: 'Anna',
                    headEmail: 'anna@schule.at'
                }
            ],
            classes: []
        });
        const withM365 = setSchulOrganigrammAreaM365(layout, SCHUL_ORG_SEC_ARGE_LEITER, {
            groupMode: 'match',
            groupKind: 'unified',
            groupId: 'gid-arge-leiter',
            groupLabel: 'ARGE-Leiter',
            mailboxMode: 'none'
        });
        const draft = buildSchulOrgM365AbgleichDraft({
            arges: [
                { code: 'MAT', name: 'Mathematik', headEmail: 'anna@schule.at', headName: 'Anna' },
                { code: 'MUS', name: 'Musik', headEmail: 'bert@schule.at', headName: 'Bert' },
                { code: 'SPO', name: 'Sport', headEmail: 'anna@schule.at', headName: 'Anna' }
            ],
            schulOrganigramm: withM365
        });
        const sync = draft.actions.find((a) => a.id === 'syncMembers:' + SCHUL_ORG_SEC_ARGE_LEITER);
        expect(sync).toBeTruthy();
        expect(sync.memberEmails.sort()).toEqual(['anna@schule.at', 'bert@schule.at']);
        expect(sync.needsGraph).toBe(true);
    });

    it('resolveArgeLeiterDesiredEmails merged Köpfe + Graph-Owner', async () => {
        const resolved = await resolveArgeLeiterDesiredEmails(
            {
                arges: [
                    { code: 'MAT', headEmail: 'anna@schule.at', headName: 'Anna' },
                    { code: 'MUS', headEmail: 'bert@schule.at', headName: 'Bert' }
                ],
                setup: {
                    catalogLinks: [
                        { kind: 'arge', code: 'MAT', graphGroupId: 'g-mat' },
                        { kind: 'arge', code: 'MUS', graphGroupId: 'g-mus' }
                    ]
                }
            },
            {
                token: 't',
                fetchGroupOwners: async (_t, gid) => {
                    if (gid === 'g-mat') return [{ mail: 'anna@schule.at' }, { mail: 'chris@schule.at' }];
                    if (gid === 'g-mus') return [{ userPrincipalName: 'dana@schule.at' }];
                    return [];
                }
            }
        );
        expect(resolved.fromHeads).toBe(2);
        expect(resolved.fromOwners).toBe(2);
        expect(resolved.emails.sort()).toEqual([
            'anna@schule.at',
            'bert@schule.at',
            'chris@schule.at',
            'dana@schule.at'
        ]);
    });

    it('enrich: ARGE-Leiter-Soll um Owner erweitern', async () => {
        const draft = {
            actions: [
                {
                    id: 'syncMembers:' + SCHUL_ORG_SEC_ARGE_LEITER,
                    type: 'syncMembers',
                    channel: 'graph',
                    nodeId: SCHUL_ORG_SEC_ARGE_LEITER,
                    label: 'ARGE-Leiter',
                    groupId: 'gid-leiter',
                    groupLabel: 'ARGE-Leiter',
                    memberEmails: ['anna@schule.at'],
                    selected: true,
                    needsGraph: true,
                    status: 'needsGraph'
                }
            ]
        };
        const enriched = await enrichSchulOrgM365AbgleichPlan(draft, {
            token: 't',
            settings: {
                arges: [{ code: 'MAT', headEmail: 'anna@schule.at' }]
            },
            setup: {
                catalogLinks: [{ kind: 'arge', code: 'MAT', graphGroupId: 'g-mat' }]
            },
            fetchGroupMembers: async () => ({
                items: [{ mail: 'anna@schule.at' }]
            }),
            fetchGroupOwners: async () => [{ mail: 'owner@schule.at' }]
        });
        expect(enriched.actions[0].desired.sort()).toEqual(['anna@schule.at', 'owner@schule.at']);
        expect(enriched.actions[0].toAdd).toEqual(['owner@schule.at']);
        expect(enriched.actions[0].status).toBe('pending');
        expect(enriched.actions[0].detail).toMatch(/fehlt \+1/);
    });

    it('Klassenvorstände-Draft nutzt Stammdaten-Köpfe der Klassen', () => {
        const layout = seedSchulOrganigrammFromStammdaten({
            adminRoles: [],
            admin: [],
            arges: [],
            classes: [
                {
                    code: '4A',
                    name: '4A',
                    headName: 'Clara',
                    headEmail: 'clara@schule.at'
                },
                {
                    code: '4B',
                    name: '4B',
                    headName: 'Dirk',
                    headEmail: 'dirk@schule.at'
                },
                {
                    code: '5A',
                    name: '5A',
                    headName: 'Clara',
                    headEmail: 'clara@schule.at'
                }
            ]
        });
        const withM365 = setSchulOrganigrammAreaM365(layout, SCHUL_ORG_SEC_KLASSENVORSTAND, {
            groupMode: 'match',
            groupKind: 'unified',
            groupId: 'gid-kv',
            groupLabel: 'Klassenvorstände',
            mailboxMode: 'none'
        });
        const draft = buildSchulOrgM365AbgleichDraft({
            classes: [
                { code: '4A', name: '4A', headEmail: 'clara@schule.at', headName: 'Clara' },
                { code: '4B', name: '4B', headEmail: 'dirk@schule.at', headName: 'Dirk' },
                { code: '5A', name: '5A', headEmail: 'clara@schule.at', headName: 'Clara' }
            ],
            schulOrganigramm: withM365
        });
        const sync = draft.actions.find(
            (a) => a.id === 'syncMembers:' + SCHUL_ORG_SEC_KLASSENVORSTAND
        );
        expect(sync).toBeTruthy();
        expect(sync.memberEmails.sort()).toEqual(['clara@schule.at', 'dirk@schule.at']);
        expect(sync.needsGraph).toBe(true);
    });

    it('resolveKlassenvorstandDesiredEmails aus Klassen-Köpfen', () => {
        const resolved = resolveKlassenvorstandDesiredEmails({
            classes: [
                { code: '1A', headEmail: 'kv1@schule.at', headName: 'Eins' },
                { code: '1B', headEmail: 'kv2@schule.at', headName: 'Zwei' },
                { code: '2A', headEmail: 'kv1@schule.at', headName: 'Eins' }
            ]
        });
        expect(resolved.fromHeads).toBe(2);
        expect(resolved.emails.sort()).toEqual(['kv1@schule.at', 'kv2@schule.at']);
    });

    it('enrich: Klassenvorstände-Soll aus Stammdaten nachziehen', async () => {
        const draft = {
            actions: [
                {
                    id: 'syncMembers:' + SCHUL_ORG_SEC_KLASSENVORSTAND,
                    type: 'syncMembers',
                    channel: 'graph',
                    nodeId: SCHUL_ORG_SEC_KLASSENVORSTAND,
                    label: 'Klassenvorstände',
                    groupId: 'gid-kv',
                    groupLabel: 'Klassenvorstände',
                    memberEmails: [],
                    selected: true,
                    needsGraph: true,
                    status: 'needsGraph'
                }
            ]
        };
        const enriched = await enrichSchulOrgM365AbgleichPlan(draft, {
            token: 't',
            settings: {
                classes: [
                    { code: '4A', headEmail: 'clara@schule.at' },
                    { code: '4B', headEmail: 'dirk@schule.at' }
                ]
            },
            fetchGroupMembers: async () => ({
                items: [{ mail: 'clara@schule.at' }]
            })
        });
        expect(enriched.actions[0].desired.sort()).toEqual(['clara@schule.at', 'dirk@schule.at']);
        expect(enriched.actions[0].toAdd).toEqual(['dirk@schule.at']);
        expect(enriched.actions[0].status).toBe('pending');
        expect(enriched.actions[0].detail).toMatch(/fehlt \+1/);
    });
});
