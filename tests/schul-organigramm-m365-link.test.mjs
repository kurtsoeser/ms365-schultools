import { describe, expect, it } from 'vitest';
import {
    isSchulOrgLocalPlaceholderEmail,
    personHasM365UserLink,
    nodeM365LinkFlags,
    normalizeSchulOrganigramm,
    setSchulOrganigrammAreaM365,
    syncSchulOrganigrammWithStammdaten,
    applyClassGroupLinksToSchulOrganigramm,
    seedSchulOrganigrammFromStammdaten,
    mergeSchulOrganigrammLivePositions,
    findSchulOrgNodesUsingM365Group,
    findSchulOrgNodesUsingM365Mailbox,
    schulOrgM365ConflictLabel,
    schulOrgGroupKindIcon,
    schulOrgGroupKindLabel,
    normalizeSchulOrgGroupKind
} from '../src/tools/schul-organigramm/schul-organigramm-logic.js';

describe('schul-organigramm M365-Link-Flags', () => {
    it('erkennt lokale Platzhalter-Mails', () => {
        expect(isSchulOrgLocalPlaceholderEmail('anna@local.schul-org')).toBe(true);
        expect(isSchulOrgLocalPlaceholderEmail('')).toBe(true);
        expect(isSchulOrgLocalPlaceholderEmail('eva@schule.at')).toBe(false);
    });

    it('personHasM365UserLink nur bei echter E-Mail', () => {
        expect(personHasM365UserLink({ email: 'eva@schule.at' })).toBe(true);
        expect(personHasM365UserLink({ email: 'x@local.schul-org' })).toBe(false);
        expect(personHasM365UserLink({ name: 'Nur Name' })).toBe(false);
    });

    it('nodeM365LinkFlags für Gruppe und Postfach', () => {
        expect(nodeM365LinkFlags({ m365: { groupMode: 'match', groupId: 'g1' } })).toMatchObject({
            hasGroup: true,
            hasMailbox: false,
            hasChat: false,
            groupId: 'g1',
            displayKind: 'unified',
            linked: true
        });
        expect(
            nodeM365LinkFlags(
                {
                    m365: {
                        groupMode: 'none',
                        mailboxMode: 'match',
                        mailboxUpn: 'sek@schule.at'
                    }
                },
                ''
            )
        ).toMatchObject({ hasGroup: false, hasMailbox: true, hasChat: false, linked: true });
        // Explizit „Keine Gruppe“ → Tenant-Match-Fallback ignorieren (Schulleitung-Fall)
        expect(
            nodeM365LinkFlags(
                {
                    id: 'sec_schulleitung',
                    refKind: 'audience',
                    refKey: 'schulleitung',
                    m365: { groupMode: 'none', groupId: null },
                    label: 'Schulleitung'
                },
                'matched-schulleitung-gid'
            )
        ).toMatchObject({ hasGroup: false, hasMailbox: false, hasChat: false, linked: false });
        // Create ohne ID = noch nicht verknüpft
        expect(
            nodeM365LinkFlags({ m365: { groupMode: 'create', groupId: null } }, 'fallback')
        ).toMatchObject({ hasGroup: false, linked: false });
        // Ohne m365-Block: nur wenn resolveGroupId wirklich etwas liefert (z. B. Sammelgruppe)
        expect(nodeM365LinkFlags({ label: 'Lehrkräfte' }, 'matched-lehrer-gid')).toMatchObject({
            hasGroup: true,
            groupId: 'matched-lehrer-gid',
            linked: true
        });
        // Ohne m365 und ohne Auflösung → kein Icon (Audience/Schulleitung ohne Match im Dialog)
        expect(
            nodeM365LinkFlags({
                id: 'sec_schulleitung',
                refKind: 'audience',
                refKey: 'schulleitung',
                label: 'Schulleitung'
            })
        ).toMatchObject({ hasGroup: false, linked: false });
        expect(nodeM365LinkFlags({ label: 'Ohne' })).toMatchObject({ linked: false });
    });

    it('nodeM365LinkFlags erkennt Teams-Gruppenchat über meta.chatId', () => {
        expect(
            nodeM365LinkFlags({
                refKind: 'klassenchatClass',
                meta: { chatId: '19:chat-abc@thread.v2' },
                m365: { groupMode: 'none', groupKind: 'groupChat' }
            })
        ).toMatchObject({
            hasChat: true,
            hasGroup: false,
            showChatMark: true,
            chatId: '19:chat-abc@thread.v2',
            displayKind: 'groupChat',
            linked: true
        });
        // KlassenChat-Kachel: Icon auch ohne chatId (Objekttyp)
        expect(
            nodeM365LinkFlags({
                id: 'klassenchat_class_1A',
                refKind: 'klassenchatClass',
                refKey: '1A',
                m365: { groupMode: 'none', groupKind: 'groupChat' }
            })
        ).toMatchObject({
            hasChat: false,
            showChatMark: true,
            displayKind: 'groupChat',
            linked: true
        });
        // Live-Auflösung über 3. Parameter
        expect(
            nodeM365LinkFlags(
                {
                    id: 'klassenchat_class_1A',
                    refKind: 'klassenchatClass',
                    refKey: '1A'
                },
                '',
                '19:live-chat@thread.v2'
            )
        ).toMatchObject({
            hasChat: true,
            chatId: '19:live-chat@thread.v2',
            showChatMark: true,
            linked: true
        });
        expect(normalizeSchulOrgGroupKind('gruppenchat')).toBe('groupChat');
        expect(schulOrgGroupKindLabel('groupChat')).toBe('Teams-Gruppenchat');
        expect(schulOrgGroupKindIcon('groupChat')).toBe('bi-chat-dots-fill');
        expect(schulOrgGroupKindIcon('security')).toBe('bi-shield-lock');
        expect(schulOrgGroupKindIcon('mailSecurity')).toBe('bi-envelope-paper');
        expect(schulOrgGroupKindIcon('distribution')).toBe('bi-mailbox');
        expect(schulOrgGroupKindIcon('unified')).toBe('bi-people-fill');
    });

    it('Gruppe lösen bleibt nach Normalize/Sync erhalten', () => {
        let layout = normalizeSchulOrganigramm({
            nodes: [
                {
                    id: 'sec_custom',
                    type: 'area',
                    label: 'Mein Bereich',
                    parentId: null,
                    order: 1,
                    m365: {
                        groupMode: 'match',
                        groupKind: 'unified',
                        groupId: 'gid-accident',
                        groupLabel: 'Zufallsgruppe',
                        mailboxMode: 'none'
                    }
                },
                {
                    id: 'klassenchat_class_1A',
                    type: 'area',
                    label: 'DEMO Klasse 1A',
                    parentId: 'sec_klassenchats',
                    order: 2,
                    tone: 'lehrer',
                    refKind: 'klassenchatClass',
                    refKey: '1A',
                    m365: {
                        groupMode: 'match',
                        groupKind: 'unified',
                        groupId: 'gid-wrong-chat',
                        groupLabel: 'Falsche Gruppe',
                        mailboxMode: 'none'
                    }
                }
            ]
        });
        layout = setSchulOrganigrammAreaM365(layout, 'sec_custom', {
            groupMode: 'none',
            groupKind: 'unified',
            groupId: null,
            mailboxMode: 'none'
        });
        layout = setSchulOrganigrammAreaM365(layout, 'klassenchat_class_1A', {
            groupMode: 'none',
            groupKind: 'unified',
            groupId: null,
            mailboxMode: 'none'
        });
        const node = layout.nodes.find((n) => n.id === 'sec_custom');
        expect(node.m365.groupMode).toBe('none');
        expect(node.m365.groupId).toBeNull();
        expect(nodeM365LinkFlags(node, 'should-ignore')).toMatchObject({
            hasGroup: false,
            linked: false
        });
        const chat = layout.nodes.find((n) => n.id === 'klassenchat_class_1A');
        expect(chat.m365.groupMode).toBe('none');
        // Entra-Gruppe gelöst: kein hasGroup trotz resolvedGroupId – Chat-Icon bleibt (KlassenChat)
        expect(nodeM365LinkFlags(chat, 'gid-wrong-chat')).toMatchObject({
            hasGroup: false,
            hasChat: false,
            showChatMark: true,
            linked: true
        });

        // Seed mit Kursteam-m365 darf die gelöste Area nicht wieder verknüpfen
        const synced = syncSchulOrganigrammWithStammdaten(layout, {
            teachers: [],
            classes: [{ code: '1A', name: 'DEMO Klasse 1A' }],
            arges: [],
            adminRoles: [],
            admin: []
        });
        const again = synced.layout.nodes.find((n) => n.id === 'sec_custom');
        expect(again.m365.groupMode).toBe('none');
        expect(again.m365.groupId).toBeNull();
        const chatAgain = synced.layout.nodes.find((n) => n.id === 'klassenchat_class_1A');
        expect(chatAgain?.m365?.groupMode).toBe('none');
        expect(chatAgain?.m365?.groupId).toBeNull();
    });

    it('applyClassGroupLinksToSchulOrganigramm übernimmt Klassengruppen-Matches', () => {
        const settings = {
            teachers: [],
            classes: [
                { code: '1A', name: 'DEMO Klasse 1A', year: '2031' },
                { code: '2B', name: 'DEMO Klasse 2B', year: '2030' }
            ],
            arges: [],
            adminRoles: [],
            admin: []
        };
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const applied = applyClassGroupLinksToSchulOrganigramm(layout, settings, {
            classTeams: [
                {
                    classCode: '1A',
                    abschlussJahr: '2031',
                    graphGroupId: 'gid-1a',
                    displayName: 'Klasse 1A'
                }
            ],
            classGroupMatchByKey: {
                '2B': { groupId: 'gid-2b', displayName: 'Klasse 2B' }
            }
        });
        expect(applied.linked).toBe(2);
        expect(applied.updated).toBe(2);
        const c1 = applied.layout.nodes.find((n) => n.id === 'class_1A');
        const c2 = applied.layout.nodes.find((n) => n.id === 'class_2B');
        expect(c1?.m365).toMatchObject({
            groupMode: 'match',
            groupId: 'gid-1a',
            groupLabel: 'Klasse 1A'
        });
        expect(c2?.m365).toMatchObject({
            groupMode: 'match',
            groupId: 'gid-2b'
        });
        expect(nodeM365LinkFlags(c1)).toMatchObject({ hasGroup: true, linked: true });
        expect(nodeM365LinkFlags(c2, 'gid-2b')).toMatchObject({ hasGroup: true, linked: true });
    });

    it('applyClassGroupLinksToSchulOrganigramm respektiert groupMode=none', () => {
        const settings = {
            teachers: [],
            classes: [{ code: '1A', name: '1A', year: '2031' }],
            arges: [],
            adminRoles: [],
            admin: []
        };
        let layout = seedSchulOrganigrammFromStammdaten(settings);
        layout = setSchulOrganigrammAreaM365(layout, 'class_1A', {
            groupMode: 'none',
            groupKind: 'unified',
            groupId: null,
            mailboxMode: 'none'
        });
        const applied = applyClassGroupLinksToSchulOrganigramm(layout, settings, {
            classTeams: [{ classCode: '1A', abschlussJahr: '2031', graphGroupId: 'gid-1a' }]
        });
        expect(applied.updated).toBe(0);
        const c1 = applied.layout.nodes.find((n) => n.id === 'class_1A');
        expect(c1?.m365?.groupMode).toBe('none');
        expect(c1?.m365?.groupId).toBeNull();
        expect(nodeM365LinkFlags(c1, 'gid-1a')).toMatchObject({ hasGroup: false, linked: false });
    });

    it('nodeM365LinkFlags: match ohne eigene ID nutzt resolvedGroupId', () => {
        expect(
            nodeM365LinkFlags(
                {
                    type: 'area',
                    refKind: 'class',
                    m365: { groupMode: 'match', groupId: null, groupKind: 'unified' }
                },
                'gid-from-teams'
            )
        ).toMatchObject({ hasGroup: true, groupId: 'gid-from-teams', linked: true });
    });

    it('findSchulOrgNodesUsingM365Group findet Doppelverknüpfungen', () => {
        const layout = normalizeSchulOrganigramm({
            nodes: [
                {
                    id: 'class_4A',
                    type: 'area',
                    label: 'DEMO Klasse 4A',
                    refKind: 'class',
                    parentId: 'sec_schueler',
                    m365: { groupMode: 'match', groupId: 'gid-shared', groupKind: 'unified' }
                },
                {
                    id: 'klassenchat_class_4A',
                    type: 'area',
                    label: 'DEMO Klasse 4A',
                    refKind: 'klassenchatClass',
                    parentId: 'sec_klassenchats',
                    m365: { groupMode: 'match', groupId: 'gid-shared', groupKind: 'unified' }
                },
                {
                    id: 'class_1A',
                    type: 'area',
                    label: '1A',
                    refKind: 'class',
                    parentId: 'sec_schueler',
                    m365: { groupMode: 'match', groupId: 'gid-other', groupKind: 'unified' }
                },
                {
                    id: 'sec_free',
                    type: 'area',
                    label: 'Gelöst',
                    parentId: null,
                    m365: { groupMode: 'none', groupId: 'gid-shared' }
                }
            ]
        });
        const hits = findSchulOrgNodesUsingM365Group(layout, 'gid-shared');
        expect(hits.map((h) => h.id).sort()).toEqual(['class_4A', 'klassenchat_class_4A']);
        expect(schulOrgM365ConflictLabel(hits.find((h) => h.id === 'klassenchat_class_4A'))).toMatch(
            /KlassenChat/
        );
        expect(
            findSchulOrgNodesUsingM365Group(layout, 'gid-shared', { excludeNodeId: 'class_4A' }).map(
                (h) => h.id
            )
        ).toEqual(['klassenchat_class_4A']);
    });

    it('findSchulOrgNodesUsingM365Mailbox findet doppelte Postfächer', () => {
        const layout = normalizeSchulOrganigramm({
            nodes: [
                {
                    id: 'sec_a',
                    type: 'area',
                    label: 'Sekretariat',
                    parentId: null,
                    m365: {
                        groupMode: 'none',
                        mailboxMode: 'match',
                        mailboxUpn: 'sek@schule.at'
                    }
                },
                {
                    id: 'sec_b',
                    type: 'area',
                    label: 'Kopie',
                    parentId: null,
                    m365: {
                        groupMode: 'none',
                        mailboxMode: 'match',
                        mailboxUpn: 'sek@schule.at'
                    }
                }
            ]
        });
        expect(findSchulOrgNodesUsingM365Mailbox(layout, 'sek@schule.at')).toHaveLength(2);
        expect(
            findSchulOrgNodesUsingM365Mailbox(layout, 'sek@schule.at', { excludeNodeId: 'sec_a' })
        ).toHaveLength(1);
    });

    it('Live-Merge überschreibt gelöste M365-Gruppe nicht (Gruppe lösen → Remount)', () => {
        const saved = normalizeSchulOrganigramm({
            nodes: [
                {
                    id: 'sec_custom',
                    type: 'area',
                    label: 'Mein Bereich',
                    parentId: null,
                    order: 1,
                    x: 10,
                    y: 20,
                    m365: {
                        groupMode: 'none',
                        groupKind: 'unified',
                        groupId: null,
                        mailboxMode: 'none'
                    }
                }
            ]
        });
        const live = normalizeSchulOrganigramm({
            nodes: [
                {
                    id: 'sec_custom',
                    type: 'area',
                    label: 'Mein Bereich',
                    parentId: null,
                    order: 1,
                    x: 40,
                    y: 50,
                    m365: {
                        groupMode: 'match',
                        groupKind: 'unified',
                        groupId: 'gid-old-mistake',
                        groupLabel: 'Falsche Gruppe',
                        mailboxMode: 'none'
                    }
                }
            ]
        });
        const merged = mergeSchulOrganigrammLivePositions(saved, live);
        const node = merged.nodes.find((n) => n.id === 'sec_custom');
        expect(node.x).toBe(40);
        expect(node.y).toBe(50);
        expect(node.m365.groupMode).toBe('none');
        expect(node.m365.groupId).toBeNull();
        expect(nodeM365LinkFlags(node, 'gid-old-mistake')).toMatchObject({
            hasGroup: false,
            linked: false
        });
    });
});
