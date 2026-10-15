import { describe, expect, it } from 'vitest';
import {
    isSchulOrgLocalPlaceholderEmail,
    personHasM365UserLink,
    nodeM365LinkFlags,
    normalizeSchulOrganigramm,
    setSchulOrganigrammAreaM365,
    syncSchulOrganigrammWithStammdaten
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
        expect(nodeM365LinkFlags({ m365: { groupMode: 'match', groupId: 'g1' } })).toEqual({
            hasGroup: true,
            hasMailbox: false,
            groupId: 'g1',
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
        ).toMatchObject({ hasGroup: false, hasMailbox: true, linked: true });
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
        ).toMatchObject({ hasGroup: false, hasMailbox: false, linked: false });
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
        expect(nodeM365LinkFlags(chat, 'gid-wrong-chat')).toMatchObject({
            hasGroup: false,
            linked: false
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
});
