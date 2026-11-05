import { describe, it, expect } from 'vitest';
import {
    defaultChatNamePattern,
    normalizeChatNamePattern,
    buildChatTopicFromPattern,
    displayNameForTarget,
    buildKvByKlasse,
    buildClassChatPlans,
    buildManualClassChatPlan,
    parseMemberEmailsText,
    sanitizeTeamsChatTopic,
    sanitizeTeamsChannelDisplayName,
    summarizePlans,
    upsertClassChatItem,
    existingByKlasseFromState,
    expandKlasseToKnownCodes,
    isCombinedClassCell,
    normalizeClassChatsState,
    normalizeTargetKind,
    kvEmailFromPlan,
    normalizeExtraMemberEmails
} from '../src/tools/klassenchats/klassenchats-logic.js';

describe('klassenchats-logic', () => {
    it('ersetzt Doppelpunkt im Topic für Teams Graph', () => {
        expect(sanitizeTeamsChatTopic('DEMO SJ26-27 | 1A | Lehrer:innen')).toBe(
            'DEMO SJ26-27 | 1A | Lehrer·innen'
        );
        const topic = buildChatTopicFromPattern(
            [
                { type: 'text', value: 'DEMO ' },
                { type: 'yearPrefix' },
                { type: 'text', value: ' | ' },
                { type: 'klasse' },
                { type: 'text', value: ' | Lehrer:innen' }
            ],
            { yearPrefix: 'SJ26-27', klasse: '1A' }
        );
        expect(topic).toBe('DEMO SJ26-27 | 1A | Lehrer·innen');
    });

    it('baut Standard-Topic mit Schuljahr', () => {
        const topic = buildChatTopicFromPattern(defaultChatNamePattern(), {
            yearPrefix: 'SJ26',
            klasse: '1A'
        });
        expect(topic).toBe('SJ26 | 1A | Lehrer');
    });

    it('normalisiert Pattern und ignoriert unbekannte Tokens', () => {
        const p = normalizeChatNamePattern([
            { type: 'yearPrefix' },
            { type: 'fach' },
            { type: 'text', value: '-' },
            { type: 'klasse' }
        ]);
        expect(p.map((t) => t.type)).toEqual(['yearPrefix', 'text', 'klasse']);
        expect(buildChatTopicFromPattern(p, { yearPrefix: 'SJ25', klasse: '2B' })).toBe('SJ25-2B');
    });

    it('ergänzt KV und filtert Klassen mit &lt; 2 Mitgliedern', () => {
        const belegung = {
            rows: [
                { klasse: '1A', lehrerCode: 'A', lehrerEmail: 'a@school.at', fach: 'D' },
                { klasse: '1A', lehrerCode: 'B', lehrerEmail: 'b@school.at', fach: 'M' },
                { klasse: '2B', lehrerCode: 'C', lehrerEmail: 'c@school.at', fach: 'D' }
            ]
        };
        const kv = buildKvByKlasse([
            { code: '1A', headEmail: 'kv1a@school.at', headName: 'KV Eins' },
            { code: '2B', headEmail: 'c@school.at', headName: 'Gleich C' }
        ]);
        const plans = buildClassChatPlans(belegung, kv, {
            yearPrefix: 'SJ26',
            namePattern: defaultChatNamePattern()
        });
        const p1 = plans.find((p) => p.klasse === '1A');
        const p2 = plans.find((p) => p.klasse === '2B');
        expect(p1.eligible).toBe(true);
        expect(p1.memberEmails).toContain('kv1a@school.at');
        expect(p1.memberEmails).toHaveLength(3);
        expect(kvEmailFromPlan(p1)).toBe('kv1a@school.at');
        expect(p1.topic).toBe('SJ26 | 1A | Lehrer');
        // 2B: nur C (+ KV ist dieselbe Mail) → 1 Mitglied → skip
        expect(p2.eligible).toBe(false);
        expect(p2.memberEmails).toHaveLength(1);
        expect(kvEmailFromPlan(p2)).toBe('c@school.at');
    });

    it('nimmt Zusatzpersonen (z. B. Direktion) in alle Klassenpläne auf', () => {
        const belegung = {
            rows: [
                { klasse: '1A', lehrerEmail: 'a@school.at', lehrerCode: 'A', fach: 'D' },
                { klasse: '1A', lehrerEmail: 'b@school.at', lehrerCode: 'B', fach: 'M' }
            ]
        };
        const plans = buildClassChatPlans(belegung, new Map(), {
            yearPrefix: 'SJ26',
            namePattern: defaultChatNamePattern(),
            extraMemberEmails: ['direktion@school.at', 'a@school.at']
        });
        const p1 = plans.find((p) => p.klasse === '1A');
        expect(p1.memberEmails).toContain('direktion@school.at');
        expect(p1.members.filter((m) => m.role === 'extra')).toHaveLength(1);
        expect(p1.members.find((m) => m.email === 'a@school.at').role).toBe('lehrer');
        expect(normalizeExtraMemberEmails('x@y.at; z@y.at')).toEqual(['x@y.at', 'z@y.at']);
        const st = normalizeClassChatsState({
            extraMemberEmails: ['dir@school.at'],
            items: []
        });
        expect(st.extraMemberEmails).toEqual(['dir@school.at']);
    });

    it('setzt Status neu/sync anhand vorhandener Chat-IDs und Topic', () => {
        const belegung = {
            rows: [
                { klasse: '1A', lehrerEmail: 'a@x.at', lehrerCode: 'A', fach: 'D' },
                { klasse: '1A', lehrerEmail: 'b@x.at', lehrerCode: 'B', fach: 'M' }
            ]
        };
        const existing = existingByKlasseFromState({
            items: [{ klasse: '1A', chatId: 'chat-1', topic: 'ALT | 1A | Lehrer' }]
        });
        const plans = buildClassChatPlans(belegung, new Map(), {
            yearPrefix: 'SJ26',
            namePattern: defaultChatNamePattern(),
            existingByKlasse: existing
        });
        expect(plans[0].status).toBe('sync');
        expect(plans[0].chatId).toBe('chat-1');
    });

    it('upsertClassChatItem schreibt und aktualisiert', () => {
        let state = upsertClassChatItem(null, {
            klasse: '1A',
            chatId: 'c1',
            topic: 'SJ26 | 1A | Lehrer',
            memberEmails: ['a@x.at', 'b@x.at'],
            yearPrefix: 'SJ26'
        });
        expect(state.items).toHaveLength(1);
        state = upsertClassChatItem(state, {
            klasse: '1A',
            chatId: 'c1',
            topic: 'SJ26 | 1A | Team',
            memberEmails: ['a@x.at', 'b@x.at', 'c@x.at']
        });
        expect(state.items).toHaveLength(1);
        expect(state.items[0].topic).toBe('SJ26 | 1A | Team');
        expect(state.items[0].memberEmails).toHaveLength(3);
        const sum = summarizePlans(
            buildClassChatPlans(
                {
                    rows: [
                        { klasse: '1A', lehrerEmail: 'a@x.at', fach: 'D' },
                        { klasse: '1A', lehrerEmail: 'b@x.at', fach: 'M' }
                    ]
                },
                new Map(),
                { yearPrefix: 'SJ26' }
            )
        );
        expect(sum.eligible).toBe(1);
    });

    it('löst Mehrklassen-Belegung auf Stammdaten-Klassen auf', () => {
        const stammdaten = [
            { code: '1AK' },
            { code: '1BK' },
            { code: '1CK' },
            { code: '1A' }
        ];
        expect(isCombinedClassCell('1AK1BK1CK')).toBe(true);
        expect(expandKlasseToKnownCodes('1AK1BK1CK', stammdaten)).toEqual(['1AK', '1BK', '1CK']);
        expect(expandKlasseToKnownCodes('1AK~1BK', stammdaten)).toEqual(['1AK', '1BK']);
        expect(isCombinedClassCell('2HMA/2HMB')).toBe(true);
        expect(expandKlasseToKnownCodes('2HMA/2HMB', [{ code: '2HMA' }, { code: '2HMB' }])).toEqual([
            '2HMA',
            '2HMB'
        ]);

        const belegung = {
            rows: [
                { klasse: '1AK1BK1CK', lehrerCode: 'X', lehrerEmail: 'x@school.at', fach: 'D' },
                { klasse: '1AK1BK1CK', lehrerCode: 'Y', lehrerEmail: 'y@school.at', fach: 'M' },
                { klasse: '1A', lehrerCode: 'Z', lehrerEmail: 'z@school.at', fach: 'D' }
            ]
        };
        const kv = buildKvByKlasse([
            { code: '1AK', headEmail: 'kvak@school.at', headName: 'KV AK' },
            { code: '1BK', headEmail: 'kvak@school.at', headName: 'KV AK' },
            { code: '1CK', headEmail: 'kvck@school.at', headName: 'KV CK' }
        ]);
        const plans = buildClassChatPlans(belegung, kv, {
            yearPrefix: 'SJ26',
            namePattern: defaultChatNamePattern(),
            knownClassCodes: stammdaten
        });
        expect(plans.find((p) => p.klasse === '1AK1BK1CK')).toBeUndefined();
        const pak = plans.find((p) => p.klasse === '1AK');
        const pbk = plans.find((p) => p.klasse === '1BK');
        const pck = plans.find((p) => p.klasse === '1CK');
        expect(pak.memberEmails).toContain('x@school.at');
        expect(pak.memberEmails).toContain('y@school.at');
        expect(pak.memberEmails).toContain('kvak@school.at');
        expect(pbk.memberEmails).toEqual(pak.memberEmails);
        expect(pck.memberEmails).toContain('kvck@school.at');
        expect(pck.memberEmails).toContain('x@school.at');
        expect(plans.find((p) => p.klasse === '1A').eligible).toBe(false);
    });

    it('baut manuellen Plan inkl. KV und Freitext-Mails', () => {
        expect(parseMemberEmailsText('a@x.at, b@x.at\nc@x.at')).toEqual([
            'a@x.at',
            'b@x.at',
            'c@x.at'
        ]);
        const kv = buildKvByKlasse([{ code: '3C', headEmail: 'kv@school.at', headName: 'KV' }]);
        const plan = buildManualClassChatPlan({
            klasse: '3C',
            membersText: 'a@school.at',
            includeKv: true,
            yearPrefix: 'SJ26',
            namePattern: defaultChatNamePattern(),
            kvByKlasse: kv,
            teacherDirectory: [{ code: 'A', email: 'a@school.at', name: 'Anna' }]
        });
        expect(plan.eligible).toBe(true);
        expect(plan.memberEmails).toContain('kv@school.at');
        expect(plan.memberEmails).toContain('a@school.at');
        expect(plan.topic).toBe('SJ26 | 3C | Lehrer');
        expect(plan.kind).toBe('groupChat');
        expect(plan.source).toBe('manual');
    });

    it('sanitized Kanalnamen (kein |, max. 50 Zeichen)', () => {
        expect(sanitizeTeamsChannelDisplayName('SJ26 | 1A | Lehrer')).toBe('SJ26 · 1A · Lehrer');
        expect(sanitizeTeamsChannelDisplayName('A:B#C')).toBe('A · B · C');
        const long = 'X'.repeat(60);
        expect(sanitizeTeamsChannelDisplayName(long).length).toBeLessThanOrEqual(50);
        expect(displayNameForTarget('SJ26 | 1A | Lehrer', 'privateChannel')).toBe(
            'SJ26 · 1A · Lehrer'
        );
        expect(normalizeTargetKind('privateChannel')).toBe('privateChannel');
        expect(normalizeTargetKind('chat')).toBe('groupChat');
    });

    it('plant private Kanäle nur mit Lehrerteam und getrennt von Gruppenchats', () => {
        const belegung = {
            rows: [
                { klasse: '1A', lehrerEmail: 'a@x.at', lehrerCode: 'A', fach: 'D' },
                { klasse: '1A', lehrerEmail: 'b@x.at', lehrerCode: 'B', fach: 'M' }
            ]
        };
        const withoutTeam = buildClassChatPlans(belegung, new Map(), {
            yearPrefix: 'SJ26',
            targetKind: 'privateChannel'
        });
        expect(withoutTeam[0].eligible).toBe(false);
        expect(withoutTeam[0].skipReason).toMatch(/Lehrerteam/i);
        expect(withoutTeam[0].topic).toBe('SJ26 · 1A · Lehrer');

        const state = upsertClassChatItem(null, {
            klasse: '1A',
            kind: 'groupChat',
            chatId: 'chat-1',
            topic: 'SJ26 | 1A | Lehrer'
        });
        const withChannel = upsertClassChatItem(state, {
            klasse: '1A',
            kind: 'privateChannel',
            channelId: 'chan-1',
            teamId: 'team-1',
            topic: 'SJ26 · 1A · Lehrer'
        });
        expect(withChannel.items).toHaveLength(2);

        const existingChat = existingByKlasseFromState(withChannel, { targetKind: 'groupChat' });
        expect(existingChat.get('1A').chatId).toBe('chat-1');
        const existingCh = existingByKlasseFromState(withChannel, {
            targetKind: 'privateChannel',
            teamId: 'team-1'
        });
        expect(existingCh.get('1A').channelId).toBe('chan-1');

        const plans = buildClassChatPlans(belegung, new Map(), {
            yearPrefix: 'SJ26',
            targetKind: 'privateChannel',
            teamId: 'team-1',
            existingByKlasse: existingCh
        });
        expect(plans[0].eligible).toBe(true);
        expect(plans[0].status).toBe('ok');
        expect(plans[0].channelId).toBe('chan-1');
    });

    it('normalisiert classChats-State mit Zieltyp und Lehrerteam', () => {
        const n = normalizeClassChatsState({
            targetKind: 'privateChannel',
            lehrerTeamId: 'tid',
            lehrerTeamDisplayName: 'Lehrkräfte',
            yearPrefix: 'SJ26',
            items: [
                {
                    klasse: '2B',
                    kind: 'privateChannel',
                    channelId: 'c2',
                    teamId: 'tid',
                    topic: 'SJ26 · 2B · Lehrer'
                }
            ]
        });
        expect(n.targetKind).toBe('privateChannel');
        expect(n.lehrerTeamId).toBe('tid');
        expect(n.items[0].channelId).toBe('c2');
        expect(n.items[0].chatId).toBe('');
    });
});
