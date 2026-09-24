import { describe, it, expect } from 'vitest';
import {
    defaultChatNamePattern,
    normalizeChatNamePattern,
    buildChatTopicFromPattern,
    buildKvByKlasse,
    buildClassChatPlans,
    buildManualClassChatPlan,
    parseMemberEmailsText,
    summarizePlans,
    upsertClassChatItem,
    existingByKlasseFromState
} from '../src/tools/klassenchats/klassenchats-logic.js';

describe('klassenchats-logic', () => {
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
        expect(p1.topic).toBe('SJ26 | 1A | Lehrer');
        // 2B: nur C (+ KV ist dieselbe Mail) → 1 Mitglied → skip
        expect(p2.eligible).toBe(false);
        expect(p2.memberEmails).toHaveLength(1);
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
        expect(plan.source).toBe('manual');
    });
});
