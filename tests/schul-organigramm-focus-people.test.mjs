import { describe, expect, it } from 'vitest';
import {
    areaHasLinkedM365Group,
    leadSeatLabelForNode,
    memberSeatLabelForNode,
    nodeUsesGraphFocusPeople,
    ensureLeadSeat,
    mergeGraphFocusPeople,
    graphUserToPerson,
    normalizeGraphMemberList
} from '../src/tools/schul-organigramm/schul-organigramm-focus-people.js';
import { computeFocusLayout } from '../src/tools/schul-organigramm/schul-organigramm-canvas.js';
import { seedSchulOrganigrammFromStammdaten } from '../src/tools/schul-organigramm/schul-organigramm-logic.js';

describe('schul-organigramm-focus-people', () => {
    it('erkennt ARGE/Klasse/Unterrichtsteam für Graph-Fokus', () => {
        expect(nodeUsesGraphFocusPeople({ type: 'arge' })).toBe(true);
        expect(
            nodeUsesGraphFocusPeople({ type: 'area', refKind: 'arge', refKey: 'MATHE' })
        ).toBe(true);
        expect(nodeUsesGraphFocusPeople({ type: 'class' })).toBe(true);
        expect(nodeUsesGraphFocusPeople({ type: 'kursteam' })).toBe(true);
        expect(nodeUsesGraphFocusPeople({ type: 'area' })).toBe(false);
        expect(nodeUsesGraphFocusPeople({ type: 'role' })).toBe(false);
        expect(leadSeatLabelForNode({ type: 'arge' })).toBe('ARGE-Leitung');
        expect(
            leadSeatLabelForNode({ type: 'area', refKind: 'arge', refKey: 'MATHE' })
        ).toBe('ARGE-Leitung');
        expect(memberSeatLabelForNode({ type: 'arge' })).toBe('ARGE-Mitglieder');
        expect(
            memberSeatLabelForNode({ type: 'area', refKind: 'arge', refKey: 'MATHE' })
        ).toBe('ARGE-Mitglieder');
        expect(memberSeatLabelForNode({ type: 'class' })).toBe('Schüler');
        expect(
            memberSeatLabelForNode({ type: 'area', refKind: 'class', refKey: '4A' })
        ).toBe('Schüler');
        expect(leadSeatLabelForNode({ type: 'class' })).toBe('Klassenvorstand');
        expect(leadSeatLabelForNode({ type: 'kursteam' })).toBe('Besitzer');
    });

    it('Bereich mit verknüpfter M365-Gruppe nutzt Graph-Mitglieder', () => {
        const linked = {
            type: 'area',
            id: 'sec_custom',
            m365: { groupMode: 'match', groupId: 'gid-demo', groupKind: 'unified' }
        };
        expect(areaHasLinkedM365Group(linked)).toBe(true);
        expect(nodeUsesGraphFocusPeople(linked)).toBe(true);
        expect(
            areaHasLinkedM365Group({
                type: 'area',
                m365: { groupMode: 'none', groupId: 'gid' }
            })
        ).toBe(false);
        expect(
            nodeUsesGraphFocusPeople({
                type: 'area',
                m365: { groupMode: 'create', groupId: null }
            })
        ).toBe(false);
        expect(
            mergeGraphFocusPeople({
                node: linked,
                localPeople: [{ kind: 'areaPerson', name: 'Lokal', email: 'lokal@schule.at' }],
                members: [
                    { displayName: 'Anna', mail: 'anna@schule.at' },
                    { displayName: 'Lokal', mail: 'lokal@schule.at' }
                ]
            }).map((p) => p.email)
        ).toEqual(['anna@schule.at', 'lokal@schule.at']);
    });

    it('ARGE-Bereich: Leitung + Graph-Mitglieder ohne Doppelung', () => {
        const argeArea = {
            type: 'area',
            refKind: 'arge',
            refKey: 'DEMO-MAT',
            id: 'arge_DEMO-MAT',
            m365: { groupMode: 'match', groupId: 'gid-mat', groupKind: 'unified' }
        };
        expect(nodeUsesGraphFocusPeople(argeArea)).toBe(true);
        expect(leadSeatLabelForNode(argeArea)).toBe('ARGE-Leitung');
        const merged = mergeGraphFocusPeople({
            node: argeArea,
            localLead: { name: 'Müller', email: 'mueller@schule.at' },
            teachers: [{ name: 'Müller', email: 'mueller@schule.at' }],
            owners: [{ displayName: 'Müller', mail: 'mueller@schule.at' }],
            members: [
                { displayName: 'Müller', mail: 'mueller@schule.at' },
                { displayName: 'Lea', mail: 'lea@schule.at' }
            ]
        });
        expect(merged[0]).toMatchObject({
            kind: 'head',
            seatLabel: 'ARGE-Leitung',
            email: 'mueller@schule.at'
        });
        const members = merged.filter((p) => p.kind !== 'head');
        expect(members.map((p) => p.email)).toEqual(['lea@schule.at']);
        expect(members[0]).toMatchObject({
            email: 'lea@schule.at',
            seatLabel: 'ARGE-Mitglieder'
        });
    });

    it('ARGE-Bereich: weitere Besitzer als ARGE-Mitglieder (nicht nur Besitzer)', () => {
        const argeArea = {
            type: 'area',
            refKind: 'arge',
            refKey: 'MUSIK',
            id: 'arge_MUSIK'
        };
        const merged = mergeGraphFocusPeople({
            node: argeArea,
            localLead: { name: 'Leitung', email: 'lead@schule.at' },
            owners: [
                { displayName: 'Leitung', mail: 'lead@schule.at' },
                { displayName: 'Zweiter', mail: 'zweit@schule.at' }
            ],
            members: [{ displayName: 'Fach', mail: 'fach@schule.at' }]
        });
        expect(merged[0].seatLabel).toBe('ARGE-Leitung');
        const rest = merged.filter((p) => p.kind !== 'head');
        expect(rest.map((p) => ({ email: p.email, seat: p.seatLabel }))).toEqual([
            { email: 'fach@schule.at', seat: 'ARGE-Mitglieder' },
            { email: 'zweit@schule.at', seat: 'ARGE-Mitglieder' }
        ]);
    });

    it('mergeGraphFocusPeople: Unterrichtsteam Besitzer + Schüler', () => {
        const merged = mergeGraphFocusPeople({
            node: { type: 'kursteam', refKey: 'mat@1AK' },
            localLead: { name: 'Müller', email: 'mueller@schule.at' },
            teachers: [{ name: 'Müller', email: 'mueller@schule.at' }],
            owners: [
                { displayName: 'Müller', mail: 'mueller@schule.at' },
                { displayName: 'Ko-Lehrer', mail: 'ko@schule.at' }
            ],
            members: [
                { displayName: 'Müller', mail: 'mueller@schule.at' },
                { displayName: 'Lea', mail: 'lea@schule.at' },
                { displayName: 'Max', mail: 'max@schule.at' }
            ]
        });
        const owners = merged.filter((p) => p.kind === 'head');
        const students = merged.filter((p) => p.kind === 'student');
        expect(owners).toHaveLength(2);
        expect(owners.every((p) => p.seatLabel === 'Besitzer')).toBe(true);
        expect(students.map((p) => p.email).sort()).toEqual(['lea@schule.at', 'max@schule.at']);
    });

    it('ensureLeadSeat stellt offenen Leitungsslot voran', () => {
        const people = ensureLeadSeat(
            [{ kind: 'member', name: 'Anna', email: 'a@schule.at' }],
            { type: 'arge' },
            null
        );
        expect(people[0].kind).toBe('head');
        expect(people[0].vacant).toBe(true);
        expect(people[0].seatLabel).toBe('ARGE-Leitung');
        expect(people).toHaveLength(2);
    });

    it('normalizeGraphMemberList: Array und { items }', () => {
        expect(normalizeGraphMemberList([{ mail: 'a@x.at' }])).toHaveLength(1);
        expect(
            normalizeGraphMemberList({
                items: [{ mail: 'a@x.at' }, { mail: 'b@x.at' }],
                truncated: false
            })
        ).toHaveLength(2);
        expect(normalizeGraphMemberList(null)).toEqual([]);
    });

    it('mergeGraphFocusPeople: Besitzer → Leitung, Mitglieder ohne Doppelung', () => {
        const merged = mergeGraphFocusPeople({
            node: { type: 'arge', refKey: 'MATHE' },
            localPeople: [],
            localLead: null,
            teachers: [{ name: 'Müller', email: 'mueller@schule.at' }],
            owners: [{ displayName: 'Müller', mail: 'mueller@schule.at' }],
            members: [
                { displayName: 'Müller', mail: 'mueller@schule.at' },
                { displayName: 'Lea', mail: 'lea@schule.at' },
                { displayName: 'Max', userPrincipalName: 'max@schule.at' }
            ]
        });
        expect(merged[0].kind).toBe('head');
        expect(merged[0].email).toBe('mueller@schule.at');
        expect(merged[0].vacant).toBeFalsy();
        const members = merged.filter((p) => p.kind !== 'head');
        expect(members.map((p) => p.email).sort()).toEqual(['lea@schule.at', 'max@schule.at']);
        expect(members.every((p) => p.fromGraph)).toBe(true);
    });

    it('mergeGraphFocusPeople: { items } von fetchGroupMembers + zweiter Besitzer', () => {
        const merged = mergeGraphFocusPeople({
            node: { type: 'area', refKind: 'arge', refKey: 'MATHE' },
            localLead: null,
            teachers: [{ name: 'Müller', email: 'mueller@schule.at' }],
            owners: [
                { displayName: 'Müller', mail: 'mueller@schule.at' },
                { displayName: 'Schmidt', mail: 'schmidt@schule.at' }
            ],
            members: {
                items: [
                    { displayName: 'Müller', mail: 'mueller@schule.at' },
                    { displayName: 'Lea', mail: 'lea@schule.at' },
                    { displayName: 'Max', mail: 'max@schule.at' },
                    { displayName: 'Anna', mail: 'anna@schule.at' },
                    { displayName: 'Ben', mail: 'ben@schule.at' },
                    { displayName: 'Schmidt', mail: 'schmidt@schule.at' }
                ],
                truncated: false
            }
        });
        expect(merged[0]).toMatchObject({ kind: 'head', email: 'mueller@schule.at' });
        const rest = merged.filter((p) => p.kind !== 'head');
        expect(rest).toHaveLength(5);
        expect(rest.map((p) => p.email).sort()).toEqual([
            'anna@schule.at',
            'ben@schule.at',
            'lea@schule.at',
            'max@schule.at',
            'schmidt@schule.at'
        ]);
    });

    it('mergeGraphFocusPeople: Klasse nutzt Schüler-Kind und behält lokale Leitung', () => {
        const merged = mergeGraphFocusPeople({
            node: { type: 'class', refKey: '1A' },
            localLead: { name: 'Schmidt', email: 'schmidt@schule.at' },
            teachers: [{ name: 'Schmidt', email: 'schmidt@schule.at' }],
            owners: [{ displayName: 'Andere', mail: 'andere@schule.at' }],
            members: [
                { displayName: 'Lea', mail: 'lea@schule.at' },
                { displayName: 'Schmidt', mail: 'schmidt@schule.at' }
            ]
        });
        expect(merged[0].seatLabel).toBe('Klassenvorstand');
        expect(merged[0].email).toBe('schmidt@schule.at');
        expect(merged.filter((p) => p.kind === 'student')).toHaveLength(1);
        expect(merged.find((p) => p.email === 'lea@schule.at')).toMatchObject({
            kind: 'student',
            seatLabel: 'Schüler'
        });
    });

    it('graphUserToPerson mappt Graph-User', () => {
        const p = graphUserToPerson({ displayName: 'Tom', mail: 'tom@schule.at' }, 'member');
        expect(p).toMatchObject({ name: 'Tom', email: 'tom@schule.at', fromGraph: true });
    });

    it('computeFocusLayout: Leitung vor Mitgliedern', () => {
        const settings = {
            adminRoles: [],
            admin: [],
            teachers: [{ code: 'MUEL', name: 'Müller', email: 'mueller@schule.at' }],
            arges: [{ code: 'MATHE', name: 'Mathematik', headName: 'Müller', headEmail: 'mueller@schule.at' }],
            classes: [],
            students: [],
            studentCouncil: [],
            sga: []
        };
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const arge = layout.nodes.find((n) => n.refKind === 'arge' && n.refKey === 'MATHE');
        expect(arge).toBeTruthy();
        const people = mergeGraphFocusPeople({
            node: arge,
            localLead: { name: 'Müller', email: 'mueller@schule.at' },
            members: [
                { displayName: 'Anna', mail: 'anna@schule.at' },
                { displayName: 'Ben', mail: 'ben@schule.at' }
            ]
        });
        const focus = computeFocusLayout(layout, arge, people, []);
        const chips = focus.boxes.filter((b) => b.kind === 'person-chip');
        expect(chips[0].leadSeat).toBe(true);
        expect(chips[0].person.kind).toBe('head');
        expect(chips.filter((b) => !b.leadSeat)).toHaveLength(2);
    });
});
