import { describe, expect, it } from 'vitest';
import {
    leadSeatLabelForNode,
    nodeUsesGraphFocusPeople,
    ensureLeadSeat,
    mergeGraphFocusPeople,
    graphUserToPerson
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
        expect(leadSeatLabelForNode({ type: 'class' })).toBe('Klassenvorstand');
        expect(leadSeatLabelForNode({ type: 'kursteam' })).toBe('Besitzer');
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
        expect(merged.find((p) => p.email === 'lea@schule.at')?.kind).toBe('student');
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
