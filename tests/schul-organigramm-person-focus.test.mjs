import { describe, expect, it } from 'vitest';
import { seedSchulOrganigrammFromStammdaten } from '../src/tools/schul-organigramm/schul-organigramm-logic.js';
import {
    schulOrgPersonFocusKey,
    schulOrgPersonFocusMatches,
    listSchulOrgDirectoryPeople,
    collectSchulOrgPersonMemberships
} from '../src/tools/schul-organigramm/schul-organigramm-person-focus.js';

describe('schul-organigramm person focus', () => {
    const settings = {
        adminRoles: [
            { code: 'PERSONAL', name: 'Personal', seatMode: 'single', tier: 'verwaltung' },
            { code: 'SEK', name: 'Sekretariat', seatMode: 'multi', tier: 'verwaltung' }
        ],
        admin: [
            { role: 'Personal', name: 'Eva Serafini', email: 'eva@schule.at' },
            { role: 'Sekretariat', name: 'Ben', email: 'ben@schule.at' }
        ],
        teachers: [{ code: 'MUEL', name: 'Müller', email: 'mueller@schule.at' }],
        arges: [
            {
                code: 'MATHE',
                name: 'Mathematik',
                subjects: ['MATHE'],
                headName: 'Müller',
                headEmail: 'mueller@schule.at'
            }
        ],
        classes: [{ code: '1A', name: '1A', headName: 'Müller', headEmail: 'mueller@schule.at' }],
        students: [{ klasse: '1A', name: 'Lea', email: 'lea@schule.at' }],
        studentCouncil: [],
        sga: [],
        adminAudienceMemberships: [
            { groupId: 'verwaltung', email: 'eva@schule.at', name: 'Eva Serafini' }
        ]
    };

    it('schulOrgPersonFocusKey / Matches', () => {
        expect(schulOrgPersonFocusKey({ email: 'Eva@Schule.at', name: 'X' })).toBe('e:eva@schule.at');
        expect(
            schulOrgPersonFocusMatches(
                { email: 'eva@schule.at' },
                { email: 'EVA@schule.at', name: 'Eva' }
            )
        ).toBe(true);
        expect(
            schulOrgPersonFocusMatches({ name: 'Nur Name' }, { name: 'nur name' })
        ).toBe(true);
    });

    it('listSchulOrgDirectoryPeople dedupliziert', () => {
        const people = listSchulOrgDirectoryPeople(settings);
        const eva = people.filter((p) => p.email === 'eva@schule.at');
        expect(eva).toHaveLength(1);
        expect(eva[0].sources.length).toBeGreaterThan(0);
        expect(people.some((p) => p.email === 'mueller@schule.at')).toBe(true);
    });

    it('collectSchulOrgPersonMemberships findet Rolle, Bereich und Leitung', () => {
        const layout = seedSchulOrganigrammFromStammdaten(settings);
        const eva = collectSchulOrgPersonMemberships(settings, layout, {
            email: 'eva@schule.at',
            name: 'Eva Serafini'
        });
        expect(eva.memberships.length).toBeGreaterThanOrEqual(1);
        expect(eva.memberships.some((m) => /personal/i.test(m.nodeLabel))).toBe(true);
        expect(eva.hitIds.length).toBeGreaterThan(0);
        expect(eva.pathIds.length).toBeGreaterThanOrEqual(eva.hitIds.length);

        const mueller = collectSchulOrgPersonMemberships(settings, layout, {
            email: 'mueller@schule.at',
            name: 'Müller'
        });
        const kinds = mueller.memberships.map((m) => m.kindLabel);
        expect(kinds.some((k) => /ARGE|Klassenvorstand|Lehrkraft/i.test(k))).toBe(true);
        // Klassen/ARGEs sind Bereichskacheln (type area, refKind class|arge); Legacy: type class|arge
        expect(
            mueller.memberships.some(
                (m) =>
                    m.nodeType === 'arge' ||
                    m.nodeType === 'class' ||
                    (m.nodeType === 'area' &&
                        (m.nodeId === 'arge_MATHE' ||
                            m.nodeId === 'class_1A' ||
                            /ARGE-Leitung|Klassenvorstand/i.test(m.kindLabel)))
            )
        ).toBe(true);
    });
});
