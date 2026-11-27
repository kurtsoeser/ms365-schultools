import { describe, expect, it } from 'vitest';
import {
    buildCohortTeamDisplayName,
    buildCohortTeamMailNickname,
    buildDiplomPlan,
    buildPrivateChannelDisplayName,
    cohortListStatus,
    diplomNameChoices,
    diplomSetupChoices,
    filterDiplomGroups,
    isDiplomGroup,
    labelsForArbeitTyp,
    maturajahrgangChoices,
    schoolYearStart
} from '../src/tools/diplomarbeiten/diplomarbeiten-logic.js';
import { embedSettingsInSchoolAppsPart } from '../src/tools/diplomarbeiten/diplomarbeiten-settings.js';

describe('diplomarbeiten-logic', () => {
    it('liefert Labels für Diplom und VWA', () => {
        expect(labelsForArbeitTyp('diplom').short).toBe('DA');
        expect(labelsForArbeitTyp('vwa').nickPrefix).toBe('vwa');
        expect(labelsForArbeitTyp('vwa').toolTitle).toBe('VWA');
    });

    it('baut Dach-Team- und Kanalnamen', () => {
        expect(
            buildCohortTeamDisplayName({ maturajahrgang: '2026/27', arbeitTyp: 'diplom' })
        ).toBe('Diplomarbeiten 2026/27');
        expect(
            buildCohortTeamMailNickname({ maturajahrgang: '2026/27', arbeitTyp: 'vwa' })
        ).toMatch(/^vwa-jahrgang-/);
        expect(
            buildPrivateChannelDisplayName({
                student: 'Müller',
                topic: 'KI',
                arbeitTyp: 'diplom'
            })
        ).toContain('DA');
    });

    it('Legacy-Plan und Erkennung', () => {
        const plan = buildDiplomPlan({
            year: '2027',
            topic: 'Nachhaltigkeit',
            student: 'Huber',
            arbeitTyp: 'vwa'
        });
        expect(plan.ok).toBe(true);
        expect(plan.displayName).toMatch(/Vorwissenschaftliche Arbeit/);
        expect(plan.mailNickname).toMatch(/^vwa-2027-/);

        expect(isDiplomGroup({ mailNickname: 'dipl-2026-x', displayName: 'x' })).toBe(true);
        expect(isDiplomGroup({ mailNickname: 'vwa-2026-y', displayName: 'x' })).toBe(true);
        expect(isDiplomGroup({ mailNickname: 'andere', displayName: 'Fachgruppe' })).toBe(false);
        expect(
            filterDiplomGroups(
                [
                    { displayName: 'Diplomarbeit 2027 – KI', mailNickname: 'dipl-2027-ki' },
                    { displayName: 'Vorwissenschaftliche Arbeit 2026', mailNickname: 'vwa-2026-x' },
                    { displayName: 'Fachgruppe', mailNickname: 'andere' }
                ],
                'ki'
            ).map((g) => g.mailNickname)
        ).toEqual(['dipl-2027-ki']);
    });

    it('erster Schritt nennt das Ergebnis, nicht den Betriebsmodus', () => {
        const diplom = diplomSetupChoices('diplom');
        const vwa = diplomSetupChoices('vwa');
        expect(diplom.map((c) => c.title)).toEqual(['Ein Team pro Jahrgang', 'Ein Team pro Arbeit']);
        expect(diplom.map((c) => c.modus)).toEqual(['channels', 'legacy']);
        expect(diplom[0].example).toContain('Diplomarbeiten 2026/27');
        expect(vwa[1].example).toContain('Vorwissenschaftliche Arbeit');
        expect(diplomNameChoices().map((n) => n.arbeitTyp)).toEqual(['diplom', 'vwa']);
    });

    it('bietet Maturajahre als Auswahl und nennt den Stand der Zeile', () => {
        const november = new Date(2026, 10, 25);
        expect(schoolYearStart(november)).toBe(2026);
        expect(schoolYearStart(new Date(2026, 7, 31))).toBe(2025);
        const choices = maturajahrgangChoices(november, ['2026/27']);
        expect(choices.map((c) => c.value)).toEqual([
            '2025/26',
            '2027/28',
            '2028/29',
            '2029/30',
            '2030/31'
        ]);
        expect(choices.some((c) => c.current)).toBe(false);
        expect(maturajahrgangChoices(november, [])[1]).toMatchObject({
            value: '2026/27',
            current: true
        });
        expect(cohortListStatus({ id: 'j1', teamId: '' }, [])).toEqual({
            key: 'jahrgang',
            label: 'Jahrgang angelegt'
        });
        expect(cohortListStatus({ id: 'j1', teamId: 't1' }, [{ cohortId: 'j1' }])).toEqual({
            key: 'team',
            label: 'Team verbunden'
        });
        expect(
            cohortListStatus({ id: 'j1', teamId: 't1' }, [
                { cohortId: 'j1', channelId: 'c1' },
                { cohortId: 'j1', channelId: 'c2' },
                { cohortId: 'andere', channelId: 'c3' }
            ])
        ).toEqual({ key: 'kanaele', label: '2 Kanäle' });
    });

    it('Schul-Datei behält fremde Schlüssel und setzt nur die Diplom-Wahl', () => {
        const next = embedSettingsInSchoolAppsPart(
            {
                localStorage: {
                    'ms365-elternbrief-settings-v1': '{"a":1}',
                    'ms365-diplomarbeiten-settings-v1': '{"alt":true}'
                },
                schoolName: 'Beispielschule'
            },
            '{"arbeitTyp":"vwa","modus":"channels","setupDone":true}'
        );
        expect(next.bundlePart).toBe('school-apps-settings');
        expect(next.localStorage['ms365-elternbrief-settings-v1']).toBe('{"a":1}');
        expect(next.localStorage['ms365-diplomarbeiten-settings-v1']).toContain('"vwa"');
        expect(next.schoolName).toBe('Beispielschule');
        expect(next.keyCount).toBe(2);
    });
});
