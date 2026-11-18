import { describe, expect, it } from 'vitest';
import {
    buildCohortTeamDisplayName,
    buildCohortTeamMailNickname,
    buildDiplomPlan,
    buildPrivateChannelDisplayName,
    isDiplomGroup,
    labelsForArbeitTyp
} from '../src/tools/diplomarbeiten/diplomarbeiten-logic.js';

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
    });
});
