import { describe, it, expect } from 'vitest';
import { loadScript } from './kursteams-vm.mjs';

describe('kursteam-teams-filter-logic', () => {
    it('filtert nach Klasse, Fach, Lehrer, Status und Freitext', () => {
        const ctx = loadScript('src/tools/kursteams/kursteam-teams-filter-logic.js');
        const { filterTeamsWithIndices, collectUniqueTeamFilterValues } = ctx.ms365KursteamTeamsFilterLogic;

        const teams = [
            {
                teamName: 'SJ26 | 1AK | D',
                gruppenmail: 'SJ26-1AK-D',
                besitzer: 'mei@schule.at',
                isValid: true,
                originalClass: '1AK',
                fach: 'D',
                lehrerCode: 'MEI',
                gruppe: ''
            },
            {
                teamName: 'SJ26 | 2BK | M',
                gruppenmail: 'SJ26-2BK-M',
                besitzer: 'fis@schule.at',
                isValid: false,
                originalClass: '2BK',
                fach: 'M',
                lehrerCode: 'FIS',
                gruppe: '',
                error: 'Fehler'
            },
            {
                teamName: 'SJ26 | 1AK | M',
                gruppenmail: 'SJ26-1AK-M',
                besitzer: 'mei@schule.at',
                isValid: true,
                originalClass: '1AK',
                fach: 'M',
                lehrerCode: 'MEI',
                gruppe: ''
            }
        ];

        expect(filterTeamsWithIndices(teams, { klasse: '1ak' })).toHaveLength(2);
        expect(filterTeamsWithIndices(teams, { fach: 'M' })).toHaveLength(2);
        expect(filterTeamsWithIndices(teams, { lehrer: 'fis' })).toHaveLength(1);
        expect(filterTeamsWithIndices(teams, { status: 'invalid' })).toHaveLength(1);
        expect(filterTeamsWithIndices(teams, { q: 'mei@' })).toHaveLength(2);
        expect(collectUniqueTeamFilterValues(teams, 'klasse')).toEqual(['1AK', '2BK']);
        expect(collectUniqueTeamFilterValues(teams, 'lehrer')).toEqual(['FIS', 'MEI']);
    });
});
