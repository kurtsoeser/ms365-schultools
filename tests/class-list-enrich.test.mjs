import { describe, expect, it } from 'vitest';
import {
    abschlussjahrFromMailNickname,
    deriveClassEntryFromGroup,
    classCodeFromDisplayName,
    pickHeadFromGraphGroupOwners,
    enrichClassesFromLinkedGroups,
    collectPriorClassFieldsByCode
} from '../src/shared/class-list-enrich.js';

describe('class-list-enrich', () => {
    it('abschlussjahrFromMailNickname liest jg-Schema', () => {
        expect(abschlussjahrFromMailNickname('jg2031hma')).toBe('2031');
        expect(abschlussjahrFromMailNickname('JG-2030-1AK')).toBe('2030');
        expect(abschlussjahrFromMailNickname('klasse1a')).toBe('');
    });

    it('deriveClassEntryFromGroup liest jg-Alias und Anzeigenamen', () => {
        expect(
            deriveClassEntryFromGroup({ mailNickname: 'jg2030-1ak', displayName: '1A-Klasse' })
        ).toEqual({ code: '1AK', name: '1A-Klasse', year: '2030' });
        expect(deriveClassEntryFromGroup({ mailNickname: 'klasse-5hma', displayName: '5 HMA' })).toMatchObject({
            code: '5HMA',
            year: ''
        });
        expect(deriveClassEntryFromGroup({ displayName: '5 HMA' })).toMatchObject({ code: '5HMA', name: '5 HMA' });
    });

    it('pickHeadFromGraphGroupOwners nutzt Stammdaten-Lehrer', () => {
        const head = pickHeadFromGraphGroupOwners(
            [
                { displayName: 'Max Mustermann', mail: 'max@schule.at' },
                { displayName: 'Andere', userPrincipalName: 'andere@schule.at' }
            ],
            [{ name: 'Mustermann Max', email: 'max@schule.at' }]
        );
        expect(head).toEqual({ headName: 'Mustermann Max', headEmail: 'max@schule.at' });
    });

    it('deriveClassEntryFromGroup: jg-Alias mit Einbuchstaben + DEMO Klasse 3A', () => {
        expect(classCodeFromDisplayName('DEMO Klasse 3A')).toBe('3A');
        expect(
            deriveClassEntryFromGroup({
                mailNickname: 'jg2030-a',
                displayName: 'DEMO Klasse 3A'
            })
        ).toEqual({ code: '3A', name: 'DEMO Klasse 3A', year: '2030' });
        expect(
            deriveClassEntryFromGroup({
                mailNickname: 'jg2031-b',
                displayName: 'DEMO Klasse 2B'
            })
        ).toMatchObject({ code: '2B', year: '2031' });
    });

    it('enrichClassesFromLinkedGroups füllt Jahr aus classTeams', () => {
        const classes = [{ code: 'HMA', name: '5HMA', year: '', headName: '', headEmail: '' }];
        const result = enrichClassesFromLinkedGroups(classes, {
            classTeams: [
                {
                    classCode: 'HMA',
                    abschlussJahr: '2031',
                    stableMailNickname: 'jg2031hma',
                    graphGroupId: 'g-1'
                }
            ],
            classGroupMatchByKey: {},
            priorByCode: new Map()
        });
        expect(result.changed).toBe(true);
        expect(result.classes[0].year).toBe('2031');
    });

    it('enrichClassesFromLinkedGroups übernimmt KV aus anderem Schuljahr', () => {
        const prior = collectPriorClassFieldsByCode({
            '2024/25': {
                classes: [{ code: '1AK', year: '2030', headName: 'Max Lehrer', headEmail: 'max@schule.at' }]
            }
        });
        const classes = [{ code: '1AK', name: '2AK', year: '', headName: '', headEmail: '' }];
        const result = enrichClassesFromLinkedGroups(classes, {
            classTeams: [],
            classGroupMatchByKey: {},
            priorByCode: prior
        });
        expect(result.classes[0].year).toBe('2030');
        expect(result.classes[0].headEmail).toBe('max@schule.at');
        expect(result.classes[0].headName).toBe('Max Lehrer');
    });
});
