import { describe, expect, it } from 'vitest';
import {
    abschlussjahrFromMailNickname,
    deriveClassEntryFromGroup,
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
