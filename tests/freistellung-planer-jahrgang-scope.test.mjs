import { describe, it, expect } from 'vitest';
import {
    normalizeAllowedJahrgang,
    normalizeJahrgangGroups,
    schulstufeFromClassCode,
    jahrgangeFromEntraMembership,
    classCodesForJahrgange,
    buildJahrgangScope,
    itemMatchesJahrgangClassCodes
} from '../src/tools/freistellung-planer/freistellung-planer-jahrgang-scope.js';
import { filterFreistellungen } from '../src/tools/freistellung-planer/freistellung-planer-logic.js';

const G5 = 'aaaaaaaa-bbbb-cccc-dddd-000000000005';

describe('freistellung-planer-jahrgang-scope', () => {
    it('normalisiert Jahrgang und Gruppen', () => {
        expect(normalizeAllowedJahrgang('5, 6,5')).toEqual(['5', '6']);
        const g = normalizeJahrgangGroups([
            { jahrgang: '5', groupId: G5, groupLabel: 'Jg 5' },
            { jahrgang: 'x', groupId: 'bad' }
        ]);
        expect(g).toHaveLength(1);
        expect(g[0].groupLabel).toBe('Jg 5');
    });

    it('leitet Schulstufe aus Klassenkürzel ab', () => {
        expect(schulstufeFromClassCode('5A')).toBe('5');
        expect(schulstufeFromClassCode('10B')).toBe('10');
    });

    it('filtert Freistellungen nach Jahrgangs-Klassen', () => {
        const config = {
            allowedJahrgang: ['5'],
            jahrgangGroups: [{ jahrgang: '5', groupId: G5, groupLabel: 'Jg5' }]
        };
        const member = new Set([G5.toLowerCase()]);
        const scope = buildJahrgangScope(config, member, [
            { code: '5A' },
            { code: '6A' }
        ]);
        expect(scope.jahrgange).toEqual(['5']);
        expect(scope.classCodes.has('5A')).toBe(true);
        expect(scope.classCodes.has('6A')).toBe(false);

        const items = [
            { klasse: '5A', kvEmail: 'other@school.at' },
            { klasse: '6A', kvEmail: 'other@school.at' }
        ];
        const filtered = filterFreistellungen(items, {}, {
            accountEmail: 'coord@school.at',
            jahrgangClassCodes: scope.classCodes
        });
        expect(filtered).toHaveLength(1);
        expect(filtered[0].klasse).toBe('5A');
        expect(itemMatchesJahrgangClassCodes(scope.classCodes, '5A')).toBe(true);
        expect(jahrgangeFromEntraMembership(member, config)).toEqual(['5']);
        expect(classCodesForJahrgange(['5'], [{ code: '5C' }])).toEqual(new Set(['5C']));
    });
});
