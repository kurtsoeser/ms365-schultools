import { describe, expect, it } from 'vitest';
import {
    localPartFromEmailPatternId,
    serializeEmailBuilderPattern,
    defaultEmailBuilderPattern
} from '../src/shared/person-email-pattern-logic.js';

describe('person-email-pattern-logic', () => {
    it('Preset vorname.nachname', () => {
        expect(
            localPartFromEmailPatternId('vorname.nachname', {
                vorname: 'Max',
                nachname: 'Mustermann',
                firstNameMode: 'first'
            })
        ).toBe('max.mustermann');
    });

    it('Preset kuerzel.nachname', () => {
        expect(
            localPartFromEmailPatternId('kuerzel.nachname', {
                vorname: 'Anna',
                nachname: 'Beispiel',
                kuerzel: 'BME',
                firstNameMode: 'first'
            })
        ).toBe('bme.beispiel');
    });

    it('Builder: kuerzel + punkt + nachname', () => {
        const id = serializeEmailBuilderPattern([
            { type: 'kuerzel' },
            { type: 'sep', value: '.' },
            { type: 'nachname' }
        ]);
        expect(
            localPartFromEmailPatternId(id, {
                vorname: 'X',
                nachname: 'Muster',
                kuerzel: 'MU',
                firstNameMode: 'first'
            })
        ).toBe('mu.muster');
    });

    it('default builder entspricht vorname.nachname', () => {
        const id = serializeEmailBuilderPattern(defaultEmailBuilderPattern());
        expect(
            localPartFromEmailPatternId(id, {
                vorname: 'Lisa',
                nachname: 'Huber',
                firstNameMode: 'first'
            })
        ).toBe('lisa.huber');
    });
});
