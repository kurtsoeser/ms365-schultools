import { describe, it, expect } from 'vitest';
import {
    mergeKategorieChoices,
    addExtraKategorie,
    isAllowedKategorie,
    normalizeExtraKategorien
} from '../src/tools/freistellung-planer/freistellung-planer-kategorien.js';
import { parseNachweiseField, stringifyNachweiseField } from '../src/tools/freistellung-planer/freistellung-planer-nachweise.js';

describe('freistellung-planer-kategorien', () => {
    it('merged Standard + Extra ohne Duplikate', () => {
        const all = mergeKategorieChoices(['Sport', 'Ärztlicher Termin']);
        expect(all).toContain('Sport');
        expect(all).toContain('Sonstiges');
        expect(all.filter((k) => k === 'Ärztlicher Termin')).toHaveLength(1);
    });

    it('isAllowedKategorie akzeptiert Extra', () => {
        const allowed = mergeKategorieChoices(['Schulausflug']);
        expect(isAllowedKategorie('Schulausflug', allowed)).toBe(true);
        expect(isAllowedKategorie('X', allowed)).toBe(false);
    });

    it('addExtraKategorie ignoriert Duplikate', () => {
        const list = normalizeExtraKategorien(['Ausflug', 'Beruf']);
        const next = addExtraKategorie('Ausflug', list);
        expect(next).toEqual(['Ausflug', 'Beruf']);
    });
});

describe('freistellung-planer-nachweise', () => {
    it('roundtrip JSON Feld', () => {
        const links = [{ name: 'attest.pdf', url: 'https://example.test/x' }];
        const raw = stringifyNachweiseField(links);
        expect(parseNachweiseField(raw)).toEqual(links);
    });
});
