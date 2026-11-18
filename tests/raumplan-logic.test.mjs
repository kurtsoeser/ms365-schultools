import { describe, expect, it } from 'vitest';
import { filterRaeume, technikMeldungUrl, validateRaum } from '../src/tools/raumplan/raumplan-logic.js';
import { DEMO_SEED_TAG, isDemoSeedItem, normalizeRaum } from '../src/tools/raumplan/raumplan-model.js';
import { mapRaumFromItem, mapRaumToFields } from '../src/tools/raumplan/raumplan-graph.js';

describe('raumplan validate + filter', () => {
    it('validateRaum verlangt Code und Name', () => {
        expect(validateRaum({ code: '', name: 'A' }).ok).toBe(false);
        expect(validateRaum({ code: 'INF1', name: 'Informatik' }).ok).toBe(true);
    });

    it('filterRaeume nach Typ und Suche', () => {
        const list = [
            normalizeRaum({ code: 'INF1', name: 'Informatik 1', typ: 'fachraum', gebaeude: 'B' }),
            normalizeRaum({ code: 'R7A', name: 'Klassenraum 7A', typ: 'klasse', gebaeude: 'A' })
        ];
        expect(filterRaeume(list, { typ: 'fachraum' }).map((r) => r.code)).toEqual(['INF1']);
        expect(filterRaeume(list, { q: '7a' }).map((r) => r.code)).toEqual(['R7A']);
    });

    it('Technik-Deep-Link behält ?raum=', () => {
        expect(technikMeldungUrl('INF1')).toBe('technik-meldungen.html?raum=INF1');
    });
});

describe('raumplan SPO mapping', () => {
    it('mapRaumToFields / mapRaumFromItem roundtrip', () => {
        const fields = mapRaumToFields({
            id: 'raum-1',
            code: 'INF1',
            name: 'Informatik 1',
            gebaeude: 'Trakt B',
            stockwerk: '1',
            kapazitaet: 28,
            typ: 'fachraum',
            notiz: 'PC-Raum'
        });
        expect(fields.Title).toBe('Informatik 1');
        expect(fields.RaumId).toBe('raum-1');
        expect(fields.Code).toBe('INF1');
        expect(fields.Kapazitaet).toBe(28);
        expect(fields.Typ).toBe('fachraum');

        const mapped = mapRaumFromItem({
            id: 77,
            fields: { ...fields }
        });
        expect(mapped.spoItemId).toBe('77');
        expect(mapped.id).toBe('raum-1');
        expect(mapped.code).toBe('INF1');
        expect(mapped.name).toBe('Informatik 1');
        expect(mapped.kapazitaet).toBe(28);
    });

    it('Demo-Seed wird erkannt und nicht als Produktiv behandelt', () => {
        const demo = normalizeRaum({
            code: 'DEMO',
            name: 'Demo',
            seedTag: DEMO_SEED_TAG
        });
        expect(isDemoSeedItem(demo)).toBe(true);
        expect(isDemoSeedItem(normalizeRaum({ code: 'X', name: 'Y' }))).toBe(false);
    });
});

describe('raumplan smoke rollen scope', () => {
    const raeume = [
        normalizeRaum({ code: 'A1', name: 'Raum A', typ: 'klasse' }),
        normalizeRaum({ code: 'B2', name: 'Raum B', typ: 'fachraum' })
    ];

    it('Lehrer-Suche sieht gefilterte Liste (Read)', () => {
        expect(filterRaeume(raeume, { q: 'fach' })).toHaveLength(1);
        expect(filterRaeume(raeume, { q: '' })).toHaveLength(2);
    });
});
