import { describe, it, expect } from 'vitest';
import {
    mapSchularbeitFromItem,
    mapSchularbeitToFields,
    mapFensterFromItem,
    mapFensterToFields,
    mapRegelwerkFromItem,
    mapRegelwerkToFields,
    isSaAccessDeniedError
} from '../src/tools/schularbeiten-planer/schularbeiten-planer-graph.js';

describe('schularbeiten-planer-graph mapping', () => {
    it('mapSchularbeitFromItem liest Kernfelder inkl. itemId', () => {
        const row = mapSchularbeitFromItem({
            id: '99',
            fields: {
                Title: 'Schularbeit - MAM - 3AK',
                Titel: 'Schularbeit Mathematik',
                Thema: 'Gleichungen',
                SchularbeitId: 'sa-abc',
                FachCode: 'MAM',
                KlasseCode: '3AK',
                LehrerCode: 'MU',
                LehrerEmail: 'mueller@schule.at',
                Datum: '2026-11-20',
                BeginnUhrzeit: '08:00',
                DauerMinuten: 100,
                Semester: 'WS',
                Status: 'beantragt',
                BeantragtVon: 'mueller@schule.at',
                Schuljahr: '2026/27'
            }
        });
        expect(row.itemId).toBe('99');
        expect(row.schularbeitId).toBe('sa-abc');
        expect(row.fachCode).toBe('MAM');
        expect(row.klasseCode).toBe('3AK');
        expect(row.lehrerEmail).toBe('mueller@schule.at');
        expect(row.datum).toBe('2026-11-20');
        expect(row.status).toBe('beantragt');
        expect(row.schuljahr).toBe('2026/27');
    });

    it('mapSchularbeitToFields setzt Ownership und Labels im Title', () => {
        const fields = mapSchularbeitToFields(
            {
                schularbeitId: 'sa-1',
                titel: 'Probe',
                thema: 'Thema A',
                fachCode: 'MAM',
                klasseCode: '3AK',
                lehrerCode: 'MU',
                lehrerEmail: 'mueller@schule.at',
                datum: '2026-11-20',
                beginnUhrzeit: '08:00',
                dauerMinuten: 100,
                semester: 'ws',
                status: 'beantragt',
                beantragtVon: 'mueller@schule.at',
                schuljahr: '2026/27'
            },
            { labels: { fach: { MAM: 'Mathematik' }, klasse: { '3AK': '3AK' } } }
        );
        expect(fields.SchularbeitId).toBe('sa-1');
        expect(fields.FachCode).toBe('MAM');
        expect(fields.LehrerEmail).toBe('mueller@schule.at');
        expect(fields.BeantragtVon).toBe('mueller@schule.at');
        expect(fields.Status).toBe('beantragt');
        expect(fields.Semester).toBe('WS');
        expect(fields.Schuljahr).toBe('2026/27');
        expect(fields.Title).toMatch(/Mathematik/);
        expect(fields.Title).toMatch(/3AK/);
        expect(fields.FixiertAm).toBeUndefined();
    });

    it('mapFenster und mapRegelwerk roundtrip Kernfelder', () => {
        const win = mapFensterFromItem({
            id: '7',
            fields: {
                Title: 'Weihnachten',
                TerminfensterId: 'tf-1',
                Typ: 'gesperrt',
                Startdatum: '2026-12-24',
                Enddatum: '2027-01-06',
                Schuljahr: '2026/27'
            }
        });
        expect(win.itemId).toBe('7');
        expect(win.typ).toBe('gesperrt');
        const winFields = mapFensterToFields(win);
        expect(winFields.Typ).toBe('gesperrt');
        expect(winFields.TerminfensterId).toBe('tf-1');

        const rw = mapRegelwerkFromItem({
            id: '3',
            fields: {
                Title: 'Standard',
                RegelwerkId: 'rw-1',
                MaxProTag: 1,
                MaxProWoche: 2,
                AnkuendigungsfristTage: 7,
                SperreVorNotenkonferenzTage: 7,
                Aktiv: true,
                Schuljahr: '2026/27'
            }
        });
        expect(rw.itemId).toBe('3');
        expect(rw.maxProTag).toBe(1);
        const rwFields = mapRegelwerkToFields(rw);
        expect(rwFields.RegelwerkId).toBe('rw-1');
        expect(rwFields.Aktiv).toBe(true);
    });

    it('isSaAccessDeniedError erkennt 403/Forbidden', () => {
        expect(isSaAccessDeniedError(new Error('403 Forbidden'))).toBe(true);
        expect(isSaAccessDeniedError(new Error('Zugriff verweigert'))).toBe(true);
        expect(isSaAccessDeniedError(new Error('Liste fehlt'))).toBe(false);
    });
});
