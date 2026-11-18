import { describe, expect, it } from 'vitest';
import {
    allowedAntwortStatusesForModus,
    parseEmpfaengerText,
    progressStats,
    resolveAntwortStatus,
    validateEinsammler
} from '../src/tools/einsammler/einsammler-logic.js';
import { ANTWORT_STATUS, MODUS } from '../src/tools/einsammler/einsammler-model.js';
import { LIST_TITLES, KOEPFE_COLUMNS, ANTWORTEN_COLUMNS } from '../src/tools/einsammler/einsammler-schema.js';

describe('einsammler-logic', () => {
    it('validiert Titel', () => {
        expect(validateEinsammler({ titel: '' }).ok).toBe(false);
        expect(validateEinsammler({ titel: 'Foto', typ: 'foto' }).ok).toBe(true);
    });

    it('parst Empfängerzeilen', () => {
        const rows = parseEmpfaengerText('Anna Berger;3AHW;a@x.at\nBen Hofer\nClara;3BHW', '1A');
        expect(rows).toHaveLength(3);
        expect(rows[0]).toEqual({
            schuelerName: 'Anna Berger',
            klasseCode: '3AHW',
            elternEmail: 'a@x.at'
        });
        expect(rows[1].klasseCode).toBe('1A');
        expect(rows[2].klasseCode).toBe('3BHW');
    });

    it('berechnet Fortschritt', () => {
        const head = { id: 'ein-1' };
        const antworten = [
            { einsammlerId: 'ein-1', status: 'offen' },
            { einsammlerId: 'ein-1', status: 'zugestimmt' },
            { einsammlerId: 'ein-1', status: 'abgelehnt' },
            { einsammlerId: 'ein-2', status: 'offen' }
        ];
        const st = progressStats(head, antworten);
        expect(st.total).toBe(3);
        expect(st.offen).toBe(1);
        expect(st.erledigt).toBe(2);
        expect(st.pct).toBe(67);
    });

    it('markiert abgelaufene offene Antworten', () => {
        const head = { faelligAm: '2020-01-01' };
        expect(resolveAntwortStatus(head, { status: ANTWORT_STATUS.OFFEN }, '2026-01-01')).toBe(
            ANTWORT_STATUS.ABGELAUFEN
        );
        expect(resolveAntwortStatus(head, { status: ANTWORT_STATUS.ZUGESTIMMT }, '2026-01-01')).toBe(
            ANTWORT_STATUS.ZUGESTIMMT
        );
    });

    it('liefert erlaubte Status je Modus', () => {
        expect(allowedAntwortStatusesForModus(MODUS.LESEBESTAETIGUNG)).toContain(ANTWORT_STATUS.GELESEN);
        expect(allowedAntwortStatusesForModus(MODUS.EINWILLIGUNG)).toContain(ANTWORT_STATUS.ZUGESTIMMT);
        expect(allowedAntwortStatusesForModus(MODUS.EINWILLIGUNG)).not.toContain(ANTWORT_STATUS.GELESEN);
        expect(allowedAntwortStatusesForModus(MODUS.ABGABE)).toContain(ANTWORT_STATUS.ABGEGEBEN);
    });

    it('validiert Abgabe-Modus', () => {
        expect(validateEinsammler({ titel: 'Foto hochladen', modus: MODUS.ABGABE, typ: 'foto' }).ok).toBe(
            true
        );
    });
});

describe('einsammler-schema', () => {
    it('hat Listen und Spalten inkl. Anhang', () => {
        expect(LIST_TITLES.koepfe).toBe('Einsammler');
        expect(LIST_TITLES.antworten).toBe('EinsammlerAntworten');
        expect(KOEPFE_COLUMNS.some((c) => c.name === 'EinsammlerId')).toBe(true);
        expect(ANTWORTEN_COLUMNS.some((c) => c.name === 'AntwortId')).toBe(true);
        expect(ANTWORTEN_COLUMNS.some((c) => c.name === 'AnhangName')).toBe(true);
        expect(KOEPFE_COLUMNS.some((c) => c.name === 'UploadErlaubt')).toBe(true);
    });
});
