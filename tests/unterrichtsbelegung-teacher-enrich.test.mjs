import { describe, it, expect } from 'vitest';
import {
    buildTeacherLookupFromList,
    enrichBelegungRow,
    enrichBelegungRows
} from '../src/shared/unterrichtsbelegung-teacher-enrich-logic.js';

describe('unterrichtsbelegung-teacher-enrich-logic', () => {
    const lookup = buildTeacherLookupFromList([
        { code: 'FRECH', email: 'frech@school.at', name: 'Maria Frech' }
    ]);

    it('ergänzt E-Mail und Name aus Kürzel', () => {
        const row = enrichBelegungRow({ lehrerCode: 'FRECH', fach: 'BW' }, lookup, {});
        expect(row.lehrerEmail).toBe('frech@school.at');
        expect(row.lehrerName).toBe('Maria Frech');
    });

    it('nutzt M365-Besitzer wenn E-Mail fehlt', () => {
        const row = enrichBelegungRow({ lehrerCode: 'FRECH' }, lookup, {
            ownerEmail: 'other@school.at'
        });
        expect(row.lehrerEmail).toBe('other@school.at');
        expect(row.lehrerName).toBe('Maria Frech');
    });

    it('ergänzt Kürzel aus Besitzer-E-Mail', () => {
        const row = enrichBelegungRow({}, lookup, { ownerEmail: 'frech@school.at' });
        expect(row.lehrerCode).toBe('FRECH');
        expect(row.lehrerName).toBe('Maria Frech');
    });

    it('mappt Besitzer pro Gruppe', () => {
        const rows = enrichBelegungRows(
            [{ graphGroupId: 'g1', lehrerCode: 'FRECH' }],
            lookup,
            new Map([['g1', 'frech@school.at']])
        );
        expect(rows[0].lehrerEmail).toBe('frech@school.at');
    });
});
