import { describe, expect, it } from 'vitest';
import {
    diffTeachersImport,
    mergeTeachersImportLists,
    replaceTeachersImportList,
    summarizeTeachersDiff
} from '../src/shared/webuntis-stammdaten-wizard-logic.js';

describe('teacher-import-diff', () => {
    it('erkennt neue Kürzel und Aktualisierungen', () => {
        const existing = [{ code: 'MU', name: 'Max Alt', email: 'max@schule.at' }];
        const incoming = [
            { code: 'MU', name: 'Max Mustermann', email: 'max@schule.at' },
            { code: 'BME', name: 'Anna Beispiel', email: 'anna@schule.at' }
        ];
        const diff = diffTeachersImport(existing, incoming);
        expect(diff.counts.added).toBe(1);
        expect(diff.counts.updated).toBe(1);
        expect(diff.added[0].code).toBe('BME');
        expect(diff.updated[0].code).toBe('MU');
        expect(summarizeTeachersDiff(diff)).toContain('neu');
    });

    it('meldet E-Mail-Konflikt bei gleichem Kürzel', () => {
        const diff = diffTeachersImport(
            [{ code: 'XY', name: 'A', email: 'a@schule.at' }],
            [{ code: 'XY', name: 'A', email: 'b@schule.at' }]
        );
        expect(diff.counts.conflicts).toBe(1);
        expect(diff.conflicts[0].code).toBe('XY');
    });

    it('führt Listen nach Kürzel zusammen', () => {
        const merged = mergeTeachersImportLists(
            [{ code: 'OLD', name: 'Bleibt', email: '' }],
            [{ code: 'NEW', name: 'Neu', email: 'n@x.at' }, { code: 'OLD', name: 'Neu Name', email: 'o@x.at' }]
        );
        expect(merged.map((r) => r.code)).toEqual(['NEW', 'OLD']);
        expect(merged.find((r) => r.code === 'OLD').name).toBe('Neu Name');
    });

    it('ersetzt die Liste vollständig', () => {
        const only = replaceTeachersImportList([{ code: 'Z', name: 'Z', email: '' }]);
        expect(only).toHaveLength(1);
        expect(only[0].code).toBe('Z');
    });
});
