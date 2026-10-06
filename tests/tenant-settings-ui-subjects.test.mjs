import { describe, expect, it } from 'vitest';
import {
    SUBJECT_LIST_TEMPLATE_HEADERS,
    subjectListTemplateAoa,
    subjectsFromSpreadsheetJsonRows,
    subjectsToLines
} from '../src/shared/tenant-settings-ui-subjects.js';

describe('tenant-settings-ui-subjects (Fächer Excel-Vorlage)', () => {
    it('Vorlage hat Kürzel und Name als erste Zeile', () => {
        const aoa = subjectListTemplateAoa();
        expect(aoa[0]).toEqual(SUBJECT_LIST_TEMPLATE_HEADERS);
        expect(aoa.length).toBeGreaterThanOrEqual(2);
    });

    it('Beispielzeilen der Vorlage importieren korrekt', () => {
        const aoa = subjectListTemplateAoa();
        const headers = aoa[0];
        const jsonRows = aoa.slice(1).map((line) => {
            const o = {};
            headers.forEach((h, i) => {
                o[h] = line[i];
            });
            return o;
        });
        const rows = subjectsFromSpreadsheetJsonRows(jsonRows);
        expect(rows).toEqual([
            { code: 'D', name: 'Deutsch' },
            { code: 'M', name: 'Mathematik' },
            { code: 'E', name: 'Englisch' }
        ]);
        expect(subjectsToLines(rows)).toBe('D;Deutsch\nM;Mathematik\nE;Englisch');
    });

    it('erkennt alternative Spaltenüberschriften wie im Import', () => {
        const rows = subjectsFromSpreadsheetJsonRows([
            { Fach: 'bio', Bezeichnung: 'Biologie' },
            { Code: 'CH', Name: 'Chemie' }
        ]);
        expect(rows).toEqual([
            { code: 'BIO', name: 'Biologie' },
            { code: 'CH', name: 'Chemie' }
        ]);
    });
});
