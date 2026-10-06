import { describe, expect, it } from 'vitest';
import {
    applyRecommendedSubjectImport,
    buildSubjectPdfReviewPreview,
    subjectImportStats
} from '../src/shared/webuntis-stammdaten-wizard-logic.js';

describe('buildSubjectPdfReviewPreview', () => {
    it('dedupliziert Kürzel und setzt bei vielen Fächern empfohlene Auswahl', () => {
        const rows = [];
        for (let i = 1; i <= 30; i++) {
            rows.push({ code: 'F' + i, name: 'Fach ' + i });
        }
        rows.push({ code: 'ADM', name: 'Verwaltung', admin: true });
        const preview = buildSubjectPdfReviewPreview(rows, { sourceFileName: 'Subject_test.pdf' });
        expect(preview.subjects.length).toBe(31);
        const stats = subjectImportStats(preview.subjects);
        expect(stats.selected).toBeLessThan(stats.total);
        expect(preview.subjectImportHint).toMatch(/Empfohlene Auswahl/);
    });

    it('empfohlene Auswahl schließt Verwaltungsfächer aus', () => {
        const preview = buildSubjectPdfReviewPreview(
            [
                { code: 'D', name: 'Deutsch' },
                { code: 'ADM', name: 'Admin', admin: true }
            ],
            { deselectThreshold: 0 }
        );
        applyRecommendedSubjectImport(preview.subjects);
        const d = preview.subjects.find((r) => r.code === 'D');
        const adm = preview.subjects.find((r) => r.code === 'ADM');
        expect(d.selected).toBe(true);
        expect(adm.selected).toBe(false);
    });
});
