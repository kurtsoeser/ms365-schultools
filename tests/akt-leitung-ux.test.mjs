import { describe, expect, it } from 'vitest';
import {
    schoolYearLabel,
    schoolYearStartFromDate,
    isoInSchoolYear,
    computeAktivitaetStats
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-logic.js';
import {
    scheduleStatusOf,
    scheduleStatusLabel,
    navSectionsForRole
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-state.js';
import {
    sanitizeFachNotiz,
    parseVerlauf
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-detail-ui.js';

describe('Leitung UX – Schuljahr & Status', () => {
    it('schoolYearStartFromDate uses Sept–Aug', () => {
        expect(schoolYearStartFromDate('2026-11-18')).toBe(2026);
        expect(schoolYearStartFromDate('2027-02-01')).toBe(2026);
        expect(schoolYearStartFromDate('2026-08-31')).toBe(2025);
        expect(schoolYearLabel(2026)).toBe('2026/27');
    });

    it('isoInSchoolYear and stats filter by Schuljahr', () => {
        expect(isoInSchoolYear('2026-10-01', 2026)).toBe(true);
        expect(isoInSchoolYear('2027-01-15', 2026)).toBe(true);
        expect(isoInSchoolYear('2026-08-01', 2026)).toBe(false);
        const stats = computeAktivitaetStats(
            [
                { typ: 'Exkursion', status: 'genehmigt', startdatum: '2026-10-01', klasseCode: '2A' },
                { typ: 'Exkursion', status: 'beantragt', startdatum: '2027-03-01', klasseCode: '2A' },
                { typ: 'Veranstaltung', status: 'abgelehnt', startdatum: '2026-08-01', klasseCode: '3A' }
            ],
            { schoolYearStart: 2026 }
        );
        expect(stats.gesamt).toBe(2);
        expect(stats.schoolYearLabel).toBe('2026/27');
        expect(stats.byStatus.genehmigt).toBe(1);
        expect(stats.byStatus.beantragt).toBe(1);
    });

    it('schedule status stays separate from approval', () => {
        expect(scheduleStatusOf({ status: 'genehmigt', untisErledigt: false })).toBe('pending');
        expect(scheduleStatusOf({ status: 'genehmigt', untisErledigt: true })).toBe('completed');
        expect(scheduleStatusOf({ status: 'beantragt' })).toBe('not_required');
        expect(scheduleStatusLabel('pending')).toBe('Eintragung ausständig');
    });

    it('sanitizeFachNotiz strips demo seed tags', () => {
        expect(sanitizeFachNotiz('Bitte Bus · akt-demo-2026-27 · SJ 2026/27')).toBe('Bitte Bus');
        expect(sanitizeFachNotiz('akt-demo-gro-own')).toBe('');
    });

    it('parseVerlauf fasst Power-Automate-Kommentare als Block', () => {
        const entries = parseVerlauf(
            '---\nDatum: 2026-11-19T18:43:25Z\nFreu mich!\nKommentar:\nPasst, legendär!\n| Approval genehmigt\n.\nGenehmiger:\nDirektor'
        );
        expect(entries.length).toBe(1);
        expect(entries[0].when).toMatch(/19\.11\.2026/);
        expect(entries[0].text).toContain('Kommentar:');
        expect(entries[0].text).toContain('Direktor');
    });

    it('Admin-Nav enthält Systemadministration, nicht Neue Aktivität', () => {
        const ids = navSectionsForRole('admin').flatMap((s) => s.items.map((v) => v.id));
        expect(ids).toContain('system');
        expect(ids).not.toContain('antrag');
        expect(ids).toContain('freigabe');
    });
});
