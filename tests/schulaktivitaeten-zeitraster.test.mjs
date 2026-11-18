import { describe, it, expect } from 'vitest';
import {
    timesFromUnterrichtseinheiten,
    applyUeToDraft,
    formatUeZeitraum,
    parseUeNumber
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-zeitraster.js';
import { computeAktivitaetStats } from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-logic.js';
import { buildIcs } from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-export.js';

describe('schulaktivitaeten zeitraster', () => {
    it('parses UE and derives clock times', () => {
        expect(parseUeNumber('3. UE')).toBe(3);
        const t = timesFromUnterrichtseinheiten({ startUe: 1, endUe: 6 });
        expect(t.ok).toBe(true);
        expect(t.startZeit).toBe('07:45');
        expect(t.endZeit).toBe('13:30');
    });

    it('rejects inverted UE range', () => {
        const t = timesFromUnterrichtseinheiten({ startUe: 5, endUe: 2 });
        expect(t.ok).toBe(false);
    });

    it('applyUeToDraft fills times', () => {
        const { draft } = applyUeToDraft({ startUe: 2, endUe: 2, titel: 'x' });
        expect(draft.startZeit).toBe('08:40');
        expect(draft.endZeit).toBe('09:30');
        expect(formatUeZeitraum(draft)).toContain('2. UE');
    });

    it('computeAktivitaetStats groups by year', () => {
        const stats = computeAktivitaetStats(
            [
                {
                    typ: 'Exkursion',
                    status: 'genehmigt',
                    startdatum: '2026-03-01',
                    klasseCode: '1AK',
                    untisErledigt: true
                },
                { typ: 'Exkursion', status: 'beantragt', startdatum: '2027-01-10', klasseCode: '2AK' },
                {
                    typ: 'Veranstaltung',
                    status: 'genehmigt',
                    startdatum: '2026-11-01',
                    klasseCodes: '3AK; 3BK',
                    nichtTeilnehmende: 'x',
                    untisErledigt: false
                }
            ],
            { year: 2026 }
        );
        expect(stats.gesamt).toBe(2);
        expect(stats.byTyp.Exkursion).toBe(1);
        expect(stats.byKlasse['3AK']).toBe(1);
        expect(stats.mitNichtteilnehmern).toBe(1);
        expect(stats.untisOffen).toBe(1);
    });

    it('ICS uses floating local times for same-day UE events', () => {
        const ics = buildIcs(
            [
                {
                    aktivitaetId: 'akt-1',
                    titel: 'Museum',
                    typ: 'Exkursion',
                    klasseCode: '1AK',
                    startdatum: '2026-11-18',
                    enddatum: '2026-11-18',
                    startUe: 1,
                    endUe: 4,
                    status: 'genehmigt'
                }
            ],
            { classes: [{ code: '1AK', name: '1AK' }], teachers: [] }
        );
        expect(ics).toContain('DTSTART:20261118T074500');
        expect(ics).toContain('DTEND:20261118T113000');
    });
});
