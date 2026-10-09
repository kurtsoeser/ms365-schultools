import { describe, it, expect } from 'vitest';
import { buildLfrCalendarEvent } from '../src/tools/lehrer-freistellung-planer/lfr-calendar-sync.js';

describe('lfr-calendar-sync', () => {
    it('baut Graph-Event für genehmigten Antrag', () => {
        const ev = buildLfrCalendarEvent({
            titel: 'Fortbildung',
            lehrerName: 'Max Mustermann',
            lehrerEmail: 'max@schule.at',
            beginn: '2026-03-10T08:00:00',
            ende: '2026-03-10T16:00:00',
            kategorie: 'Fortbildung',
            antragId: 'lfr-1'
        });
        expect(ev.subject).toContain('Max Mustermann');
        expect(ev.start.dateTime).toContain('2026-03-10');
        expect(ev.categories).toContain('Lehrer-Freistellung');
    });
});
