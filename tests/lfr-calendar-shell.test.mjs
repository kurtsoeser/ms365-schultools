import { describe, it, expect } from 'vitest';
import {
    buildLfrCalendarEvents,
    lfrItemToCalEvent
} from '../src/tools/lehrer-freistellung-planer/lfr-calendar.js';

describe('lfr calendar shell events', () => {
    it('baut Event mit data-lfr-detail', () => {
        const ev = lfrItemToCalEvent({
            antragId: 'LFR-1',
            itemId: 'spo-1',
            beginn: '2026-10-16T08:00',
            ende: '2026-10-16T16:00',
            status: 'Genehmigt',
            lehrerName: 'Kurt',
            titel: 'Fortbildung'
        });
        expect(ev).toBeTruthy();
        expect(ev.tone).toBe('genehmigt');
        expect(ev.openAttr).toContain('data-lfr-detail="LFR-1"');
        expect(ev.allDay).toBe(false);
    });

    it('blendet Abgelehnt aus', () => {
        const events = buildLfrCalendarEvents([
            {
                antragId: 'a',
                beginn: '2026-10-16',
                ende: '2026-10-16',
                status: 'Abgelehnt',
                lehrerName: 'A'
            },
            {
                antragId: 'b',
                beginn: '2026-10-17',
                ende: '2026-10-17',
                status: 'Ausstehend',
                lehrerName: 'B'
            }
        ]);
        expect(events).toHaveLength(1);
        expect(events[0].id).toBe('b');
        expect(events[0].tone).toBe('beantragt');
    });
});
