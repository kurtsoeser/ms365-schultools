import { describe, expect, it } from 'vitest';
import {
    countOpenDeviations,
    extractFirstName,
    formatGreetingBandText,
    greetingForHour
} from '../src/shared/dashboard-greeting-band.js';

describe('dashboard-greeting-band', () => {
    it('grüßt nach Tageszeit', () => {
        expect(greetingForHour(8)).toBe('Guten Morgen');
        expect(greetingForHour(14)).toBe('Guten Tag');
        expect(greetingForHour(20)).toBe('Guten Abend');
    });

    it('formatiert die Statuszeile', () => {
        const line = formatGreetingBandText({
            greeting: 'Guten Morgen',
            firstName: 'Kurt',
            deviations: 2,
            schuljahrSteps: '4/8 Schritte',
            hasSchoolData: true
        });
        expect(line).toContain('Guten Morgen, Kurt');
        expect(line).toContain('2 offene Abweichungen');
        expect(line).toContain('Schuljahresstart: 4/8 Schritte');
    });

    it('zählt Abweichungen aus Tönen und Hygiene', () => {
        expect(countOpenDeviations(['warn', 'ok'], { mismatch: 1 })).toBe(2);
        expect(extractFirstName('Kurt Söser')).toBe('Kurt');
    });
});
