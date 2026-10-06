import { describe, expect, it } from 'vitest';
import { navBadgeFromProgress, navStatusDotTone } from '../src/shared/dashboard-tasks-split.js';

describe('dashboard-tasks-split', () => {
    it('mappt Fortschritt auf Nav-Badges', () => {
        expect(navBadgeFromProgress('ok', '')).toEqual({ text: 'Konsistent', tone: 'ok' });
        expect(navBadgeFromProgress('warn', '2/4 Sammelgruppen')).toEqual({
            text: 'Abweichung',
            tone: 'warn'
        });
        expect(navBadgeFromProgress('', 'Playbook: 4 / 8 Schritte')).toEqual({
            text: '4/8 Schritte',
            tone: 'warn'
        });
    });

    it('mappt Status-Dots für die Sidebar', () => {
        expect(navStatusDotTone('ok', '')).toBe('ok');
        expect(navStatusDotTone('warn', 'x')).toBe('warn');
        expect(navStatusDotTone('', '')).toBe('none');
    });
});
