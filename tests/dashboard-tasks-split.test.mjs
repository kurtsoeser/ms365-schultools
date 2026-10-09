import { describe, expect, it } from 'vitest';
import {
    clampDashTasksSplitNavWidth,
    navBadgeFromProgress,
    navStatusDotTone,
    readStoredDashTasksSplitNavWidth,
    DASH_TASKS_SPLIT_NAV_WIDTH_DEFAULT,
    DASH_TASKS_SPLIT_NAV_WIDTH_MIN
} from '../src/shared/dashboard-tasks-split.js';

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

    it('begrenzt gespeicherte Nav-Breite', () => {
        expect(clampDashTasksSplitNavWidth(100)).toBe(DASH_TASKS_SPLIT_NAV_WIDTH_MIN);
        expect(clampDashTasksSplitNavWidth(999)).toBe(520);
        expect(clampDashTasksSplitNavWidth(312.7)).toBe(313);
        expect(readStoredDashTasksSplitNavWidth({ getItem: () => '280' })).toBe(280);
        expect(readStoredDashTasksSplitNavWidth({ getItem: () => null })).toBe(
            DASH_TASKS_SPLIT_NAV_WIDTH_DEFAULT
        );
    });
});
