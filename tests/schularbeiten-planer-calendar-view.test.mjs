import { describe, it, expect, beforeEach } from 'vitest';
import {
    filterItemsForCalendarView,
    loadCalModeSettings,
    persistCalModeSettings
} from '../src/tools/schularbeiten-planer/schularbeiten-planer-state.js';

describe('filterItemsForCalendarView', () => {
    const rows = [
        { itemId: '1', status: 'beantragt' },
        { itemId: '2', status: 'fixiert' },
        { itemId: '3', status: 'abgelehnt' }
    ];

    it('Admin filtert nach calShow', () => {
        const out = filterItemsForCalendarView(rows, {
            role: 'admin',
            calShow: { beantragt: true, fixiert: false, abgelehnt: false }
        });
        expect(out.map((r) => r.itemId)).toEqual(['1']);
    });

    it('Lehrer sieht keine abgelehnten', () => {
        const out = filterItemsForCalendarView(rows, { role: 'lehrer' });
        expect(out.map((r) => r.itemId)).toEqual(['1', '2']);
    });
});

describe('calMode settings', () => {
    beforeEach(() => {
        const store = new Map();
        globalThis.localStorage = {
            getItem: (k) => store.get(k) ?? null,
            setItem: (k, v) => store.set(k, String(v))
        };
    });

    it('speichert Monat/Woche', () => {
        persistCalModeSettings('week');
        expect(loadCalModeSettings()).toBe('week');
        persistCalModeSettings('month');
        expect(loadCalModeSettings()).toBe('month');
    });
});
