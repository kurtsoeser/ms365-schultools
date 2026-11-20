import { describe, it, expect, beforeEach } from 'vitest';
import {
    createInitialState,
    filterItemsForCalendarView,
    loadCalModeSettings,
    persistCalModeSettings
} from '../src/tools/schularbeiten-planer/schularbeiten-planer-state.js';
import { renderApp } from '../src/tools/schularbeiten-planer/schularbeiten-planer-ui.js';

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

    it('speichert Agenda, Woche und Monat', () => {
        persistCalModeSettings('week');
        expect(loadCalModeSettings()).toBe('week');
        persistCalModeSettings('agenda');
        expect(loadCalModeSettings()).toBe('agenda');
        persistCalModeSettings('month');
        expect(loadCalModeSettings()).toBe('month');
        persistCalModeSettings('nope');
        expect(loadCalModeSettings()).toBe('month');
    });
});

describe('Schularbeiten-Ansichten', () => {
    beforeEach(() => {
        const store = new Map();
        globalThis.localStorage = {
            getItem: (k) => store.get(k) ?? null,
            setItem: (k, v) => store.set(k, String(v))
        };
    });

    it('Kalender und Liste folgen der gemeinsamen Planer-Darstellung', () => {
        const state = createInitialState();
        state.bootstrapped = true;
        state.localDemoOnly = true;
        state.role = 'lehrer';
        state.view = 'kalender';
        state.items = [
            {
                itemId: 'sa-1',
                datum: '2026-11-20',
                status: 'fixiert',
                fachCode: 'D',
                klasseCode: '3A',
                titel: 'Grammatik',
                beginnUhrzeit: '08:00',
                dauerMinuten: 100
            }
        ];
        const root = { innerHTML: '' };
        renderApp(state, root);
        expect(root.innerHTML).toContain('class="ms365-cal"');
        expect(root.innerHTML).toContain('data-sa-cal-mode="agenda"');
        expect(root.innerHTML).not.toContain('sa-nav__session');
        expect(root.innerHTML).toContain('sa-top__title');

        state.view = 'liste';
        renderApp(state, root);
        expect(root.innerHTML).toContain('sa-listbar');
        expect(root.innerHTML).toContain('sa-table--rows');
        expect(root.innerHTML).toContain('data-sa-detail="sa-1"');
    });
});
