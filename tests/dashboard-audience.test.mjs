import { describe, it, expect } from 'vitest';
import {
    isToolVisibleForView,
    resolveActiveView,
    listAvailableViews
} from '../src/shared/dashboard-audience-catalog.js';

describe('dashboard-audience-catalog', () => {
    it('IT sieht alle Werkzeuge', () => {
        expect(isToolVisibleForView('jahrgang', {}, 'it')).toBe(true);
        expect(isToolVisibleForView('schularbeiten-planer', {}, 'it')).toBe(true);
    });

    it('Lehrkraft und Schüler sehen nur Planer-Apps', () => {
        expect(isToolVisibleForView('schularbeiten-planer', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('freistellung-planer', {}, 'schueler')).toBe(true);
        expect(isToolVisibleForView('kursteams', {}, 'lehrer')).toBe(false);
        expect(isToolVisibleForView('sharepoint-intranet-hub', {}, 'schueler')).toBe(false);
        expect(isToolVisibleForView('jahrgang', {}, 'lehrer')).toBe(false);
    });

    it('Planer-Apps immer in Lehrkraft/Schüler-Ansicht (Rechte in der App)', () => {
        expect(isToolVisibleForView('schularbeiten-planer', {}, 'lehrer')).toBe(true);
        expect(isToolVisibleForView('freistellung-planer', {}, 'schueler')).toBe(true);
    });

    it('resolveActiveView: IT-Default vor Lehrkraft', () => {
        expect(resolveActiveView(['it', 'lehrer', 'schueler'], null, { isIt: true })).toBe('it');
        expect(resolveActiveView(['it', 'lehrer'], 'lehrer', { isIt: true })).toBe('lehrer');
        expect(resolveActiveView(['lehrer', 'schueler'], null, { isIt: false })).toBe('lehrer');
        expect(resolveActiveView(['schueler'], null, { isIt: false })).toBe('schueler');
    });

    it('listAvailableViews: eine Hauptansicht pro Stufe', () => {
        expect(listAvailableViews({ isIt: true, isLehrer: true, isSchueler: true })).toEqual(['it']);
        expect(listAvailableViews({ isIt: false, isLehrer: true, isSchueler: false })).toEqual(['lehrer']);
        expect(listAvailableViews({ isIt: false, isLehrer: false, isSchueler: true })).toEqual(['schueler']);
    });

});
