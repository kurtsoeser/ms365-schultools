import { describe, expect, it } from 'vitest';
import {
    chipLabelFromPhase,
    clampTourStep,
    isFirstSaveTipDone,
    isLastTourStep,
    isLehrerWelcomeTourDone,
    isWelcomeTourDone,
    lehrerChipLabelFromPhase,
    markFirstSaveTipDone,
    markLehrerWelcomeTourDone,
    markWelcomeTourDone,
    normalizeItFirstRunState,
    normalizeLehrerFirstRunState,
    patchItFirstRunState,
    patchLehrerFirstRunState,
    readItFirstRunState,
    readLehrerFirstRunState,
    resumeTargetFromPhase,
    shouldShowFirstSaveTip,
    shouldShowItFirstRunChip,
    shouldShowLehrerFirstRunChip,
    shouldShowLehrerTourMenuButton,
    shouldStartItStammdatenTour,
    shouldStartItWelcomeTour,
    shouldStartLehrerWelcomeTour
} from '../src/shared/product-tour-logic.js';

function memoryStore(seed) {
    const map = new Map(Object.entries(seed || {}));
    return {
        getItem(k) {
            return map.has(k) ? map.get(k) : null;
        },
        setItem(k, v) {
            map.set(k, String(v));
        },
        removeItem(k) {
            map.delete(k);
        }
    };
}

describe('product-tour-logic', () => {
    it('erkennt Willkommens-Tour als erledigt', () => {
        const store = memoryStore({ 'ms365-onboarding-welcome-v1': 'seen' });
        expect(isWelcomeTourDone(store)).toBe(true);
        markWelcomeTourDone(memoryStore({}), 'skipped');
        expect(isWelcomeTourDone(memoryStore({ 'ms365-onboarding-welcome-v1': 'skipped' }))).toBe(true);
    });

    it('startet IT-Welcome nur für angemeldete IT ohne erledigte Tour', () => {
        expect(
            shouldStartItWelcomeTour({
                loggedIn: true,
                isIt: true,
                welcomeDone: false,
                phase: 'idle'
            })
        ).toBe(true);
        expect(
            shouldStartItWelcomeTour({
                loggedIn: true,
                isIt: false,
                welcomeDone: false
            })
        ).toBe(false);
        expect(
            shouldStartItWelcomeTour({
                loggedIn: true,
                isIt: true,
                welcomeDone: true
            })
        ).toBe(false);
        expect(
            shouldStartItWelcomeTour({
                force: true,
                loggedIn: false,
                isIt: false,
                welcomeDone: true
            })
        ).toBe(true);
    });

    it('startet Stammdaten-Tour nur in Phase stammdaten', () => {
        expect(
            shouldStartItStammdatenTour({
                loggedIn: true,
                isIt: true,
                phase: 'stammdaten'
            })
        ).toBe(true);
        expect(
            shouldStartItStammdatenTour({
                loggedIn: true,
                isIt: true,
                phase: 'dashboard'
            })
        ).toBe(false);
    });

    it('normalisiert und speichert First-Run-State', () => {
        expect(normalizeItFirstRunState({ phase: 'nope', step: -2 }).phase).toBe('idle');
        const store = memoryStore({});
        const next = patchItFirstRunState(store, { phase: 'stammdaten', step: 2 });
        expect(next.phase).toBe('stammdaten');
        expect(readItFirstRunState(store).step).toBe(2);
    });

    it('clamped Tour-Schritte', () => {
        expect(clampTourStep(9, 3)).toBe(2);
        expect(isLastTourStep(2, 3)).toBe(true);
        expect(isLastTourStep(0, 3)).toBe(false);
    });

    it('zeigt Fortschritts-Chip bei offener oder abgebrochener Tour', () => {
        expect(
            shouldShowItFirstRunChip({
                loggedIn: true,
                isIt: true,
                phase: 'stammdaten'
            })
        ).toBe(true);
        expect(
            shouldShowItFirstRunChip({
                loggedIn: true,
                isIt: true,
                phase: 'skipped'
            })
        ).toBe(true);
        expect(
            shouldShowItFirstRunChip({
                loggedIn: true,
                isIt: true,
                phase: 'done'
            })
        ).toBe(false);
        expect(
            shouldShowItFirstRunChip({
                loggedIn: true,
                isIt: true,
                phase: 'dashboard',
                chipDismissed: true
            })
        ).toBe(false);
        expect(resumeTargetFromPhase('stammdaten')).toBe('stammdaten');
        expect(resumeTargetFromPhase('skipped')).toBe('dashboard-restart');
        expect(chipLabelFromPhase('stammdaten')).toContain('Stammdaten');
    });

    it('startet Lehrer-Welcome nur für angemeldete Lehrkräfte ohne IT', () => {
        expect(
            shouldStartLehrerWelcomeTour({
                loggedIn: true,
                isLehrer: true,
                isIt: false,
                welcomeDone: false,
                phase: 'idle'
            })
        ).toBe(true);
        expect(
            shouldStartLehrerWelcomeTour({
                loggedIn: true,
                isLehrer: true,
                isIt: true,
                welcomeDone: false
            })
        ).toBe(false);
        expect(
            shouldStartLehrerWelcomeTour({
                loggedIn: true,
                isLehrer: false,
                isIt: false,
                welcomeDone: false
            })
        ).toBe(false);
        expect(
            shouldStartLehrerWelcomeTour({
                force: true,
                loggedIn: false,
                isLehrer: false,
                welcomeDone: true
            })
        ).toBe(true);
        expect(
            shouldStartLehrerWelcomeTour({
                loggedIn: true,
                isLehrer: true,
                isIt: false,
                welcomeDone: false,
                phase: 'idle',
                portalReady: false
            })
        ).toBe(false);
        expect(
            shouldStartLehrerWelcomeTour({
                force: true,
                loggedIn: true,
                isLehrer: true,
                isIt: false,
                welcomeDone: false,
                portalReady: false
            })
        ).toBe(true);
    });

    it('normalisiert Lehrer-First-Run und Chip/Menü', () => {
        expect(normalizeLehrerFirstRunState({ phase: 'nope' }).phase).toBe('idle');
        const store = memoryStore({});
        patchLehrerFirstRunState(store, { phase: 'dashboard', step: 1 });
        expect(readLehrerFirstRunState(store).phase).toBe('dashboard');
        markLehrerWelcomeTourDone(store, 'seen');
        expect(isLehrerWelcomeTourDone(store)).toBe(true);
        expect(
            shouldShowLehrerFirstRunChip({
                loggedIn: true,
                isLehrer: true,
                isIt: false,
                phase: 'skipped'
            })
        ).toBe(true);
        expect(
            shouldShowLehrerFirstRunChip({
                loggedIn: true,
                isLehrer: true,
                isIt: true,
                phase: 'dashboard'
            })
        ).toBe(false);
        expect(lehrerChipLabelFromPhase('dashboard')).toContain('fortsetzen');
        expect(
            shouldShowLehrerTourMenuButton({
                loggedIn: true,
                isLehrer: true,
                isIt: false
            })
        ).toBe(true);
        expect(
            shouldShowLehrerTourMenuButton({
                loggedIn: true,
                isLehrer: false,
                isIt: true,
                portalPreview: true
            })
        ).toBe(true);
        expect(
            shouldShowLehrerTourMenuButton({
                loggedIn: true,
                isLehrer: false,
                isIt: true,
                portalPreview: false
            })
        ).toBe(false);
    });

    it('zeigt First-Save-Tip nur beim ersten manuellen Speichern', () => {
        expect(
            shouldShowFirstSaveTip({
                tipDone: false,
                reason: 'manual-save',
                isTenantPage: true
            })
        ).toBe(true);
        expect(
            shouldShowFirstSaveTip({
                tipDone: false,
                reason: 'autosave',
                isTenantPage: true
            })
        ).toBe(false);
        expect(
            shouldShowFirstSaveTip({
                tipDone: false,
                reason: 'backup-header',
                isTenantPage: false
            })
        ).toBe(true);
        expect(
            shouldShowFirstSaveTip({
                tipDone: true,
                reason: 'manual-save',
                isTenantPage: true
            })
        ).toBe(false);
        const store = memoryStore({});
        markFirstSaveTipDone(store, 'seen');
        expect(isFirstSaveTipDone(store)).toBe(true);
    });
});
