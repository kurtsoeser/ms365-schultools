import { describe, it, expect } from 'vitest';
import {
    shouldShowItLibraryPromptOnThisPage,
    shouldPromptForMissingItLibrary
} from '../src/shared/stammdaten-sharepoint-it-library-prompt.js';

describe('stammdaten-it-library-prompt', () => {
    it('zeigt Prompt nur auf Dashboard/Tenant', () => {
        expect(shouldShowItLibraryPromptOnThisPage('/index.html')).toBe(true);
        expect(shouldShowItLibraryPromptOnThisPage('/')).toBe(true);
        expect(shouldShowItLibraryPromptOnThisPage('/MS365schule/')).toBe(true);
        expect(shouldShowItLibraryPromptOnThisPage('/tenant.html')).toBe(true);
        expect(shouldShowItLibraryPromptOnThisPage('/dashboard-werkzeug-zugriff.html')).toBe(true);
        expect(shouldShowItLibraryPromptOnThisPage('/tools/')).toBe(false);
        expect(shouldShowItLibraryPromptOnThisPage('/tools/freistellung-planer.html')).toBe(false);
        expect(shouldShowItLibraryPromptOnThisPage('/tools/stammdaten-uebergabe.html')).toBe(false);
    });

    it('entscheidet Prompt anhand Auto-Link-Ergebnis', () => {
        expect(shouldPromptForMissingItLibrary({ linked: true }, '/index.html')).toBe(false);
        expect(shouldPromptForMissingItLibrary({ skipped: 'no-hints' }, '/index.html')).toBe(true);
        expect(shouldPromptForMissingItLibrary({ skipped: 'library-not-found' }, '/index.html')).toBe(
            true
        );
        expect(shouldPromptForMissingItLibrary({ skipped: 'no-hints' }, '/tools/pa-antraege.html')).toBe(
            false
        );
        expect(shouldPromptForMissingItLibrary({ skipped: 'already-configured' }, '/index.html')).toBe(
            false
        );
    });
});
