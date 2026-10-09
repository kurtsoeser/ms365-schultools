import { describe, it, expect } from 'vitest';
import {
    validateWizardStep,
    buildWizardSummary,
    clampWizardStep
} from '../src/tools/sharepoint/sharepoint-liste-stammdaten-wizard.js';

describe('sharepoint-liste-stammdaten-wizard', () => {
    it('clampWizardStep begrenzt auf 1–5', () => {
        expect(clampWizardStep(0)).toBe(1);
        expect(clampWizardStep(3)).toBe(3);
        expect(clampWizardStep(9)).toBe(5);
    });

    it('validateWizardStep verlangt URL in Schritt 1', () => {
        expect(validateWizardStep(1, { siteUrl: '' })).toMatch(/Website/);
        expect(validateWizardStep(1, { siteUrl: 'https://x.sharepoint.com/sites/i' })).toBeNull();
    });

    it('validateWizardStep verlangt Listen in Schritt 2', () => {
        expect(
            validateWizardStep(2, {
                lists: { schueler: false, faecher: false }
            })
        ).toMatch(/mindestens eine/);
        expect(
            validateWizardStep(2, {
                lists: { schueler: true },
                listTitles: { schueler: 'Schülerinnen' }
            })
        ).toBeNull();
        expect(
            validateWizardStep(2, {
                lists: { schueler: true },
                listTitles: { schueler: '' }
            })
        ).toMatch(/Listenname/);
    });

    it('buildWizardSummary fasst Auswahl zusammen', () => {
        const s = buildWizardSummary({
            siteUrl: 'https://schule.sharepoint.com/sites/intranet',
            syncMode: true,
            removeOrphans: true,
            lists: { schueler: true, faecher: true },
            listTitles: { schueler: 'Schülerinnen', faecher: 'Fächer' },
            skipPerms: true
        });
        expect(s.site).toContain('intranet');
        expect(s.listsText).toMatch(/Schülerinnen/);
        expect(s.perms).toMatch(/überspringen/);
    });
});
