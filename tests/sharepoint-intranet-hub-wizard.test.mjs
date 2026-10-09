import { describe, expect, it } from 'vitest';
import {
    IH_WIZARD_STEP_COUNT,
    clampWizardStep,
    validateWizardStep,
    wizardPhaseHint
} from '../src/tools/sharepoint/sharepoint-intranet-hub-wizard.js';

describe('sharepoint-intranet-hub-wizard', () => {
    it('clampWizardStep bounds', () => {
        expect(clampWizardStep(0)).toBe(1);
        expect(clampWizardStep(99)).toBe(IH_WIZARD_STEP_COUNT);
        expect(clampWizardStep(2)).toBe(2);
    });

    it('wizardPhaseHint covers four steps', () => {
        expect(wizardPhaseHint(1)).toMatch(/Schritt 1/);
        expect(wizardPhaseHint(4)).toMatch(/fertigstellen/i);
    });

    it('validateWizardStep requires site URL before listen step', () => {
        expect(validateWizardStep(2, { siteUrl: '' })).toMatch(/Site-URL/);
        expect(validateWizardStep(2, { siteUrl: 'https://x.sharepoint.com/sites/intranet' })).toBeNull();
    });
});
