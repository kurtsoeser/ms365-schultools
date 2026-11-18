import { describe, it, expect } from 'vitest';
import {
    clampWizardStep,
    validateWizardStep,
    buildWizardSummary,
    SPAKT_WIZARD_STEP_COUNT,
    SPAKT_WIZARD_STEP_LABELS
} from '../src/tools/sharepoint/sharepoint-liste-schulaktivitaeten-wizard.js';
import {
    AKT_SETUP_GROUP_FIELDS,
    AKT_STAMMDATEN_GROUP_ROLES
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-permissions-ui.js';

describe('schulaktivitaeten listen wizard', () => {
    it('hat 4 Schritte inkl. Einreicher', () => {
        expect(SPAKT_WIZARD_STEP_COUNT).toBe(4);
        expect(SPAKT_WIZARD_STEP_LABELS[1]).toMatch(/Genehmigung/i);
        expect(SPAKT_WIZARD_STEP_LABELS[2]).toMatch(/Einreicher/i);
        expect(clampWizardStep(99)).toBe(4);
    });

    it('fordert Website in Schritt 1', () => {
        expect(validateWizardStep(1, { siteUrl: '' })).toMatch(/Website/i);
        expect(validateWizardStep(1, { siteUrl: 'https://x.sharepoint.com/sites/a' })).toBeNull();
        expect(validateWizardStep(2, { siteUrl: 'https://x.sharepoint.com/sites/a' })).toBeNull();
    });

    it('baut Zusammenfassung mit Genehmigung und Einreicher', () => {
        const s = buildWizardSummary({
            siteUrl: 'https://schule.sharepoint.com/sites/Intranet',
            adminLabel: 'Direktion',
            einreicherLabel: 'DEMO Lehrer:innen',
            skipPerms: false
        });
        expect(s.site).toContain('Intranet');
        expect(s.admin).toBe('Direktion');
        expect(s.einreicher).toBe('DEMO Lehrer:innen');
        expect(s.skipPerms).toBe(false);
    });

    it('Einreicher-Gruppe ist wählbar (nicht gesperrt)', () => {
        expect(AKT_STAMMDATEN_GROUP_ROLES.size).toBe(0);
        expect(AKT_SETUP_GROUP_FIELDS.some((f) => f.role === 'groupLehrer')).toBe(true);
        expect(AKT_SETUP_GROUP_FIELDS.find((f) => f.role === 'groupLehrer').dialogTitle).toMatch(
            /Einreicher/i
        );
    });
});
