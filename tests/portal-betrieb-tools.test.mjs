import { describe, it, expect } from 'vitest';
import { isToolVisibleForView, DASHBOARD_LEHRER_SCHUELER_APPS } from '../src/shared/dashboard-audience-catalog.js';
import { LICENSE_PACKAGES } from '../src/shared/license-feature-catalog.js';
import { classifyKostenAmpel, suggestTeachersForKlassen } from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-logic.js';

describe('Lehrer-Portal P1–P5 Wiring', () => {
    it('Whitelist enthält neue Betrieb-Tools', () => {
        for (const id of ['termine-hub', 'aufgaben', 'technik-meldungen', 'raumplan', 'ausgaben']) {
            expect(DASHBOARD_LEHRER_SCHUELER_APPS).toContain(id);
            expect(isToolVisibleForView(id, {}, 'lehrer')).toBe(true);
        }
        expect(isToolVisibleForView('termine-hub', {}, 'schueler')).toBe(true);
        expect(isToolVisibleForView('technik-meldungen', {}, 'schueler')).toBe(true);
        for (const id of ['aufgaben', 'raumplan', 'ausgaben']) {
            expect(isToolVisibleForView(id, {}, 'schueler')).toBe(false);
        }
    });

    it('Lizenzpaket betrieb existiert', () => {
        const pkg = LICENSE_PACKAGES.find((p) => p.id === 'betrieb');
        expect(pkg).toBeTruthy();
        expect(pkg.toolIds).toEqual(
            expect.arrayContaining([
                'termine-hub',
                'aufgaben',
                'sharepoint-liste-aufgaben',
                'technik-meldungen',
                'sharepoint-liste-technik-meldungen',
                'raumplan',
                'sharepoint-liste-raumplan'
            ])
        );
    });
});

describe('Schulaktivitäten P2 Helpers', () => {
    it('classifyKostenAmpel', () => {
        expect(classifyKostenAmpel('Skikurs', 250)).toBe('ok');
        expect(classifyKostenAmpel('Skikurs', 350)).toBe('warn');
        expect(classifyKostenAmpel('Skikurs', 450)).toBe('teuer');
    });

    it('suggestTeachersForKlassen', () => {
        const codes = suggestTeachersForKlassen(['3A'], [
            { code: 'AM', klassen: ['3A', '3B'] },
            { code: 'XY', classCodes: '4A' }
        ]);
        expect(codes).toEqual(['AM']);
    });
});
