import { describe, expect, it } from 'vitest';
import { specToolDescription } from '../src/shared/dashboard-tool-spec-copy.js';

describe('dashboard-tool-spec-copy', () => {
    it('liefert Spec-Beschreibungen für Kern-Werkzeuge', () => {
        expect(specToolDescription('slg-schueler')).toMatch(/Sammelgruppe/i);
        expect(specToolDescription('freistellung-planer')).toMatch(/Freistellung/i);
        expect(specToolDescription('pa-schularbeiten-mail')).toMatch(/Entwicklung/i);
    });
});
