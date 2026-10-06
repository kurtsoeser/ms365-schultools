import { describe, expect, it } from 'vitest';
import {
    catalogToolActionLabel,
    catalogCardDescription,
    truncateCatalogDescription
} from '../src/shared/dashboard-tool-card.js';

describe('dashboard-tool-card', () => {
    it('liefert kontextuelle Button-Labels', () => {
        expect(catalogToolActionLabel('personen-verwaltung', 'Öffnen')).toBe('Suchen');
        expect(catalogToolActionLabel('playbook-intranet', 'Öffnen')).toBe('Playbook');
        expect(catalogToolActionLabel('namenskonvention-audit', 'Prüfen')).toBe('Prüfen');
        expect(catalogToolActionLabel('bildungsportal-stammdaten', 'Roadmap')).toBe('Roadmap');
        expect(catalogToolActionLabel('slg-schueler', 'Öffnen')).toBe('Öffnen');
    });

    it('kürzt lange Beschreibungen', () => {
        const long = 'A'.repeat(200);
        expect(truncateCatalogDescription(long, 50).endsWith('…')).toBe(true);
        expect(catalogCardDescription('x', 'Kurz')).toBe('Kurz');
    });
});
