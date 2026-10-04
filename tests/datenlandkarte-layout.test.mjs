import { describe, it, expect } from 'vitest';
import { layoutDatenlandkarte } from '../src/tools/datenlandkarte/datenlandkarte-layout.js';
import { DATEN_BLOECKE, DATEN_LINKS } from '../src/tools/datenlandkarte/datenlandkarte-catalog.js';

describe('datenlandkarte-layout', () => {
    it('platziert alle Blöcke und erzeugt Pfade für Links', () => {
        const layout = layoutDatenlandkarte(DATEN_BLOECKE, DATEN_LINKS);
        expect(layout.blocks.length).toBe(DATEN_BLOECKE.length);
        expect(layout.links.length).toBeGreaterThan(10);
        expect(layout.width).toBeGreaterThan(400);
        const studentsClasses = layout.links.find((l) => l.id === 'students-classes');
        expect(studentsClasses && studentsClasses.path).toMatch(/^M /);
    });

    it('übernimmt gespeicherte Blockpositionen', () => {
        const saved = { 'stamm-classes': { x: 120, y: 80 } };
        const layout = layoutDatenlandkarte(DATEN_BLOECKE, DATEN_LINKS, saved);
        const b = layout.blocks.find((x) => x.id === 'stamm-classes');
        expect(b).toBeTruthy();
        expect(b.x).toBe(120);
        expect(b.y).toBe(80);
        expect(layout.customized).toBe(true);
    });
});
