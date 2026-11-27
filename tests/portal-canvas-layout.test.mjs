import { describe, expect, it } from 'vitest';
import {
    canvasLayoutUserKey,
    defaultCanvasTiles,
    mergeCanvasTiles,
    placeCanvasTile,
    readCanvasLayout,
    reorderCanvasTiles,
    writeCanvasLayout
} from '../src/shared/portal-canvas-layout.js';
import { renderPortalBriefingHtml } from '../src/shared/portal-briefing-ui.js';

function memoryStorage() {
    const data = new Map();
    return {
        getItem(key) {
            return data.has(key) ? data.get(key) : null;
        },
        setItem(key, value) {
            data.set(key, String(value));
        }
    };
}

describe('Portal-Canvas', () => {
    it('legt neue Kacheln zu zweit auf zwölf Spalten', () => {
        const tiles = defaultCanvasTiles(['termine-hub', 'schulaktivitaeten-planer', 'elternbrief']);
        expect(tiles.map((t) => [t.id, t.x, t.y, t.w, t.h])).toEqual([
            ['termine-hub', 0, 0, 6, 4],
            ['schulaktivitaeten-planer', 6, 0, 6, 4],
            ['elternbrief', 0, 4, 6, 4]
        ]);
    });

    it('behält gespeicherte Plätze und setzt neue Kacheln daneben', () => {
        const tiles = mergeCanvasTiles(['termine-hub', 'elternbrief'], [
            { id: 'termine-hub', x: 6, y: 0, w: 6, h: 3 },
            { id: 'fremd', x: 0, y: 0, w: 6, h: 4 }
        ]);
        expect(tiles.find((t) => t.id === 'termine-hub')).toMatchObject({ x: 6, y: 0, w: 6, h: 3 });
        expect(tiles.find((t) => t.id === 'elternbrief')).toMatchObject({ x: 0, y: 0, w: 6, h: 4 });
        expect(tiles.some((t) => t.id === 'fremd')).toBe(false);
    });

    it('schiebt eine überdeckte Kachel nach unten', () => {
        const tiles = placeCanvasTile(
            [
                { id: 'a', x: 0, y: 0, w: 6, h: 4 },
                { id: 'b', x: 6, y: 0, w: 6, h: 4 }
            ],
            'a',
            { x: 6, y: 0 }
        );
        expect(tiles.find((t) => t.id === 'a')).toMatchObject({ x: 6, y: 0 });
        expect(tiles.find((t) => t.id === 'b').y).toBeGreaterThanOrEqual(4);
    });

    it('hält die Mindestgröße', () => {
        const tiles = placeCanvasTile([{ id: 'a', x: 0, y: 0, w: 6, h: 4 }], 'a', { w: 1, h: 1, x: 20 });
        expect(tiles[0]).toMatchObject({ w: 2, h: 2, x: 10 });
    });

    it('ordnet auf schmalen Flächen neu, ohne Kacheln zu verlieren', () => {
        const tiles = reorderCanvasTiles(
            [
                { id: 'a', x: 0, y: 0, w: 6, h: 4 },
                { id: 'b', x: 6, y: 0, w: 4, h: 3 }
            ],
            'b',
            0
        );
        expect(tiles.map((t) => t.id)).toEqual(['b', 'a']);
        expect(tiles.find((t) => t.id === 'b').w).toBe(4);
    });

    it('speichert die Anordnung je Konto und Mandant', () => {
        const storage = memoryStorage();
        const kurt = { email: 'Kurt@schule.at', tenantId: 'tenant-a', persona: 'lehrer' };
        const other = { email: 'kurt@schule.at', tenantId: 'tenant-b', persona: 'lehrer' };
        writeCanvasLayout(storage, kurt, [{ id: 'termine-hub', x: 3, y: 1, w: 6, h: 4 }], '2026-11-23T22:00:00.000Z');
        expect(readCanvasLayout(storage, kurt)).toEqual([{ id: 'termine-hub', x: 3, y: 1, w: 6, h: 4 }]);
        expect(readCanvasLayout(storage, other)).toEqual([]);
        expect(canvasLayoutUserKey(kurt)).not.toBe(canvasLayoutUserKey(other));
    });

    it('zeigt Anordnen und im Raster die Fertig-Ansicht', () => {
        const card = {
            toolId: 'termine-hub',
            title: 'Nächste Termine',
            href: 'tools/termine-hub.html',
            lines: [],
            emptyText: 'Keine Termine'
        };
        const calm = renderPortalBriefingHtml([card]);
        expect(calm).toContain('Anordnen');
        expect(calm).toContain('grid-column:1 / span 6');
        const editing = renderPortalBriefingHtml([card], {
            arranging: true,
            tiles: [{ id: 'termine-hub', x: 0, y: 0, w: 6, h: 4 }]
        });
        expect(editing).toContain('Fertig');
        expect(editing).toContain('Größe ändern');
        expect(editing).toContain('auf diesem Gerät');
        expect(editing).not.toContain('href="tools/termine-hub.html"');
    });
});
