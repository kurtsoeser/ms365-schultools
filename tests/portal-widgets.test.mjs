import { describe, expect, it } from 'vitest';
import {
    addCanvasTile,
    hasCanvasLayout,
    mergeCanvasTiles,
    removeCanvasTile
} from '../src/shared/portal-canvas-layout.js';
import { PORTAL_WIDGETS, portalWidgetSizes } from '../src/shared/portal-widget-catalog.js';
import { expandPortalWidgets } from '../src/shared/portal-widgets.js';
import { renderPortalBriefingHtml } from '../src/shared/portal-briefing-ui.js';

const NOW = new Date(2026, 10, 23, 8, 0, 0);

describe('Portal-Widgets', () => {
    it('bietet mindestens 15 auswählbare Kacheln', () => {
        expect(PORTAL_WIDGETS.length).toBeGreaterThanOrEqual(15);
        const ids = new Set(PORTAL_WIDGETS.map((widget) => widget.id));
        expect(ids.size).toBe(PORTAL_WIDGETS.length);
    });

    it('baut aus einem Kalender drei Ansichten und eine Zahl', () => {
        const cards = expandPortalWidgets({
            view: 'lehrer',
            now: NOW,
            termine: [
                { title: 'Heute früh', start: '2026-11-23T09:00:00', end: '2026-11-23T10:00:00', sichtbarkeit: 'intern', sourceId: 'me' },
                { title: 'Mittwoch', start: '2026-11-25T11:00:00', end: '2026-11-25T12:00:00', sichtbarkeit: 'oeffentlich' }
            ]
        });
        const byId = new Map(cards.map((card) => [card.widgetId, card]));
        expect(byId.get('termine-heute').lines.map((line) => line.label)).toEqual(['Heute früh']);
        expect(byId.get('termine-liste').lines.map((line) => line.label)).toEqual(['Heute früh', 'Mittwoch']);
        expect(byId.get('termine-woche').days.find((day) => day.today).items.join(' ')).toContain('Heute früh');
        expect(byId.get('termine-zahl').stat).toBe('2');
    });

    it('lässt eine gespeicherte Auswahl streng und ergänzt sie nur auf Wunsch', () => {
        const saved = [{ id: 'termine-liste', x: 0, y: 0, w: 8, h: 4 }];
        const strict = mergeCanvasTiles(['termine-liste', 'termine-woche'], saved, { strict: true });
        expect(strict.map((tile) => tile.id)).toEqual(['termine-liste']);
        const added = addCanvasTile(strict, 'termine-woche', portalWidgetSizes()['termine-woche']);
        expect(added.map((tile) => tile.id)).toEqual(['termine-liste', 'termine-woche']);
        expect(added.find((tile) => tile.id === 'termine-woche').w).toBe(8);
        expect(removeCanvasTile(added, 'termine-liste').map((tile) => tile.id)).toEqual(['termine-woche']);
    });

    it('erkennt ein leeres Raster nur ohne gespeicherten Satz', () => {
        const data = new Map();
        const storage = {
            getItem: (key) => (data.has(key) ? data.get(key) : null),
            setItem: (key, value) => data.set(key, String(value))
        };
        const identity = { email: 'kurt@schule.at', tenantId: 't' };
        expect(hasCanvasLayout(storage, identity)).toBe(false);
        storage.setItem(
            'ms365-portal-canvas-layout-v2',
            JSON.stringify({ v: 2, users: { 't|kurt@schule.at': { tiles: [] } } })
        );
        expect(hasCanvasLayout(storage, identity)).toBe(true);
    });

    it('zeigt im Anordnen-Modus den Kachelkatalog und die Wochenansicht', () => {
        const cards = expandPortalWidgets({
            view: 'lehrer',
            now: NOW,
            termine: [{ title: 'Konferenz', start: '2026-11-23T08:30:00', end: '2026-11-23T09:00:00', sichtbarkeit: 'intern' }],
            freistellung: {
                toolId: 'freistellung-planer',
                title: 'Freistellungen',
                href: 'tools/freistellung-planer.html',
                lines: [{ kicker: '25.11.', label: 'Anna', meta: '4A' }],
                summary: '1 wartet auf Entscheidung',
                attention: true
            }
        });
        const html = renderPortalBriefingHtml(cards, {
            arranging: true,
            picker: true,
            tiles: [
                { id: 'termine-woche', x: 0, y: 0, w: 12, h: 3 },
                { id: 'freistellung-liste', x: 0, y: 3, w: 5, h: 5 }
            ]
        });
        expect(html).toContain('Kalender');
        expect(html).toContain('data-canvas-pick="termine-heute"');
        expect(html).toContain('data-canvas-pick="termine-woche" disabled');
        expect(html).toContain('dash-portal-briefing__week');
        expect(html).toContain('25.11.');
        expect(html).toContain('grid-column:1 / span 5');
        expect(html).toContain('Kachel entfernen');
    });
});
