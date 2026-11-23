import { describe, expect, it } from 'vitest';
import {
    lookupToolCopy,
    normalizeToolHref,
    stripHygieneNoiseFromTitle,
    taskRowDescription,
    taskRowToolId,
    TOOL_PRACTICE_COPY
} from '../src/shared/dashboard-task-tool-copy.js';

describe('dashboard-task-tool-copy', () => {
    it('normalisiert Tool-Pfade', () => {
        expect(normalizeToolHref('./tools/kursteams.html?mode=single')).toBe('tools/kursteams.html');
    });

    it('liefert Schulpraxis-Texte für Klassengruppen', () => {
        const copy = lookupToolCopy('tools/jahrgangsgruppen.html');
        expect(copy).toBeTruthy();
        expect(copy.title).toContain('Klassengruppen');
        expect(copy.desc.length).toBeGreaterThan(10);
    });

    it('unterstützt Query-Varianten und hält Schulleitung vs. Verwaltung getrennt', () => {
        const copy = lookupToolCopy('tools/personen-verwaltung.html?create=1');
        expect(copy.title).toBe('Konto anlegen');

        const sl = lookupToolCopy('tools/verwaltung.html?kind=schulleitung');
        const vw = lookupToolCopy('tools/verwaltung.html?kind=verwaltung');
        expect(sl.title).toBe('Schulleitung');
        expect(vw.title).toBe('Verwaltung (Personal)');
        expect(sl.title).not.toBe(vw.title);
    });

    it('hat Einträge für alle Kern-Werkzeuge der Aufgaben-Kacheln', () => {
        expect(TOOL_PRACTICE_COPY['tools/schueler-sammelgruppe.html']).toBeTruthy();
        expect(TOOL_PRACTICE_COPY['tools/webuntis-sync-monitor.html'].title).toMatch(/Stundenplan/i);
    });

    it('entfernt wiederholte Hygiene-Badge-Texte aus Titeln', () => {
        expect(stripHygieneNoiseFromTitle('Datenhygiene ⚠ Abweichung ⚠ Abweichung')).toBe('Datenhygiene');
    });

    it('liefert Beschreibung und Tool-ID für Sammelgruppen-Kacheln', () => {
        expect(taskRowDescription('tools/schueler-sammelgruppe.html', 'slg-schueler', '')).toMatch(
            /Sammelgruppe/i
        );
        expect(
            taskRowToolId({ getAttribute: (k) => (k === 'data-dash-hygiene-id' ? 'slg-schueler' : '') })
        ).toBe('slg-schueler');
    });
});
