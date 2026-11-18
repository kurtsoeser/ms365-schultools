/**
 * Demo-Daten Kalender und Termine.
 */

import { normalizeTermin, toIsoDateOnly } from './termine-hub-model.js';

function atDay(offsetDays, hour, minute) {
    const d = new Date();
    d.setDate(d.getDate() + offsetDays);
    d.setHours(hour, minute || 0, 0, 0);
    return d.toISOString();
}

function allDayRange(offsetDays, lengthDays) {
    const start = new Date();
    start.setDate(start.getDate() + offsetDays);
    const end = new Date(start);
    end.setDate(end.getDate() + (lengthDays || 0));
    return {
        start: toIsoDateOnly(start),
        end: toIsoDateOnly(end)
    };
}

export function buildTermineHubDemoState() {
    const arge = allDayRange(3, 0);
    const ferienHinweis = allDayRange(10, 1);

    const termine = [
        normalizeTermin({
            title: 'Lehrerkonferenz',
            start: atDay(1, 15, 0),
            end: atDay(1, 17, 0),
            allDay: false,
            ort: 'Konferenzzimmer',
            beschreibung: 'Monatskonferenz – Tagesordnung folgt per Mail.',
            sichtbarkeit: 'intern',
            terminTyp: 'konferenz',
            zielgruppe: 'Kollegium',
            syncStatus: 'lokal'
        }),
        normalizeTermin({
            title: 'Elternabend 1A',
            start: atDay(2, 18, 0),
            end: atDay(2, 19, 30),
            allDay: false,
            ort: 'Klassenzimmer 1A',
            beschreibung: 'Informationen zum Semesterstart.',
            sichtbarkeit: 'oeffentlich',
            terminTyp: 'sonstiges',
            zielgruppe: 'Eltern 1A',
            syncStatus: 'pending'
        }),
        normalizeTermin({
            title: 'ARGE Mathematik',
            start: arge.start,
            end: arge.end,
            allDay: true,
            ort: 'Bibliothek',
            beschreibung: 'Schwerpunkt Schularbeitsplanung.',
            sichtbarkeit: 'intern',
            terminTyp: 'arge',
            zielgruppe: 'Fachgruppe Mathematik',
            syncStatus: 'lokal'
        }),
        normalizeTermin({
            title: 'Bildungsberatung 3. Jahrgang',
            start: atDay(5, 9, 0),
            end: atDay(5, 12, 0),
            allDay: false,
            ort: 'Beratungsraum',
            beschreibung: 'Einzeltermine nach Anmeldung.',
            sichtbarkeit: 'oeffentlich',
            terminTyp: 'beratung',
            zielgruppe: '3. Jahrgang',
            syncStatus: 'lokal'
        }),
        normalizeTermin({
            title: 'Schulfest – Aufbau',
            start: atDay(8, 13, 0),
            end: atDay(8, 16, 0),
            allDay: false,
            ort: 'Turnsaal / Hof',
            beschreibung: 'Helfer:innen bitte melden.',
            sichtbarkeit: 'intern',
            terminTyp: 'sonstiges',
            zielgruppe: 'Kollegium',
            syncStatus: 'synced',
            graphEventId: 'demo-graph-event-001'
        }),
        normalizeTermin({
            title: 'Bewegliche Ferien (Hinweis)',
            start: ferienHinweis.start,
            end: ferienHinweis.end,
            allDay: true,
            ort: '',
            beschreibung: 'Kein Unterricht – siehe Schulkalender.',
            sichtbarkeit: 'oeffentlich',
            terminTyp: 'sonstiges',
            zielgruppe: 'Alle',
            syncStatus: 'lokal'
        })
    ];

    return {
        termine,
        updatedAt: new Date().toISOString(),
        _demoMeta: { seededOn: toIsoDateOnly(new Date()) }
    };
}
