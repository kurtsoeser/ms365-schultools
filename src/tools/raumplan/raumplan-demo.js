/**
 * Demo-Daten Raumplan (local-first, seedTag, nicht nach SPO pushen).
 */

import { DEMO_SEED_TAG, RAUM_TYP, normalizeRaum } from './raumplan-model.js';

export function buildRaumplanDemoState() {
    const raeume = [
        normalizeRaum({
            code: 'INF1',
            name: 'Informatik 1',
            gebaeude: 'Trakt B',
            stockwerk: '1',
            kapazitaet: 28,
            typ: RAUM_TYP.FACHRAUM,
            notiz: 'PC-Raum, Beamer',
            seedTag: DEMO_SEED_TAG
        }),
        normalizeRaum({
            code: 'INF3',
            name: 'Informatik 3',
            gebaeude: 'Trakt B',
            stockwerk: '1',
            kapazitaet: 24,
            typ: RAUM_TYP.FACHRAUM,
            notiz: 'Notebooks Wagen',
            seedTag: DEMO_SEED_TAG
        }),
        normalizeRaum({
            code: 'PHS',
            name: 'Physik / Chemie',
            gebaeude: 'Trakt C',
            stockwerk: 'EG',
            kapazitaet: 26,
            typ: RAUM_TYP.FACHRAUM,
            notiz: 'Labortische',
            seedTag: DEMO_SEED_TAG
        }),
        normalizeRaum({
            code: 'R7A',
            name: 'Klassenraum 7A',
            gebaeude: 'Hauptgebäude',
            stockwerk: '2',
            kapazitaet: 30,
            typ: RAUM_TYP.KLASSE,
            seedTag: DEMO_SEED_TAG
        }),
        normalizeRaum({
            code: 'R7B',
            name: 'Klassenraum 7B',
            gebaeude: 'Hauptgebäude',
            stockwerk: '2',
            kapazitaet: 30,
            typ: RAUM_TYP.KLASSE,
            seedTag: DEMO_SEED_TAG
        }),
        normalizeRaum({
            code: 'TURNHALLE',
            name: 'Turnhalle',
            gebaeude: 'Sport',
            stockwerk: 'EG',
            kapazitaet: 60,
            typ: RAUM_TYP.SPORTHALLE,
            notiz: 'Große Halle',
            seedTag: DEMO_SEED_TAG
        }),
        normalizeRaum({
            code: 'MUS',
            name: 'Musikraum',
            gebaeude: 'Trakt A',
            stockwerk: 'EG',
            kapazitaet: 22,
            typ: RAUM_TYP.FACHRAUM,
            seedTag: DEMO_SEED_TAG
        }),
        normalizeRaum({
            code: 'SEK',
            name: 'Sekretariat',
            gebaeude: 'Hauptgebäude',
            stockwerk: 'EG',
            kapazitaet: 4,
            typ: RAUM_TYP.BUERO,
            seedTag: DEMO_SEED_TAG
        }),
        normalizeRaum({
            code: 'DIR',
            name: 'Direktion',
            gebaeude: 'Hauptgebäude',
            stockwerk: '1',
            kapazitaet: 6,
            typ: RAUM_TYP.BUERO,
            seedTag: DEMO_SEED_TAG
        }),
        normalizeRaum({
            code: 'AULA',
            name: 'Aula',
            gebaeude: 'Hauptgebäude',
            stockwerk: 'EG',
            kapazitaet: 180,
            typ: RAUM_TYP.SONSTIGES,
            notiz: 'Veranstaltungen',
            seedTag: DEMO_SEED_TAG
        })
    ];

    return {
        raeume,
        planImageDataUrl: '',
        planPdfName: '',
        localDemoOnly: true,
        updatedAt: new Date().toISOString()
    };
}
