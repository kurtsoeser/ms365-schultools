/**
 * Demo-Daten Einsammler (local-first, Quelle „demo“).
 */

import { ANTWORT_STATUS, EINSAMMLER_STATUS, MODUS, newId } from './einsammler-model.js';

/** Seed-Tag für Reset nur der Demo-Zeilen. */
export const EINSAMMLER_DEMO_SEED_TAG = 'einsammler-demo-v1';

/**
 * @returns {{ einsammler: object[], antworten: object[] }}
 */
export function buildEinsammlerDemoState() {
    const ein1 = newId('ein');
    const ein2 = newId('ein');
    const today = new Date();
    const in5 = new Date(today.getTime() + 5 * 86400000).toISOString().slice(0, 10);
    const in10 = new Date(today.getTime() + 10 * 86400000).toISOString().slice(0, 10);

    const einsammler = [
        {
            id: ein1,
            titel: 'Einwilligung Museumsbesuch 3AHW',
            modus: MODUS.EINWILLIGUNG,
            typ: 'exkursion',
            textInhalt:
                'Wir besuchen am kommenden Donnerstag das Technisches Museum. Abfahrt 8:00, Rückkehr ca. 15:30. Bitte stimmen Sie der Teilnahme zu.',
            klasseCodes: '3AHW',
            faelligAm: in5,
            status: EINSAMMLER_STATUS.AKTIV,
            erstellerEmail: 'demo@schule.at',
            erstellerName: 'Demo Lehrkraft',
            quelle: 'demo',
            seedTag: EINSAMMLER_DEMO_SEED_TAG
        },
        {
            id: ein2,
            titel: 'Elternbrief: Schularbeitenplan Q2',
            modus: MODUS.LESEBESTAETIGUNG,
            typ: 'elternbrief',
            textInhalt:
                'Anbei der Schularbeitenplan für das zweite Quartal. Bitte bestätigen Sie die Kenntnisnahme.',
            klasseCodes: '3AHW;3BHW',
            faelligAm: in10,
            status: EINSAMMLER_STATUS.ENTWURF,
            erstellerEmail: 'demo@schule.at',
            erstellerName: 'Demo Direktion',
            quelle: 'demo',
            seedTag: EINSAMMLER_DEMO_SEED_TAG
        }
    ];

    const antworten = [
        {
            id: newId('ant'),
            einsammlerId: ein1,
            schuelerName: 'Anna Berger',
            klasseCode: '3AHW',
            elternEmail: 'berger@example.com',
            status: ANTWORT_STATUS.ZUGESTIMMT,
            antwortAm: new Date().toISOString(),
            kommentar: '',
            quelle: 'demo',
            seedTag: EINSAMMLER_DEMO_SEED_TAG
        },
        {
            id: newId('ant'),
            einsammlerId: ein1,
            schuelerName: 'Ben Hofer',
            klasseCode: '3AHW',
            elternEmail: '',
            status: ANTWORT_STATUS.OFFEN,
            antwortAm: '',
            kommentar: '',
            quelle: 'demo',
            seedTag: EINSAMMLER_DEMO_SEED_TAG
        },
        {
            id: newId('ant'),
            einsammlerId: ein1,
            schuelerName: 'Clara Moser',
            klasseCode: '3AHW',
            elternEmail: 'moser@example.com',
            status: ANTWORT_STATUS.ABGELEHNT,
            antwortAm: new Date().toISOString(),
            kommentar: 'Arzttermin',
            quelle: 'demo',
            seedTag: EINSAMMLER_DEMO_SEED_TAG
        },
        {
            id: newId('ant'),
            einsammlerId: ein2,
            schuelerName: 'Anna Berger',
            klasseCode: '3AHW',
            elternEmail: 'berger@example.com',
            status: ANTWORT_STATUS.OFFEN,
            antwortAm: '',
            kommentar: '',
            quelle: 'demo',
            seedTag: EINSAMMLER_DEMO_SEED_TAG
        },
        {
            id: newId('ant'),
            einsammlerId: ein2,
            schuelerName: 'David Steiner',
            klasseCode: '3BHW',
            elternEmail: '',
            status: ANTWORT_STATUS.OFFEN,
            antwortAm: '',
            kommentar: '',
            quelle: 'demo',
            seedTag: EINSAMMLER_DEMO_SEED_TAG
        }
    ];

    return { einsammler, antworten };
}
