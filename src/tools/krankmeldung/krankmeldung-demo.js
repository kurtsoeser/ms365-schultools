/**
 * Demo-Daten Krankmeldung (local-first, Quelle „demo“).
 */

import {
    DEMO_SEED_TAG,
    KATEGORIE,
    STATUS,
    newId,
    toIsoDateOnly
} from './krankmeldung-model.js';

/** Seed-Tag für Reset nur der Demo-Zeilen. */
export const KRANKMELDUNG_DEMO_SEED_TAG = DEMO_SEED_TAG;

/**
 * @returns {{ meldungen: object[], localDemoOnly: boolean }}
 */
export function buildKrankmeldungDemoState() {
    const today = toIsoDateOnly(new Date());
    const tomorrow = (() => {
        const d = new Date();
        d.setDate(d.getDate() + 1);
        return toIsoDateOnly(d);
    })();

    return {
        localDemoOnly: true,
        meldungen: [
            {
                id: newId('km'),
                schuelerName: 'Anna Berger',
                klasseCode: '3AHW',
                von: today,
                bis: today,
                kategorie: KATEGORIE.KRANKHEIT,
                gemeldetAm: new Date().toISOString(),
                gemeldetVonEmail: 'berger@example.com',
                gemeldetVonName: 'Eltern Berger',
                status: STATUS.NEU,
                kenntnisLehrerEmail: '',
                notiz: 'Fieber seit gestern',
                quelle: 'demo',
                seedTag: KRANKMELDUNG_DEMO_SEED_TAG
            },
            {
                id: newId('km'),
                schuelerName: 'Ben Hofer',
                klasseCode: '3AHW',
                von: today,
                bis: tomorrow,
                kategorie: KATEGORIE.ARZT,
                gemeldetAm: new Date(Date.now() - 3600000).toISOString(),
                gemeldetVonEmail: '',
                gemeldetVonName: 'Mutter Hofer',
                status: STATUS.ZUR_KENNTNIS,
                kenntnisLehrerEmail: 'kv@schule.at',
                notiz: '',
                quelle: 'demo',
                seedTag: KRANKMELDUNG_DEMO_SEED_TAG
            },
            {
                id: newId('km'),
                schuelerName: 'Clara Moser',
                klasseCode: '2BHW',
                von: tomorrow,
                bis: tomorrow,
                kategorie: KATEGORIE.SONSTIGES,
                gemeldetAm: new Date(Date.now() - 86400000).toISOString(),
                gemeldetVonEmail: 'moser@example.com',
                gemeldetVonName: '',
                status: STATUS.NEU,
                kenntnisLehrerEmail: '',
                notiz: 'Familienangelegenheit',
                quelle: 'demo',
                seedTag: KRANKMELDUNG_DEMO_SEED_TAG
            }
        ]
    };
}
