/**
 * Demo-Daten Kollegiumsumfragen.
 */

import { draftFromVorlage } from './kollegium-umfragen-templates.js';
import {
    ANTWORT_STATUS,
    DEMO_SEED_TAG,
    STATUS,
    normalizeAntwort,
    normalizeUmfrage,
    toIsoDateOnly
} from './kollegium-umfragen-model.js';

export function buildKollegiumDemoState() {
    const heute = toIsoDateOnly(new Date());
    const d1 = draftFromVorlage('ja-nein-beschluss');
    const d2 = draftFromVorlage('schilf-themen');
    const d3 = draftFromVorlage('kv-rueckmeldung');

    const u1 = normalizeUmfrage({
        ...d1,
        status: STATUS.AKTIV,
        faelligAm: heute,
        empfaengerMails: 'mueller@schule.demo; schmidt@schule.demo; berg@schule.demo',
        erstellerName: 'Direktion Demo',
        erstellerEmail: 'direktion@schule.demo',
        anonym: false,
        seedTag: DEMO_SEED_TAG
    });
    const u2 = normalizeUmfrage({
        ...d2,
        status: STATUS.AKTIV,
        faelligAm: heute,
        empfaengerMails: 'mueller@schule.demo; schmidt@schule.demo',
        erstellerName: 'SCHILF-Team',
        erstellerEmail: 'schilf@schule.demo',
        anonym: true,
        seedTag: DEMO_SEED_TAG
    });
    const u3 = normalizeUmfrage({
        ...d3,
        status: STATUS.ENTWURF,
        empfaengerMails: 'kv.3a@schule.demo; kv.2b@schule.demo',
        erstellerName: 'Direktion Demo',
        erstellerEmail: 'direktion@schule.demo',
        seedTag: DEMO_SEED_TAG
    });

    const q1 = u1.fragen[0];
    const antworten = [
        normalizeAntwort({
            umfrageId: u1.id,
            lehrerEmail: 'mueller@schule.demo',
            lehrerName: 'Müller',
            status: ANTWORT_STATUS.ABGEGEBEN,
            abgegebenAm: new Date().toISOString(),
            seedTag: DEMO_SEED_TAG,
            werte: {
                [q1.id]: 'Ja',
                [u1.fragen[1].id]: 'Sinnvoller Schritt.'
            }
        }),
        normalizeAntwort({
            umfrageId: u1.id,
            lehrerEmail: 'schmidt@schule.demo',
            lehrerName: 'Schmidt',
            status: ANTWORT_STATUS.ABGEGEBEN,
            abgegebenAm: new Date().toISOString(),
            seedTag: DEMO_SEED_TAG,
            werte: {
                [q1.id]: 'Enthaltung'
            }
        }),
        normalizeAntwort({
            umfrageId: u1.id,
            lehrerEmail: 'berg@schule.demo',
            lehrerName: 'Berg',
            status: ANTWORT_STATUS.OFFEN,
            seedTag: DEMO_SEED_TAG,
            werte: {}
        }),
        normalizeAntwort({
            umfrageId: u2.id,
            lehrerEmail: 'mueller@schule.demo',
            lehrerName: 'Müller',
            status: ANTWORT_STATUS.ABGEGEBEN,
            abgegebenAm: new Date().toISOString(),
            seedTag: DEMO_SEED_TAG,
            werte: {
                [u2.fragen[0].id]: ['Digitale Medien / KI', 'Klassenführung'],
                [u2.fragen[2].id]: 4
            }
        }),
        normalizeAntwort({
            umfrageId: u2.id,
            lehrerEmail: 'schmidt@schule.demo',
            lehrerName: 'Schmidt',
            status: ANTWORT_STATUS.OFFEN,
            seedTag: DEMO_SEED_TAG,
            werte: {}
        })
    ];

    return {
        umfragen: [u1, u2, u3],
        antworten,
        localDemoOnly: true,
        updatedAt: new Date().toISOString()
    };
}
