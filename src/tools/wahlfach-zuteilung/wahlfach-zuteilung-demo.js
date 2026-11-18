/**
 * Demo-Daten Wahlfach-Zuteilung (Linz-Auhof Showcase) – local-first, Opt-in.
 */

import {
    normalizeAnmeldung,
    normalizeAngebot,
    normalizeRunde,
    normalizeZuteilung
} from './wahlfach-zuteilung-model.js';

export const DEMO_SEED_TAG = 'wz-demo-auhof-v1';

/**
 * @param {object|null|undefined} state
 * @returns {boolean}
 */
export function isWzDemoSeedState(state) {
    if (!state) return false;
    const r = state.runde;
    if (r && String(r.id || '').indexOf('runde-demo-') === 0) return true;
    return (state.anmeldungen || []).some(
        (a) => a && (a.seedTag === DEMO_SEED_TAG || String(a.id || '').indexOf('anm-demo-') === 0)
    );
}

/**
 * Fester Demo-Snapshot: 6 Angebote, 15 Anmeldungen, teils zugeteilt.
 * @returns {{ runde: object, angebote: object[], anmeldungen: object[], zuteilungen: object[] }}
 */
export function buildWahlfachDemoState() {
    const runde = normalizeRunde({
        id: 'runde-demo-auhof-2627',
        schuljahr: '2026/27',
        titel: 'Wahlfach-Runde Linz-Auhof (Demo)',
        /* Fenster um „heute“ (Nov 2026), damit Schüler-Formular in der Live-Demo offen ist */
        anmeldeVon: '2026-11-01',
        anmeldeBis: '2026-11-30',
        status: 'offen'
    });

    const chor = normalizeAngebot({
        id: 'ang-demo-chor',
        bezeichnung: 'Chor',
        kapazitaetMin: 12,
        kapazitaetMax: 25,
        zielJahrgaenge: '1–4',
        aktiv: true
    });
    const robotik = normalizeAngebot({
        id: 'ang-demo-robotik',
        bezeichnung: 'Robotik',
        kapazitaetMin: 8,
        kapazitaetMax: 16,
        zielJahrgaenge: '3–4',
        aktiv: true
    });
    const theater = normalizeAngebot({
        id: 'ang-demo-theater',
        bezeichnung: 'Theater',
        kapazitaetMin: 10,
        kapazitaetMax: 20,
        zielJahrgaenge: '1–8',
        aktiv: true
    });
    const debating = normalizeAngebot({
        id: 'ang-demo-debating',
        bezeichnung: 'Debating / Rhetorik',
        kapazitaetMin: 8,
        kapazitaetMax: 14,
        zielJahrgaenge: '5–8',
        aktiv: true
    });
    const orchester = normalizeAngebot({
        id: 'ang-demo-orchester',
        bezeichnung: 'Schulorchester',
        kapazitaetMin: 10,
        kapazitaetMax: 18,
        zielJahrgaenge: '1–8',
        aktiv: true
    });
    const sport = normalizeAngebot({
        id: 'ang-demo-sport',
        bezeichnung: 'Sport & Outdoor',
        kapazitaetMin: 12,
        kapazitaetMax: 22,
        zielJahrgaenge: '3–6',
        aktiv: true
    });

    const angebote = [chor, robotik, theater, debating, orchester, sport];

    /** @type {Array<object>} */
    const rawAnmeldungen = [
        {
            id: 'anm-demo-01',
            schuelerName: 'Anna Berger',
            klasse: '3AK',
            schuelerEmail: 'anna.berger@demo.schule',
            wunsch1Id: chor.id,
            wunsch2Id: theater.id,
            wunsch3Id: orchester.id,
            status: 'zugeordnet',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-03T09:12:00Z'
        },
        {
            id: 'anm-demo-02',
            schuelerName: 'Lukas Hofer',
            klasse: '3BK',
            schuelerEmail: 'lukas.hofer@demo.schule',
            wunsch1Id: robotik.id,
            wunsch2Id: sport.id,
            wunsch3Id: chor.id,
            status: 'zugeordnet',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-03T10:05:00Z'
        },
        {
            id: 'anm-demo-03',
            schuelerName: 'Mia Steiner',
            klasse: '4AK',
            schuelerEmail: 'mia.steiner@demo.schule',
            wunsch1Id: theater.id,
            wunsch2Id: chor.id,
            wunsch3Id: debating.id,
            status: 'zugeordnet',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-04T08:40:00Z'
        },
        {
            id: 'anm-demo-04',
            schuelerName: 'Jonas Weber',
            klasse: '5AK',
            schuelerEmail: 'jonas.weber@demo.schule',
            wunsch1Id: debating.id,
            wunsch2Id: theater.id,
            wunsch3Id: robotik.id,
            status: 'zugeordnet',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-04T11:22:00Z'
        },
        {
            id: 'anm-demo-05',
            schuelerName: 'Emma Gruber',
            klasse: '2AK',
            schuelerEmail: 'emma.gruber@demo.schule',
            wunsch1Id: chor.id,
            wunsch2Id: orchester.id,
            wunsch3Id: theater.id,
            status: 'zugeordnet',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-05T07:55:00Z'
        },
        {
            id: 'anm-demo-06',
            schuelerName: 'Noah Leitner',
            klasse: '4BK',
            schuelerEmail: 'noah.leitner@demo.schule',
            wunsch1Id: robotik.id,
            wunsch2Id: theater.id,
            wunsch3Id: sport.id,
            status: 'zugeordnet',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-05T14:10:00Z'
        },
        {
            id: 'anm-demo-07',
            schuelerName: 'Sophie Bauer',
            klasse: '6AK',
            schuelerEmail: 'sophie.bauer@demo.schule',
            wunsch1Id: debating.id,
            wunsch2Id: orchester.id,
            wunsch3Id: theater.id,
            status: 'zugeordnet',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-06T09:30:00Z'
        },
        {
            id: 'anm-demo-08',
            schuelerName: 'Felix Moser',
            klasse: '3AK',
            schuelerEmail: 'felix.moser@demo.schule',
            wunsch1Id: sport.id,
            wunsch2Id: robotik.id,
            wunsch3Id: chor.id,
            status: 'zugeordnet',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-06T12:00:00Z'
        },
        {
            id: 'anm-demo-09',
            schuelerName: 'Lara Fuchs',
            klasse: '1BK',
            schuelerEmail: 'lara.fuchs@demo.schule',
            wunsch1Id: theater.id,
            wunsch2Id: chor.id,
            wunsch3Id: orchester.id,
            status: 'eingereicht',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-07T08:15:00Z'
        },
        {
            id: 'anm-demo-10',
            schuelerName: 'Paul Wagner',
            klasse: '7AK',
            schuelerEmail: 'paul.wagner@demo.schule',
            wunsch1Id: debating.id,
            wunsch2Id: sport.id,
            wunsch3Id: robotik.id,
            status: 'eingereicht',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-07T10:45:00Z'
        },
        {
            id: 'anm-demo-11',
            schuelerName: 'Hannah Mayr',
            klasse: '2BK',
            schuelerEmail: 'hannah.mayr@demo.schule',
            wunsch1Id: orchester.id,
            wunsch2Id: chor.id,
            wunsch3Id: theater.id,
            status: 'eingereicht',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-08T09:00:00Z'
        },
        {
            id: 'anm-demo-12',
            schuelerName: 'David Schmid',
            klasse: '4AK',
            schuelerEmail: 'david.schmid@demo.schule',
            wunsch1Id: robotik.id,
            wunsch2Id: sport.id,
            wunsch3Id: debating.id,
            status: 'eingereicht',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-08T13:20:00Z'
        },
        {
            id: 'anm-demo-13',
            schuelerName: 'Lea Huber',
            klasse: '5BK',
            schuelerEmail: 'lea.huber@demo.schule',
            wunsch1Id: theater.id,
            wunsch2Id: debating.id,
            wunsch3Id: chor.id,
            status: 'ersatz',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-09T08:50:00Z'
        },
        {
            id: 'anm-demo-14',
            schuelerName: 'Tim Aigner',
            klasse: '3BK',
            schuelerEmail: 'tim.aigner@demo.schule',
            wunsch1Id: sport.id,
            wunsch2Id: chor.id,
            wunsch3Id: robotik.id,
            status: 'eingereicht',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-09T15:05:00Z'
        },
        {
            id: 'anm-demo-15',
            schuelerName: 'Nina Pichler',
            klasse: '8AK',
            schuelerEmail: 'nina.pichler@demo.schule',
            wunsch1Id: orchester.id,
            wunsch2Id: debating.id,
            wunsch3Id: theater.id,
            status: 'eingereicht',
            seedTag: DEMO_SEED_TAG,
            eingereichtAm: '2026-02-10T09:40:00Z'
        }
    ];

    const anmeldungen = rawAnmeldungen.map(normalizeAnmeldung);

    /** Teilweise Zuteilung – Rest offen für Live-Demo in „Zuteilung“. */
    const zuteilungen = [
        normalizeZuteilung({
            anmeldungId: 'anm-demo-01',
            angebotId: chor.id,
            begruendung: '1. Wunsch',
            zugeteiltVon: 'SLG Demo',
            zugeteiltAm: '2026-03-01T10:00:00Z'
        }),
        normalizeZuteilung({
            anmeldungId: 'anm-demo-02',
            angebotId: robotik.id,
            begruendung: '1. Wunsch',
            zugeteiltVon: 'SLG Demo',
            zugeteiltAm: '2026-03-01T10:02:00Z'
        }),
        normalizeZuteilung({
            anmeldungId: 'anm-demo-03',
            angebotId: theater.id,
            begruendung: '1. Wunsch',
            zugeteiltVon: 'SLG Demo',
            zugeteiltAm: '2026-03-01T10:04:00Z'
        }),
        normalizeZuteilung({
            anmeldungId: 'anm-demo-04',
            angebotId: debating.id,
            begruendung: '1. Wunsch',
            zugeteiltVon: 'SLG Demo',
            zugeteiltAm: '2026-03-01T10:06:00Z'
        }),
        normalizeZuteilung({
            anmeldungId: 'anm-demo-05',
            angebotId: chor.id,
            begruendung: '1. Wunsch',
            zugeteiltVon: 'SLG Demo',
            zugeteiltAm: '2026-03-01T10:08:00Z'
        }),
        normalizeZuteilung({
            anmeldungId: 'anm-demo-06',
            angebotId: robotik.id,
            begruendung: '1. Wunsch',
            zugeteiltVon: 'SLG Demo',
            zugeteiltAm: '2026-03-01T10:10:00Z'
        }),
        normalizeZuteilung({
            anmeldungId: 'anm-demo-07',
            angebotId: debating.id,
            begruendung: '1. Wunsch',
            zugeteiltVon: 'SLG Demo',
            zugeteiltAm: '2026-03-01T10:12:00Z'
        }),
        normalizeZuteilung({
            anmeldungId: 'anm-demo-08',
            angebotId: sport.id,
            begruendung: '1. Wunsch',
            zugeteiltVon: 'SLG Demo',
            zugeteiltAm: '2026-03-01T10:14:00Z'
        }),
        normalizeZuteilung({
            anmeldungId: 'anm-demo-13',
            angebotId: debating.id,
            begruendung: '2. Wunsch (Theater voll / Demo-Ersatz)',
            zugeteiltVon: 'SLG Demo',
            zugeteiltAm: '2026-03-01T10:20:00Z'
        })
    ];

    return { runde, angebote, anmeldungen, zuteilungen };
}
