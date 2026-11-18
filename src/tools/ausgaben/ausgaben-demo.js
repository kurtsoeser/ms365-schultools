/**
 * Demo-Daten Ausgabenverwaltung.
 */

import {
    DEMO_SEED_TAG,
    EMPFAENGER_TYP,
    VORGANG_STATUS,
    normalizeGegenstand,
    normalizeTyp,
    normalizeVorgang,
    toIsoDateOnly
} from './ausgaben-model.js';
import { TYP_VORLAGEN } from './ausgaben-templates.js';

export function buildAusgabenDemoState() {
    const heute = toIsoDateOnly(new Date());
    const gestern = (() => {
        const d = new Date();
        d.setDate(d.getDate() - 1);
        return toIsoDateOnly(d);
    })();
    const bald = (() => {
        const d = new Date();
        d.setDate(d.getDate() + 3);
        return toIsoDateOnly(d);
    })();
    const ueberfaellig = (() => {
        const d = new Date();
        d.setDate(d.getDate() - 5);
        return toIsoDateOnly(d);
    })();

    const typen = TYP_VORLAGEN.filter((t) => t.id !== 'sonstiges').map((t) =>
        normalizeTyp({ id: t.id, label: t.label, icon: t.icon, hinweis: t.hinweis })
    );

    const g1 = normalizeGegenstand({
        typ: 'schluessel',
        typLabel: 'Schlüssel',
        kennung: 'S-12b',
        bezeichnung: 'Klassenzimmer 12b',
        ortZone: 'Trakt A',
        quelle: 'vorlage',
        seedTag: DEMO_SEED_TAG
    });
    const g2 = normalizeGegenstand({
        typ: 'schluessel',
        typLabel: 'Schlüssel',
        kennung: 'S-GEN',
        bezeichnung: 'Generalschlüssel Gebäude A',
        ortZone: 'Pforte',
        quelle: 'vorlage',
        seedTag: DEMO_SEED_TAG
    });
    const g3 = normalizeGegenstand({
        typ: 'raumkarte',
        typLabel: 'Raumkarte',
        kennung: 'RK-IT-01',
        bezeichnung: 'IT-Raum Karte 01',
        ortZone: 'IT-Trakt',
        quelle: 'vorlage',
        seedTag: DEMO_SEED_TAG
    });
    const g4 = normalizeGegenstand({
        typ: 'schrankenkarte',
        typLabel: 'Schrankenkarte',
        kennung: 'SK-07',
        bezeichnung: 'Schrankenkarte 07',
        ortZone: 'Tor Nord',
        quelle: 'vorlage',
        seedTag: DEMO_SEED_TAG
    });
    const g5 = normalizeGegenstand({
        typ: 'parkplatz',
        typLabel: 'Parkplatz / Zufahrt',
        kennung: 'P-A12',
        bezeichnung: 'Parkplatz A12',
        ortZone: 'Hof West',
        quelle: 'vorlage',
        seedTag: DEMO_SEED_TAG
    });
    const g6 = normalizeGegenstand({
        typ: 'laptop',
        typLabel: 'Laptop / Tablet',
        kennung: 'NB-104',
        bezeichnung: 'Notebook 104',
        ortZone: 'Medienwagen',
        quelle: 'vorlage',
        seedTag: DEMO_SEED_TAG
    });

    const vorgaenge = [
        normalizeVorgang({
            gegenstandId: g1.id,
            kennung: g1.kennung,
            empfaengerTyp: EMPFAENGER_TYP.LEHRER,
            empfaengerName: 'Müller',
            empfaengerEmail: 'mueller@schule.demo',
            klasseCode: '3A',
            ausgegebenAm: gestern,
            faelligAm: bald,
            status: VORGANG_STATUS.AUSGEGEBEN,
            ausgegebenVon: 'Hausverwaltung',
            seedTag: DEMO_SEED_TAG
        }),
        normalizeVorgang({
            gegenstandId: g4.id,
            kennung: g4.kennung,
            empfaengerTyp: EMPFAENGER_TYP.LEHRER,
            empfaengerName: 'Schmidt',
            empfaengerEmail: 'schmidt@schule.demo',
            ausgegebenAm: ueberfaellig,
            faelligAm: ueberfaellig,
            status: VORGANG_STATUS.AUSGEGEBEN,
            kautionHinweis: 'Kaution dokumentiert (kein Zahlungsfluss)',
            ausgegebenVon: 'Sekretariat',
            seedTag: DEMO_SEED_TAG
        }),
        normalizeVorgang({
            gegenstandId: g6.id,
            kennung: g6.kennung,
            empfaengerTyp: EMPFAENGER_TYP.SCHUELER,
            empfaengerName: 'Anna Beispiel',
            klasseCode: '2B',
            ausgegebenAm: gestern,
            zurueckAm: heute,
            status: VORGANG_STATUS.RETOUR,
            unterschriftEmpfang: true,
            ausgegebenVon: 'IT',
            seedTag: DEMO_SEED_TAG
        })
    ];

    return {
        typen,
        gegenstaende: [g1, g2, g3, g4, g5, g6],
        vorgaenge,
        localDemoOnly: true,
        updatedAt: new Date().toISOString()
    };
}
