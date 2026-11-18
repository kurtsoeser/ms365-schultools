/**
 * Demo-Daten Aufgaben (local-first, seedTag, nur ?demo=1 / Button).
 */

import {
    DEMO_ME_EMAIL,
    DEMO_ME_NAME,
    DEMO_SEED_TAG,
    PRIORITAET,
    ROLLE,
    STATUS,
    normalizeAufgabe,
    toIsoDateOnly
} from './aufgaben-model.js';

export function buildAufgabenDemoState() {
    const heute = toIsoDateOnly(new Date());
    const bald = (() => {
        const d = new Date();
        d.setDate(d.getDate() + 2);
        return toIsoDateOnly(d);
    })();
    const spaeter = (() => {
        const d = new Date();
        d.setDate(d.getDate() + 7);
        return toIsoDateOnly(d);
    })();
    const ueberfaellig = (() => {
        const d = new Date();
        d.setDate(d.getDate() - 3);
        return toIsoDateOnly(d);
    })();

    const aufgaben = [
        normalizeAufgabe({
            title: 'Unterlagen für Elternabend nachreichen',
            beschreibung: 'Teilnehmerliste und Tagesordnung an das Sekretariat senden.',
            zuweisungName: DEMO_ME_NAME,
            zuweisungEmail: DEMO_ME_EMAIL,
            rolle: ROLLE.LEHRER,
            prioritaet: PRIORITAET.HOCH,
            frist: bald,
            status: STATUS.OFFEN,
            erstelltVon: 'Direktion',
            erstelltVonEmail: 'direktion@schule.demo',
            erstelltAm: new Date(Date.now() - 86400000).toISOString(),
            kommentar: 'Bitte vor Freitag.',
            seedTag: DEMO_SEED_TAG
        }),
        normalizeAufgabe({
            title: 'Raum 12b auf Vollständigkeit prüfen',
            beschreibung: 'Stühle und Beamer-Fernbedienung nach dem Projekttag kontrollieren.',
            zuweisungName: DEMO_ME_NAME,
            zuweisungEmail: DEMO_ME_EMAIL,
            rolle: ROLLE.KV,
            prioritaet: PRIORITAET.MITTEL,
            frist: heute,
            status: STATUS.IN_ARBEIT,
            erstelltVon: 'Hausverwaltung',
            erstelltVonEmail: 'verwaltung@schule.demo',
            erstelltAm: new Date(Date.now() - 2 * 86400000).toISOString(),
            kommentar: '',
            seedTag: DEMO_SEED_TAG
        }),
        normalizeAufgabe({
            title: 'Elternkontakt wegen Fehlstunden',
            beschreibung: 'KV bittet um Rückruf bei Erziehungsberechtigten von Max Beispiel.',
            zuweisungName: 'Schmidt',
            zuweisungEmail: 'schmidt@schule.demo',
            rolle: ROLLE.KV,
            prioritaet: PRIORITAET.DRINGEND,
            frist: ueberfaellig,
            status: STATUS.OFFEN,
            erstelltVon: DEMO_ME_NAME,
            erstelltVonEmail: DEMO_ME_EMAIL,
            erstelltAm: new Date(Date.now() - 4 * 86400000).toISOString(),
            kommentar: 'Telefonnummer liegt im Sekretariat.',
            seedTag: DEMO_SEED_TAG
        }),
        normalizeAufgabe({
            title: 'Bitte Unterschrift Direktion',
            beschreibung: 'Exkursionsformular Klasse 3A – Unterschrift fehlt noch.',
            zuweisungName: 'Direktion',
            zuweisungEmail: 'direktion@schule.demo',
            rolle: ROLLE.DIREKTION,
            prioritaet: PRIORITAET.NIEDRIG,
            frist: spaeter,
            status: STATUS.ERLEDIGT,
            erstelltVon: DEMO_ME_NAME,
            erstelltVonEmail: DEMO_ME_EMAIL,
            erstelltAm: new Date(Date.now() - 10 * 86400000).toISOString(),
            kommentar: 'Erledigt – Formular liegt im Sekretariat.',
            seedTag: DEMO_SEED_TAG
        })
    ];

    return {
        aufgaben,
        localDemoOnly: true,
        updatedAt: new Date().toISOString()
    };
}
