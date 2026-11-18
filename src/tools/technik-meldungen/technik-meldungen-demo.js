/**
 * Demo-Daten Technik-Meldungen (local-first, seedTag).
 */

import {
    DEMO_ME_EMAIL,
    DEMO_ME_NAME,
    DEMO_SCHUELER_EMAIL,
    DEMO_SCHUELER_NAME,
    DEMO_SEED_TAG,
    KATEGORIE,
    PRIORITAET,
    STATUS,
    normalizeMeldung,
    toIsoDateOnly
} from './technik-meldungen-model.js';

export function buildTechnikDemoState() {
    const heute = toIsoDateOnly(new Date());
    const gestern = (() => {
        const d = new Date();
        d.setDate(d.getDate() - 1);
        return d.toISOString();
    })();
    const vor3 = (() => {
        const d = new Date();
        d.setDate(d.getDate() - 3);
        return d.toISOString();
    })();
    const vor7 = (() => {
        const d = new Date();
        d.setDate(d.getDate() - 7);
        return d.toISOString();
    })();

    const meldungen = [
        normalizeMeldung({
            title: 'Beamer flackert',
            beschreibung: 'Bild flackert stark, Unterricht kaum möglich.',
            raumCode: 'PHY1',
            raumName: 'Physiksaal',
            kategorie: KATEGORIE.BEAMER,
            prioritaet: PRIORITAET.DRINGEND,
            status: STATUS.OFFEN,
            melderName: DEMO_ME_NAME,
            melderEmail: DEMO_ME_EMAIL,
            gemeldetAm: gestern,
            kommentar: 'Fernbedienung funktioniert.',
            seedTag: DEMO_SEED_TAG
        }),
        normalizeMeldung({
            title: 'PC startet nicht',
            beschreibung: 'Lehrer-PC bleibt beim Boot-Logo hängen.',
            raumCode: 'INF3',
            raumName: 'Informatik 3',
            kategorie: KATEGORIE.PC,
            prioritaet: PRIORITAET.HOCH,
            status: STATUS.IN_ARBEIT,
            melderName: 'Schmidt',
            melderEmail: 'schmidt@schule.demo',
            gemeldetAm: vor3,
            kommentar: 'IT: Festplatte prüfen – Ticket übernommen.',
            bearbeitetVonName: 'IT Support',
            bearbeitetVonEmail: 'it@schule.demo',
            bearbeitetAm: vor3,
            seedTag: DEMO_SEED_TAG
        }),
        normalizeMeldung({
            title: 'WLAN im Trakt B schwach',
            beschreibung: 'Mehrere Geräte fallen im 2. Stock ständig vom WLAN.',
            raumCode: 'B-2',
            raumName: 'Trakt B 2. Stock',
            kategorie: KATEGORIE.NETZWERK,
            prioritaet: PRIORITAET.MITTEL,
            status: STATUS.OFFEN,
            melderName: DEMO_ME_NAME,
            melderEmail: DEMO_ME_EMAIL,
            gemeldetAm: vor3,
            kommentar: '',
            seedTag: DEMO_SEED_TAG
        }),
        normalizeMeldung({
            title: 'Laptop lädt nicht',
            beschreibung: 'Schul-Laptop aus dem Pool lädt nicht mehr – Ladekabel ok.',
            raumCode: 'AULA',
            raumName: 'Aula',
            kategorie: KATEGORIE.PC,
            prioritaet: PRIORITAET.MITTEL,
            status: STATUS.OFFEN,
            melderName: DEMO_SCHUELER_NAME,
            melderEmail: DEMO_SCHUELER_EMAIL,
            gemeldetAm: gestern,
            kommentar: '',
            seedTag: DEMO_SEED_TAG
        }),
        normalizeMeldung({
            title: 'Lautsprecher knackt',
            beschreibung: 'Audio-Anlage knackt bei Lautstärke über 40 %.',
            raumCode: 'MUS1',
            raumName: 'Musiksaal',
            kategorie: KATEGORIE.AUDIO,
            prioritaet: PRIORITAET.NIEDRIG,
            status: STATUS.ERLEDIGT,
            melderName: 'Weber',
            melderEmail: 'weber@schule.demo',
            gemeldetAm: vor7,
            erledigtAm: heute,
            kommentar: 'Kabel getauscht – erledigt.',
            bearbeitetVonName: 'IT Support',
            bearbeitetVonEmail: 'it@schule.demo',
            bearbeitetAm: heute,
            seedTag: DEMO_SEED_TAG
        })
    ];

    return {
        meldungen,
        localDemoOnly: true,
        updatedAt: new Date().toISOString()
    };
}
