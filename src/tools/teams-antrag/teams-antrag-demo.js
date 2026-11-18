/**
 * Demo-Daten Teams-Antrag.
 */
import { ART, STATUS, newId, normalizeAntrag } from './teams-antrag-model.js';

export function buildTeamsAntragDemoState() {
    const now = Date.now();
    const iso = (offsetMs) => new Date(now - offsetMs).toISOString();
    const antraege = [
        normalizeAntrag({
            id: newId('ta'),
            title: 'Robotik-AG',
            zweck: 'Wöchentliches Treffen der Robotik-AG mit Dateiablage und Planung.',
            art: ART.AG,
            namensvorschlag: 'AG Robotik 2026/27',
            schuljahr: '2026/27',
            status: STATUS.BEANTRAGT,
            antragstellerEmail: 'max.mustermann@schueler.example.at',
            antragstellerName: 'Max Mustermann',
            besitzerMails: 'lehrer.robotik@schule.example.at',
            mitgliederMails: '',
            quelle: 'demo',
            createdAt: iso(86400000 * 2),
            updatedAt: iso(86400000)
        }),
        normalizeAntrag({
            id: newId('ta'),
            title: 'Projekt Nachhaltigkeit',
            zweck: 'Gemeinsames Projektteam für den Nachhaltigkeitstag der 4. Klassen.',
            art: ART.PROJEKT,
            namensvorschlag: 'Projekt Nachhaltigkeit 4. Klassen',
            schuljahr: '2026/27',
            status: STATUS.IN_PRUEFUNG,
            antragstellerEmail: 'anna.lehrer@schule.example.at',
            antragstellerName: 'Anna Lehrer',
            besitzerMails: 'anna.lehrer@schule.example.at\nit@schule.example.at',
            finalDisplayName: 'Projekt Nachhaltigkeit 2026/27',
            finalMailNickname: 'prj-202627-nachhaltigkeit',
            quelle: 'demo',
            createdAt: iso(86400000 * 5),
            updatedAt: iso(3600000)
        }),
        normalizeAntrag({
            id: newId('ta'),
            title: 'Fachgruppe Deutsch',
            zweck: 'Austausch und Materialien der Fachgruppe Deutsch.',
            art: ART.FACHGRUPPE,
            namensvorschlag: 'FG Deutsch',
            schuljahr: '2026/27',
            status: STATUS.ANGELEGT,
            antragstellerEmail: 'deutsch@schule.example.at',
            antragstellerName: 'Fachgruppe Deutsch',
            finalDisplayName: 'Fachgruppe Deutsch',
            finalMailNickname: 'fg-deutsch',
            teamId: 'demo-team-id',
            teamWebUrl: 'https://teams.microsoft.com/',
            angelegtAt: iso(86400000 * 10),
            genehmigtAt: iso(86400000 * 11),
            quelle: 'demo',
            createdAt: iso(86400000 * 12),
            updatedAt: iso(86400000 * 10)
        })
    ];
    return { antraege, updatedAt: new Date().toISOString() };
}
