/**
 * Persona-Screenshots für die Homepage (Einblick).
 * Capture + Sync-Landing lesen dieselbe Liste.
 */

/** @typedef {{
 *   id: string,
 *   file: string,
 *   label: string,
 *   urlLabel: string,
 *   caption: string,
 *   captureUrl: string,
 *   prepare?: 'basic'|'demo-click'|'stammdaten'|'organigramm'|'fr-antrag'|'fr-liste'|'fr-schueler'|'sa-list'|'akt-list'|'eb'|'srdp'|'krank'|'ausgaben'
 * }} ScreenShot
 *
 * @typedef {{
 *   id: string,
 *   label: string,
 *   short: string,
 *   lead: string,
 *   screens: ScreenShot[]
 * }} ScreenPersona
 */

/** @type {ScreenPersona[]} */
export const SCREEN_PERSONAS = [
  {
    id: 'it',
    label: 'IT-Admin',
    short: 'IT',
    lead: 'Tenant, Stammdaten, Gruppen und Hygiene – die Werkstatt hinter dem Lehrer-Portal.',
    screens: [
      {
        id: 'dash-it',
        file: 'dash-it.png',
        label: 'IT-Werkzeuge',
        urlLabel: 'Dashboard · IT',
        caption: 'IT-Werkstatt: Mitgliedschaften, Klassen, Schuljahr und Aufräumen in einer Leiste.',
        captureUrl: '/index.html?landingShot=1&persona=it',
        prepare: 'dash-it'
      },
      {
        id: 'stammdaten',
        file: 'einrichtung.png',
        label: 'Stammdaten',
        urlLabel: 'Stammdaten',
        caption: 'Schulprofil, Domain, Import und lokale Stammdaten – Grundlage für alle Tools.',
        captureUrl: '/tenant.html?landingShot=1&persona=it',
        prepare: 'stammdaten'
      },
      {
        id: 'organigramm',
        file: 'organigramm.png',
        label: 'Organigramm',
        urlLabel: 'Schul-Organigramm',
        caption: 'Schulstruktur sichtbar: Klassen, Teams und Zuständigkeiten als Landkarte.',
        captureUrl: '/tenant.html?landingShot=1&persona=it#organigramm',
        prepare: 'organigramm'
      },
      {
        id: 'gruppen',
        file: 'gruppen.png',
        label: 'Gruppen & Teams',
        urlLabel: 'MS 365 Gruppenverwaltung',
        caption: 'Alle Gruppen und Teams einlesen, filtern und live pflegen.',
        captureUrl: '/tools/schulstruktur-sync.html',
        prepare: 'gruppen'
      },
      {
        id: 'kursteams',
        file: 'kursteams.png',
        label: 'Kursteams',
        urlLabel: 'Unterrichtsteams',
        caption: 'Unterrichtsteams aus Exporten vorbereiten und in Microsoft 365 anlegen.',
        captureUrl: '/tools/kursteams.html',
        prepare: 'kursteams'
      },
      {
        id: 'aufraeumen',
        file: 'aufraeumen.png',
        label: 'Aufräumen',
        urlLabel: 'Cleanup-Playbook',
        caption: 'Hygiene und SharePoint-Sicherung in sinnvoller Reihenfolge zum Schuljahr.',
        captureUrl: '/tools/cleanup-playbook.html',
        prepare: 'basic'
      }
    ]
  },
  {
    id: 'direktion',
    label: 'Direktion',
    short: 'Direktion',
    lead: 'Freigaben, Überblick und Entscheidungen – ohne in jeder Liste zu versinken.',
    screens: [
      {
        id: 'fr-direktion',
        file: 'freistellung.png',
        label: 'Freistellungen',
        urlLabel: 'Freistellungen · Alle Anträge',
        caption: 'Alle Schüler-Freistellungen mit Status, Klasse und KV – freigabereif.',
        captureUrl:
          '/tools/freistellung-planer.html?landingShot=1&demoRole=1&role=direktion&view=liste',
        prepare: 'fr-liste'
      },
      {
        id: 'sa-direktion',
        file: 'schularbeiten.png',
        label: 'Schularbeiten',
        urlLabel: 'Schularbeiten-Planer',
        caption: 'Planung und Freigabe von Schularbeitsterminen über das Schuljahr.',
        captureUrl:
          '/tools/schularbeiten-planer.html?landingShot=1&demoRole=1&role=admin',
        prepare: 'sa-list'
      },
      {
        id: 'akt-direktion',
        file: 'schulaktivitaeten.png',
        label: 'Schulaktivitäten',
        urlLabel: 'Schulaktivitäten',
        caption: 'Exkursionen und Aktivitäten mit Genehmigung, Kosten und Nacharbeit.',
        captureUrl:
          '/tools/schulaktivitaeten-planer.html?landingShot=1&demoRole=1&role=admin',
        prepare: 'akt-list'
      },
      {
        id: 'projektwochen',
        file: 'projektwochen.png',
        label: 'Projektwochen',
        urlLabel: 'Projektwochen',
        caption: 'Angebote, Kapazitäten und Buchung – Überblick für die Leitung.',
        captureUrl: '/tools/projektwochen.html?demo=1',
        prepare: 'demo-click'
      }
    ]
  },
  {
    id: 'verwaltung',
    label: 'Verwaltung',
    short: 'Verwaltung',
    lead: 'Sekretariat und Administration: Elternkanal, Ausgaben, Krankmeldung, Termine.',
    screens: [
      {
        id: 'elternbrief',
        file: 'elternbrief.png',
        label: 'Elternbrief',
        urlLabel: 'Elternbrief',
        caption: 'Briefe mit Lesebestätigung – persönliche Links statt Papierchaos.',
        captureUrl: '/tools/elternbrief.html?landingShot=1&demoRole=1',
        prepare: 'eb'
      },
      {
        id: 'krankmeldung',
        file: 'krankmeldung.png',
        label: 'Krankmeldung',
        urlLabel: 'Krankmeldung',
        caption: 'Krankmeldungen digital erfassen und den Stand im Überblick behalten.',
        captureUrl: '/tools/krankmeldung.html?demo=1&demoRole=1',
        prepare: 'krank'
      },
      {
        id: 'ausgaben',
        file: 'ausgaben.png',
        label: 'Ausgaben',
        urlLabel: 'Ausgaben & Zugang',
        caption: 'Schlüssel, Karten und Zugänge – wer was hat, nachvollziehbar.',
        captureUrl: '/tools/ausgaben.html?demo=1&demoRole=1',
        prepare: 'ausgaben'
      },
      {
        id: 'termine',
        file: 'termine.png',
        label: 'Termine',
        urlLabel: 'Kalender und Termine',
        caption: 'Schultermine und Kalenderbezüge für den Betriebsalltag.',
        captureUrl: '/tools/termine-hub.html?landingShot=1&persona=lehrer',
        prepare: 'basic'
      }
    ]
  },
  {
    id: 'lehrer',
    label: 'Lehrkraft',
    short: 'Lehrkraft',
    lead: 'Portal für den Alltag: Planer, Freistellungen, Elternbrief, Matura – ohne IT-Katalog.',
    screens: [
      {
        id: 'dash-lehrer',
        file: 'dash-lehrer.png',
        label: 'Lehrer-Portal',
        urlLabel: 'Dashboard · Lehrkraft',
        caption: 'Guten-Morgen-Portal: die wichtigsten Schul-Apps für den Unterrichtsalltag.',
        captureUrl: '/index.html?landingShot=1&persona=lehrer',
        prepare: 'dash-lehrer'
      },
      {
        id: 'fr-antrag',
        file: 'freistellung-antrag.png',
        label: 'Freistellung beantragen',
        urlLabel: 'Freistellung · Antrag',
        caption: 'Antrag stellen mit Zeitraum und Grund – Freigabe über den bekannten Weg.',
        captureUrl:
          '/tools/freistellung-planer.html?landingShot=1&demoRole=1&role=kv&view=antrag',
        prepare: 'fr-antrag'
      },
      {
        id: 'srdp',
        file: 'srdp.png',
        label: 'Maturaanmeldung',
        urlLabel: 'sRDP-Anmeldung',
        caption: 'Matura-/Diplomprüfung: Anmeldung und Übersicht im Formular.',
        captureUrl: '/tools/srdp-anmeldung.html?demo=1&demoRole=1&role=lehrer',
        prepare: 'srdp'
      },
      {
        id: 'einsammler',
        file: 'einsammler.png',
        label: 'Einsammler',
        urlLabel: 'Einsammler',
        caption: 'Unterschriften und Rückläufe digital einsammeln – Fortschritt sichtbar.',
        captureUrl: '/tools/einsammler.html?landingShot=1&demo=1&demoRole=1',
        prepare: 'demo-click'
      },
      {
        id: 'lehrer-fr',
        file: 'lehrer-freistellung.png',
        label: 'Lehrer-Freistellung',
        urlLabel: 'Freistellungen Lehrkräfte',
        caption: 'Eigene Abwesenheiten der Lehrkräfte – Kalender und Genehmigung.',
        captureUrl: '/tools/lehrer-freistellung-planer.html?landingShot=1&demoRole=1&role=direktion',
        prepare: 'basic'
      },
      {
        id: 'aufgaben',
        file: 'aufgaben.png',
        label: 'Aufgaben',
        urlLabel: 'Aufgaben',
        caption: 'Schulinterne Aufgaben zuweisen und Status nachverfolgen.',
        captureUrl: '/tools/aufgaben.html?demo=1',
        prepare: 'demo-click'
      }
    ]
  },
  {
    id: 'schueler',
    label: 'Schüler/in',
    short: 'Schüler/in',
    lead: 'Antrag, Status, Matura und Wahlen – was Schülerinnen und Schüler selbst erledigen.',
    screens: [
      {
        id: 'fr-schueler',
        file: 'freistellung-schueler.png',
        label: 'Meine Freistellungen',
        urlLabel: 'Freistellung · Meine Anträge',
        caption: 'Eigene Anträge und Status – transparent statt Zettelchaos.',
        captureUrl:
          '/tools/freistellung-planer.html?landingShot=1&demoRole=1&role=schueler&view=meine',
        prepare: 'fr-schueler'
      },
      {
        id: 'fr-schueler-antrag',
        file: 'freistellung-schueler-antrag.png',
        label: 'Antrag stellen',
        urlLabel: 'Freistellung · Neuer Antrag',
        caption: 'Formular für den Freistellungsantrag aus Schülersicht.',
        captureUrl:
          '/tools/freistellung-planer.html?landingShot=1&demoRole=1&role=schueler&view=antrag',
        prepare: 'fr-antrag'
      },
      {
        id: 'srdp-schueler',
        file: 'srdp-schueler.png',
        label: 'Maturaanmeldung',
        urlLabel: 'sRDP · Schüler',
        caption: 'Eigene Matura-/Diplom-Anmeldung im geführten Formular.',
        captureUrl: '/tools/srdp-anmeldung.html?demo=1&demoRole=1&role=schueler',
        prepare: 'srdp'
      },
      {
        id: 'wahl',
        file: 'schulsprecherwahl.png',
        label: 'Schulwahl',
        urlLabel: 'Schulwahlen',
        caption: 'Schulwahlen digital: Kandidaturen, Stimme, Auswertung.',
        captureUrl: '/tools/schulsprecherwahl.html?demo=1',
        prepare: 'demo-click'
      },
      {
        id: 'sa-schueler',
        file: 'schularbeiten-schueler.png',
        label: 'Schularbeiten',
        urlLabel: 'Schularbeiten · Schüler',
        caption: 'Schularbeitstermine der eigenen Klasse einsehen.',
        captureUrl:
          '/tools/schularbeiten-planer.html?landingShot=1&demoRole=1&role=schueler&klasse=3AK',
        prepare: 'sa-list'
      }
    ]
  }
];

export function allLandingShots() {
  /** @type {ScreenShot[]} */
  const out = [];
  const seen = new Set();
  for (const p of SCREEN_PERSONAS) {
    for (const s of p.screens) {
      if (seen.has(s.file)) continue;
      seen.add(s.file);
      out.push(s);
    }
  }
  return out;
}
