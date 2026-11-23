/**
 * Auto-generated – nicht von Hand editieren.
 * Quelle: license-feature-catalog + dashboard-audience-catalog
 * Neu erzeugen: node scripts/sync-landing-offer-catalog.mjs
 */
window.MS365_OFFER_CATALOG = {
  "generatedAt": "2026-11-23T15:41:14.600Z",
  "contactMail": "kontakt@ms365.schule",
  "bundles": [
    {
      "id": "starter-it",
      "label": "Starter IT",
      "audience": "IT & Administration",
      "priceEuro": 390,
      "priceLabel": "390 € / Schuljahr",
      "packageIds": [
        "core",
        "personen",
        "schuljahr",
        "unterricht"
      ],
      "featured": false,
      "tagline": "Schuljahr, Personen, Kursteams und Hygiene – der IT-Einstieg.",
      "highlights": [
        "Stammdaten, Personen, Gäste, Lizenzen",
        "Schuljahreswechsel & Import (WebUntis/Sokrates)",
        "Kursteams, Klassengruppen, Diplom, ARGEs",
        "Hygiene, Organigramm, Teams-Antrag"
      ]
    },
    {
      "id": "schulalltag",
      "label": "Schulalltag",
      "audience": "IT + Lehrkräfte / Direktion",
      "priceEuro": 890,
      "priceLabel": "890 € / Schuljahr",
      "packageIds": [
        "core",
        "personen",
        "schuljahr",
        "unterricht",
        "intranet",
        "schulapps",
        "freistellung"
      ],
      "featured": true,
      "tagline": "Starter plus Lehrer-Portal: Planer, Freistellungen, Intranet.",
      "highlights": [
        "Alles aus Starter IT",
        "Schularbeiten, Aktivitäten, Projektwochen, Schulwahlen",
        "Schüler- & Lehrer-Freistellungen",
        "Intranet-Listen + Lehrer-Portal"
      ]
    },
    {
      "id": "schule-plus",
      "label": "Schule Plus",
      "audience": "Gesamtschule inkl. Support-Story",
      "priceEuro": 1690,
      "priceLabel": "1.690 € / Schuljahr",
      "packageIds": [
        "core",
        "personen",
        "schuljahr",
        "unterricht",
        "intranet",
        "schulapps",
        "freistellung",
        "ablauf",
        "ausgaben",
        "betrieb",
        "automationen"
      ],
      "featured": false,
      "tagline": "Schulalltag plus Elternkanal, Ausgaben, Betrieb und Automationen.",
      "highlights": [
        "Alles aus Schulalltag",
        "Elternbrief, Einsammler, Krankmeldung, Kollegium",
        "Ausgaben (Schlüssel/Karten)",
        "Automationen & Betrieb (Termine, Technik, Raum)"
      ]
    }
  ],
  "modules": [
    {
      "id": "core",
      "label": "Basis (Core)",
      "priceEuro": 250,
      "blurb": "Gruppen, Hygiene, Organigramm und schulweite Grundeinstellungen.",
      "requires": [],
      "highlights": [
        "Alle Gruppen und Teams",
        "Schüler:innen-Sammelgruppe",
        "Lehrer:innen-Sammelgruppe",
        "Schulleitung & Verwaltung",
        "Klassenvorstände"
      ],
      "tools": [
        {
          "id": "slg-schueler",
          "label": "Schüler:innen-Sammelgruppe",
          "blurb": "",
          "setup": false
        },
        {
          "id": "slg-lehrer",
          "label": "Lehrer:innen-Sammelgruppe",
          "blurb": "",
          "setup": false
        },
        {
          "id": "verwaltung",
          "label": "Schulleitung & Verwaltung",
          "blurb": "",
          "setup": false
        },
        {
          "id": "klassenvorstaende",
          "label": "Klassenvorstände",
          "blurb": "",
          "setup": false
        },
        {
          "id": "arge-fachgruppen",
          "label": "Fächer und ARGEs",
          "blurb": "",
          "setup": false
        },
        {
          "id": "klassenchats",
          "label": "KlassenChats",
          "blurb": "KlassenChats für den Alltag anlegen und pflegen.",
          "setup": false
        },
        {
          "id": "weitere-teams-gruppen",
          "label": "Weitere Teams & Gruppen",
          "blurb": "",
          "setup": false
        },
        {
          "id": "teams-antrag",
          "label": "Teams-Antrag",
          "blurb": "Neues Team beantragen – IT prüft und legt es an.",
          "setup": false
        },
        {
          "id": "sharepoint-liste-teams-antrag",
          "label": "Teams-Antrag-Liste",
          "blurb": "",
          "setup": true
        },
        {
          "id": "datenhygiene",
          "label": "Datenhygiene",
          "blurb": "",
          "setup": false
        },
        {
          "id": "schulstruktur-sync",
          "label": "Alle Gruppen und Teams",
          "blurb": "",
          "setup": false
        },
        {
          "id": "datenlandkarte",
          "label": "Datenlandkarte",
          "blurb": "",
          "setup": false
        },
        {
          "id": "schul-organigramm",
          "label": "Schul-Organigramm",
          "blurb": "",
          "setup": false
        },
        {
          "id": "schulgraph",
          "label": "Schuldaten-Karte",
          "blurb": "",
          "setup": false
        },
        {
          "id": "stammdaten-uebergabe",
          "label": "Stammdaten-Übergabe",
          "blurb": "",
          "setup": false
        },
        {
          "id": "it-berechtigungen",
          "label": "Berechtigungen",
          "blurb": "",
          "setup": false
        },
        {
          "id": "stammdaten-backup-abgleich",
          "label": "Backup-Abgleich",
          "blurb": "",
          "setup": false
        },
        {
          "id": "cleanup-playbook",
          "label": "Cleanup-Playbook",
          "blurb": "",
          "setup": false
        },
        {
          "id": "datei-migration",
          "label": "Datei-Migration",
          "blurb": "",
          "setup": false
        },
        {
          "id": "nextcloud-teams-migrate",
          "label": "Nextcloud → Teams",
          "blurb": "",
          "setup": false
        },
        {
          "id": "leere-gruppen-report",
          "label": "Leere Gruppen finden",
          "blurb": "",
          "setup": false
        },
        {
          "id": "teams-archiv",
          "label": "Teams archivieren",
          "blurb": "",
          "setup": false
        },
        {
          "id": "teams-kanaele-bulk",
          "label": "Kanäle bulk anlegen",
          "blurb": "",
          "setup": false
        },
        {
          "id": "gruppenerstellung",
          "label": "Wer darf Teams anlegen?",
          "blurb": "",
          "setup": false
        },
        {
          "id": "sharepoint-mandant-website",
          "label": "Neue Websites",
          "blurb": "",
          "setup": false
        },
        {
          "id": "sharepoint-mandant-teilen",
          "label": "Dateien teilen",
          "blurb": "",
          "setup": false
        },
        {
          "id": "schul-baseline",
          "label": "Schul-Baseline",
          "blurb": "",
          "setup": false
        }
      ],
      "searchText": "basis (core) gruppen, hygiene, organigramm und schulweite grundeinstellungen. schüler:innen-sammelgruppe lehrer:innen-sammelgruppe schulleitung & verwaltung klassenvorstände fächer und arges klassenchats weitere teams & gruppen teams-antrag teams-antrag-liste datenhygiene alle gruppen und teams datenlandkarte schul-organigramm schuldaten-karte stammdaten-übergabe berechtigungen backup-abgleich cleanup-playbook datei-migration nextcloud → teams leere gruppen finden teams archivieren kanäle bulk anlegen wer darf teams anlegen? neue websites dateien teilen schul-baseline slg schueler slg lehrer verwaltung klassenvorstaende arge fachgruppen klassenchats weitere teams gruppen teams antrag sharepoint liste teams antrag datenhygiene schulstruktur sync datenlandkarte schul organigramm schulgraph stammdaten uebergabe it berechtigungen stammdaten backup abgleich cleanup playbook datei migration nextcloud teams migrate leere gruppen report teams archiv teams kanaele bulk gruppenerstellung sharepoint mandant website sharepoint mandant teilen schul baseline"
    },
    {
      "id": "personen",
      "label": "Personen",
      "priceEuro": 100,
      "blurb": "Konten, Gäste, Lifecycle und M365-Lizenzübersicht.",
      "requires": [],
      "highlights": [
        "Personen suchen",
        "Lizenzverwaltung",
        "Namenskonvention prüfen",
        "Gäste",
        "Schüler-Lifecycle"
      ],
      "tools": [
        {
          "id": "lizenzverwaltung",
          "label": "Lizenzverwaltung",
          "blurb": "",
          "setup": false
        },
        {
          "id": "personen-verwaltung",
          "label": "Personen suchen",
          "blurb": "",
          "setup": false
        },
        {
          "id": "namenskonvention-audit",
          "label": "Namenskonvention prüfen",
          "blurb": "",
          "setup": false
        },
        {
          "id": "gaeste-verwalten",
          "label": "Gäste",
          "blurb": "",
          "setup": false
        },
        {
          "id": "schueler-lifecycle",
          "label": "Schüler-Lifecycle",
          "blurb": "",
          "setup": false
        },
        {
          "id": "pa-gast-erinnerung",
          "label": "Gast-Erinnerung Flow",
          "blurb": "",
          "setup": true
        }
      ],
      "searchText": "personen konten, gäste, lifecycle und m365-lizenzübersicht. lizenzverwaltung personen suchen namenskonvention prüfen gäste schüler-lifecycle gast-erinnerung flow lizenzverwaltung personen verwaltung namenskonvention audit gaeste verwalten schueler lifecycle pa gast erinnerung"
    },
    {
      "id": "schuljahr",
      "label": "Schuljahr",
      "priceEuro": 130,
      "blurb": "Import, Schuljahreswechsel und Cleanup.",
      "requires": [],
      "highlights": [
        "Schuljahr wechseln",
        "Bildungsportal (Vorschau)",
        "Klassen umbenennen",
        "WebUntis-Sync-Monitor",
        "Daten importieren"
      ],
      "tools": [
        {
          "id": "playbook-schuljahresstart",
          "label": "Playbook Schuljahresstart",
          "blurb": "",
          "setup": true
        },
        {
          "id": "playbook-daten-import-verknuepfen",
          "label": "Erste Einrichtung (Blaupause)",
          "blurb": "",
          "setup": true
        },
        {
          "id": "bildungsportal-stammdaten",
          "label": "Bildungsportal (Vorschau)",
          "blurb": "",
          "setup": false
        },
        {
          "id": "organisations-assistent",
          "label": "Schuljahr wechseln",
          "blurb": "",
          "setup": false
        },
        {
          "id": "klassen-umbenennen",
          "label": "Klassen umbenennen",
          "blurb": "",
          "setup": false
        },
        {
          "id": "webuntis-sync-monitor",
          "label": "WebUntis-Sync-Monitor",
          "blurb": "",
          "setup": false
        },
        {
          "id": "webuntis-stammdaten-import",
          "label": "Daten importieren",
          "blurb": "",
          "setup": false
        },
        {
          "id": "cleanup-playbook-schuljahr",
          "label": "Cleanup zum Schuljahr",
          "blurb": "",
          "setup": false
        }
      ],
      "searchText": "schuljahr import, schuljahreswechsel und cleanup. playbook schuljahresstart erste einrichtung (blaupause) bildungsportal (vorschau) schuljahr wechseln klassen umbenennen webuntis-sync-monitor daten importieren cleanup zum schuljahr playbook schuljahresstart playbook daten import verknuepfen bildungsportal stammdaten organisations assistent klassen umbenennen webuntis sync monitor webuntis stammdaten import cleanup playbook schuljahr"
    },
    {
      "id": "unterricht",
      "label": "Unterricht",
      "priceEuro": 200,
      "blurb": "Unterrichtsteams, OneNote, Diplom und Spielwiesen.",
      "requires": [],
      "highlights": [
        "Unterrichtsteams",
        "Diplomarbeiten / VWA",
        "OneNote-Inhalte verteilen",
        "Klassengruppen",
        "Klassen zusammenlegen"
      ],
      "tools": [
        {
          "id": "jahrgang",
          "label": "Klassengruppen",
          "blurb": "",
          "setup": false
        },
        {
          "id": "klassen-merge",
          "label": "Klassen zusammenlegen",
          "blurb": "",
          "setup": false
        },
        {
          "id": "kursteams",
          "label": "Unterrichtsteams",
          "blurb": "Unterrichtsteams anlegen und Mitglieder steuern.",
          "setup": false
        },
        {
          "id": "playbook-kursteams",
          "label": "Playbook Unterrichtsteams",
          "blurb": "",
          "setup": true
        },
        {
          "id": "unterrichtsteams-katalog",
          "label": "Unterrichtsteams-Katalog",
          "blurb": "Übersicht aller Unterrichtsteams der Schule.",
          "setup": false
        },
        {
          "id": "kursteam-einzeln",
          "label": "Einzelne Unterrichtsteams",
          "blurb": "Einzelnes Unterrichtsteam anlegen oder nachziehen.",
          "setup": false
        },
        {
          "id": "kursteam-templates",
          "label": "Kursteam-Vorlagen",
          "blurb": "",
          "setup": false
        },
        {
          "id": "onenote-verteilung",
          "label": "OneNote-Inhalte verteilen",
          "blurb": "OneNote-Inhalte an Klassen und Kurse verteilen.",
          "setup": false
        },
        {
          "id": "diplomarbeiten",
          "label": "Diplomarbeiten / VWA",
          "blurb": "Abschlussarbeiten: Jahrgangs-Team und private Kanäle.",
          "setup": false
        },
        {
          "id": "pa-diplom-ordner",
          "label": "Diplom-Ordner Flow",
          "blurb": "",
          "setup": true
        },
        {
          "id": "spielwiesen",
          "label": "Spielwiesen-Teams",
          "blurb": "",
          "setup": false
        }
      ],
      "searchText": "unterricht unterrichtsteams, onenote, diplom und spielwiesen. klassengruppen klassen zusammenlegen unterrichtsteams playbook unterrichtsteams unterrichtsteams-katalog einzelne unterrichtsteams kursteam-vorlagen onenote-inhalte verteilen diplomarbeiten / vwa diplom-ordner flow spielwiesen-teams jahrgang klassen merge kursteams playbook kursteams unterrichtsteams katalog kursteam einzeln kursteam templates onenote verteilung diplomarbeiten pa diplom ordner spielwiesen matura kursteam unterrichtsteam onenote"
    },
    {
      "id": "kommunikation",
      "label": "Kommunikation",
      "priceEuro": 150,
      "blurb": "Postfächer, Verteiler, Elternsprechtag und Räume.",
      "requires": [],
      "highlights": [
        "Gemeinsame Postfächer",
        "Kalender-Postfächer",
        "E-Mail-Verteiler",
        "Eltern-Verteiler",
        "Elternsprechtag Bookings"
      ],
      "tools": [
        {
          "id": "postfaecher",
          "label": "Gemeinsame Postfächer",
          "blurb": "",
          "setup": false
        },
        {
          "id": "kalender-postfaecher",
          "label": "Kalender-Postfächer",
          "blurb": "",
          "setup": false
        },
        {
          "id": "verteilerlisten",
          "label": "E-Mail-Verteiler",
          "blurb": "",
          "setup": false
        },
        {
          "id": "eltern-verteiler",
          "label": "Eltern-Verteiler",
          "blurb": "",
          "setup": false
        },
        {
          "id": "elternsprechtag-bookings",
          "label": "Elternsprechtag Bookings",
          "blurb": "Elternsprechtag mit Bookings organisieren.",
          "setup": false
        },
        {
          "id": "playbook-eltern",
          "label": "Playbook Elternkommunikation",
          "blurb": "",
          "setup": true
        },
        {
          "id": "playbook-elternsprechtag",
          "label": "Playbook Elternsprechtag",
          "blurb": "",
          "setup": true
        },
        {
          "id": "raeume-ressourcen",
          "label": "Räume und Ressourcen",
          "blurb": "",
          "setup": false
        }
      ],
      "searchText": "kommunikation postfächer, verteiler, elternsprechtag und räume. gemeinsame postfächer kalender-postfächer e-mail-verteiler eltern-verteiler elternsprechtag bookings playbook elternkommunikation playbook elternsprechtag räume und ressourcen postfaecher kalender postfaecher verteilerlisten eltern verteiler elternsprechtag bookings playbook eltern playbook elternsprechtag raeume ressourcen eltern raum"
    },
    {
      "id": "intranet",
      "label": "Intranet & Listen",
      "priceEuro": 130,
      "blurb": "Schul-Intranet und SharePoint-Listen.",
      "requires": [],
      "highlights": [
        "Maturaanmeldung (sRDP)",
        "Schul-Intranet"
      ],
      "tools": [
        {
          "id": "sharepoint-intranet-hub",
          "label": "Schul-Intranet",
          "blurb": "Schul-Intranet öffnen und Navigation nutzen.",
          "setup": false
        },
        {
          "id": "playbook-intranet",
          "label": "Playbook Intranet",
          "blurb": "",
          "setup": true
        },
        {
          "id": "sharepoint-liste-lehrer",
          "label": "Lehrerliste",
          "blurb": "",
          "setup": true
        },
        {
          "id": "sharepoint-liste-stammdaten",
          "label": "Stammdaten-Listen",
          "blurb": "",
          "setup": true
        },
        {
          "id": "sharepoint-liste-schultermine",
          "label": "Schultermine-Liste",
          "blurb": "",
          "setup": true
        },
        {
          "id": "sharepoint-liste-srdp",
          "label": "sRDP-Liste einrichten",
          "blurb": "",
          "setup": true
        },
        {
          "id": "srdp-anmeldung",
          "label": "Maturaanmeldung (sRDP)",
          "blurb": "Matura-/Diplomprüfungsanmeldung (sRDP) für Abschlussjahrgänge.",
          "setup": false
        },
        {
          "id": "sharepoint-liste-vertretung",
          "label": "Vertretungsplan-Liste",
          "blurb": "",
          "setup": true
        }
      ],
      "searchText": "intranet & listen schul-intranet und sharepoint-listen. schul-intranet playbook intranet lehrerliste stammdaten-listen schultermine-liste srdp-liste einrichten maturaanmeldung (srdp) vertretungsplan-liste sharepoint intranet hub playbook intranet sharepoint liste lehrer sharepoint liste stammdaten sharepoint liste schultermine sharepoint liste srdp srdp anmeldung sharepoint liste vertretung matura srdp"
    },
    {
      "id": "schulapps",
      "label": "Schul-Apps",
      "priceEuro": 300,
      "blurb": "Schularbeiten, Aktivitäten, Projektwochen und Wahl.",
      "requires": [
        "intranet"
      ],
      "highlights": [
        "Schularbeiten-Planer",
        "Schulaktivitäten",
        "Projektwochen",
        "Schulwahlen",
        "Wahlfach-Zuteilung"
      ],
      "tools": [
        {
          "id": "playbook-schularbeiten-planer",
          "label": "Playbook Schularbeiten-Planer",
          "blurb": "",
          "setup": true
        },
        {
          "id": "sharepoint-liste-schularbeiten",
          "label": "Schularbeiten-Listen (IT)",
          "blurb": "",
          "setup": true
        },
        {
          "id": "sharepoint-liste-schulaktivitaeten",
          "label": "Schulaktivitäten-Listen",
          "blurb": "",
          "setup": true
        },
        {
          "id": "sharepoint-liste-projektwochen",
          "label": "Projektwochen-Listen",
          "blurb": "",
          "setup": true
        },
        {
          "id": "projektwochen",
          "label": "Projektwochen",
          "blurb": "Projektwochen planen, Gruppen und Termine steuern.",
          "setup": false
        },
        {
          "id": "schularbeiten-planer",
          "label": "Schularbeiten-Planer",
          "blurb": "Termine, Anträge und Kalender für Ihre Klasse.",
          "setup": false
        },
        {
          "id": "schulaktivitaeten-planer",
          "label": "Schulaktivitäten",
          "blurb": "Exkursionen und Schulaktivitäten beantragen und verfolgen.",
          "setup": false
        },
        {
          "id": "schulaktivitaeten-setup",
          "label": "Schulaktivitäten Setup",
          "blurb": "",
          "setup": true
        },
        {
          "id": "termine-hub",
          "label": "Kalender und Termine",
          "blurb": "Schul-, Lehrer- und Schülerkalender aus Outlook – rollenbasiert, Woche/Monat, anlegen und pflegen.",
          "setup": false
        },
        {
          "id": "aufgaben",
          "label": "Aufgaben",
          "blurb": "Aufgaben an Personen oder Rollen zuweisen und abhaken.",
          "setup": false
        },
        {
          "id": "technik-meldungen",
          "label": "Technik-Meldungen",
          "blurb": "Defekte oder problematische Technik melden – Lehrkräfte und Schüler; IT bearbeitet die Inbox.",
          "setup": false
        },
        {
          "id": "raumplan",
          "label": "Raumplan",
          "blurb": "Räume suchen, Plan ansehen, CSV importieren.",
          "setup": false
        },
        {
          "id": "schulsprecherwahl",
          "label": "Schulwahlen",
          "blurb": "Schulwahlen digital durchführen und auswerten.",
          "setup": false
        },
        {
          "id": "wahlfach-zuteilung",
          "label": "Wahlfach-Zuteilung",
          "blurb": "Wahlfächer erfassen, zuteilen und Übersicht behalten.",
          "setup": false
        },
        {
          "id": "wahlfach-anmeldung",
          "label": "Wahlfach anmelden",
          "blurb": "1.–3. Wunsch für Wahlfächer eintragen.",
          "setup": false
        },
        {
          "id": "sharepoint-liste-wahlfach",
          "label": "Wahlfach-Listen",
          "blurb": "",
          "setup": true
        }
      ],
      "searchText": "schul-apps schularbeiten, aktivitäten, projektwochen und wahl. playbook schularbeiten-planer schularbeiten-listen (it) schulaktivitäten-listen projektwochen-listen projektwochen schularbeiten-planer schulaktivitäten schulaktivitäten setup kalender und termine aufgaben technik-meldungen raumplan schulwahlen wahlfach-zuteilung wahlfach anmelden wahlfach-listen playbook schularbeiten planer sharepoint liste schularbeiten sharepoint liste schulaktivitaeten sharepoint liste projektwochen projektwochen schularbeiten planer schulaktivitaeten planer schulaktivitaeten setup termine hub aufgaben technik meldungen raumplan schulsprecherwahl wahlfach zuteilung wahlfach anmeldung sharepoint liste wahlfach schulwahl schulwahlen wahl schulsprecher technik raum"
    },
    {
      "id": "freistellung",
      "label": "Freistellungen",
      "priceEuro": 150,
      "blurb": "Schüler- und Lehrer-Freistellungen inkl. Setup.",
      "requires": [
        "schulapps"
      ],
      "highlights": [
        "Freistellungen (Schüler)",
        "Freistellungen Lehrkräfte"
      ],
      "tools": [
        {
          "id": "freistellung-planer",
          "label": "Freistellungen (Schüler)",
          "blurb": "Freistellungen beantragen und Status verfolgen (Schüler/KV).",
          "setup": false
        },
        {
          "id": "lehrer-freistellung-planer",
          "label": "Freistellungen Lehrkräfte",
          "blurb": "Freistellungen für Lehrkräfte – Direktion genehmigt, Kalender und iCal.",
          "setup": false
        },
        {
          "id": "freistellung-setup",
          "label": "Freistellungen Setup",
          "blurb": "",
          "setup": true
        },
        {
          "id": "lehrer-freistellung-setup",
          "label": "Lehrer-Freistellungen Setup",
          "blurb": "",
          "setup": true
        },
        {
          "id": "playbook-freistellungen",
          "label": "Playbook Freistellungen",
          "blurb": "",
          "setup": true
        }
      ],
      "searchText": "freistellungen schüler- und lehrer-freistellungen inkl. setup. freistellungen (schüler) freistellungen lehrkräfte freistellungen setup lehrer-freistellungen setup playbook freistellungen freistellung planer lehrer freistellung planer freistellung setup lehrer freistellung setup playbook freistellungen freistellung"
    },
    {
      "id": "automationen",
      "label": "Automationen",
      "priceEuro": 130,
      "blurb": "Power-Automate-Rezepte und Flows.",
      "requires": [],
      "highlights": [
        "Automationen-Übersicht"
      ],
      "tools": [
        {
          "id": "pa-erst-setup",
          "label": "Power Platform Erst-Setup",
          "blurb": "",
          "setup": true
        },
        {
          "id": "power-automate-rezepte",
          "label": "Automationen-Übersicht",
          "blurb": "",
          "setup": false
        },
        {
          "id": "pa-termine-sync",
          "label": "Termine → Kalender",
          "blurb": "",
          "setup": true
        },
        {
          "id": "pa-schularbeiten-mail",
          "label": "Schularbeiten-Mail Flow",
          "blurb": "",
          "setup": true
        },
        {
          "id": "pa-projektwochen-mail",
          "label": "Projektwochen-Mail Flow",
          "blurb": "",
          "setup": true
        },
        {
          "id": "pa-antraege",
          "label": "Forms-Anträge Flow",
          "blurb": "",
          "setup": true
        },
        {
          "id": "pa-seminar",
          "label": "Seminar Flow",
          "blurb": "",
          "setup": true
        },
        {
          "id": "pa-schilf",
          "label": "Schilf-Checkliste",
          "blurb": "",
          "setup": true
        },
        {
          "id": "srdp-mail-setup",
          "label": "sRDP Bestätigungsmail",
          "blurb": "",
          "setup": true
        }
      ],
      "searchText": "automationen power-automate-rezepte und flows. power platform erst-setup automationen-übersicht termine → kalender schularbeiten-mail flow projektwochen-mail flow forms-anträge flow seminar flow schilf-checkliste srdp bestätigungsmail pa erst setup power automate rezepte pa termine sync pa schularbeiten mail pa projektwochen mail pa antraege pa seminar pa schilf srdp mail setup"
    },
    {
      "id": "stammdaten-portal",
      "label": "Stammdaten-Portal",
      "priceEuro": 390,
      "blurb": "Öffentliches Schüler-Stammdaten-/Einschreibeportal.",
      "requires": [
        "intranet"
      ],
      "highlights": [
        "Schüler-Stammdaten-Portal"
      ],
      "tools": [
        {
          "id": "schueler-stammdaten-portal",
          "label": "Schüler-Stammdaten-Portal",
          "blurb": "",
          "setup": false
        }
      ],
      "searchText": "stammdaten-portal öffentliches schüler-stammdaten-/einschreibeportal. schüler-stammdaten-portal schueler stammdaten portal"
    },
    {
      "id": "ablauf",
      "label": "Eltern & Ablauf",
      "priceEuro": 290,
      "blurb": "Einsammler, Elternbrief, Einwilligungen, Lesebestätigungen, Krankmeldung, Kollegiumsumfragen.",
      "requires": [
        "intranet"
      ],
      "highlights": [
        "Elternbrief",
        "Einsammler",
        "Krankmeldung",
        "Kollegiumsumfragen",
        "Teams-Antrag"
      ],
      "tools": [
        {
          "id": "einsammler",
          "label": "Einsammler",
          "blurb": "Einwilligungen und Rückläufe digital einsammeln.",
          "setup": false
        },
        {
          "id": "sharepoint-liste-einsammler",
          "label": "Einsammler-Listen",
          "blurb": "",
          "setup": true
        },
        {
          "id": "elternbrief",
          "label": "Elternbrief",
          "blurb": "Briefe an Eltern mit Lesebestätigung.",
          "setup": false
        },
        {
          "id": "sharepoint-liste-elternbrief",
          "label": "Elternbrief-Listen",
          "blurb": "",
          "setup": true
        },
        {
          "id": "krankmeldung",
          "label": "Krankmeldung",
          "blurb": "Krankmeldungen der Klasse einsehen und bearbeiten.",
          "setup": false
        },
        {
          "id": "sharepoint-liste-krankmeldung",
          "label": "Krankmeldungen-Liste",
          "blurb": "",
          "setup": true
        },
        {
          "id": "kollegium-umfragen",
          "label": "Kollegiumsumfragen",
          "blurb": "Umfragen nur für Lehrkräfte – Antworten und Auswertung.",
          "setup": false
        },
        {
          "id": "sharepoint-liste-kollegium-umfragen",
          "label": "Kollegiumsumfragen-Listen",
          "blurb": "",
          "setup": true
        },
        {
          "id": "teams-antrag",
          "label": "Teams-Antrag",
          "blurb": "Neues Team beantragen – IT prüft und legt es an.",
          "setup": false
        },
        {
          "id": "sharepoint-liste-teams-antrag",
          "label": "Teams-Antrag-Liste",
          "blurb": "",
          "setup": true
        }
      ],
      "searchText": "eltern & ablauf einsammler, elternbrief, einwilligungen, lesebestätigungen, krankmeldung, kollegiumsumfragen. einsammler einsammler-listen elternbrief elternbrief-listen krankmeldung krankmeldungen-liste kollegiumsumfragen kollegiumsumfragen-listen teams-antrag teams-antrag-liste einsammler sharepoint liste einsammler elternbrief sharepoint liste elternbrief krankmeldung sharepoint liste krankmeldung kollegium umfragen sharepoint liste kollegium umfragen teams antrag sharepoint liste teams antrag eltern"
    },
    {
      "id": "ausgaben",
      "label": "Ausgaben & Zugang",
      "priceEuro": 150,
      "blurb": "Schlüssel, Raum-/Schrankenkarten, Parkplatz und eigene Inventar-Typen.",
      "requires": [
        "intranet"
      ],
      "highlights": [
        "Ausgabenverwaltung"
      ],
      "tools": [
        {
          "id": "ausgaben",
          "label": "Ausgabenverwaltung",
          "blurb": "Schlüssel, Karten und Zugänge verwalten.",
          "setup": false
        },
        {
          "id": "sharepoint-liste-ausgaben",
          "label": "Ausgaben-Listen",
          "blurb": "",
          "setup": true
        }
      ],
      "searchText": "ausgaben & zugang schlüssel, raum-/schrankenkarten, parkplatz und eigene inventar-typen. ausgabenverwaltung ausgaben-listen ausgaben sharepoint liste ausgaben ausgabe schlüssel"
    },
    {
      "id": "betrieb",
      "label": "Betrieb & Portal",
      "priceEuro": 130,
      "blurb": "Kalender und Termine, Aufgaben, Technik-Meldungen und Raumplan für den Schulalltag.",
      "requires": [
        "intranet"
      ],
      "highlights": [
        "Kalender und Termine",
        "Aufgaben",
        "Technik-Meldungen",
        "Raumplan"
      ],
      "tools": [
        {
          "id": "termine-hub",
          "label": "Kalender und Termine",
          "blurb": "Schul-, Lehrer- und Schülerkalender aus Outlook – rollenbasiert, Woche/Monat, anlegen und pflegen.",
          "setup": false
        },
        {
          "id": "aufgaben",
          "label": "Aufgaben",
          "blurb": "Aufgaben an Personen oder Rollen zuweisen und abhaken.",
          "setup": false
        },
        {
          "id": "sharepoint-liste-aufgaben",
          "label": "Aufgaben-Liste",
          "blurb": "",
          "setup": true
        },
        {
          "id": "technik-meldungen",
          "label": "Technik-Meldungen",
          "blurb": "Defekte oder problematische Technik melden – Lehrkräfte und Schüler; IT bearbeitet die Inbox.",
          "setup": false
        },
        {
          "id": "sharepoint-liste-technik-meldungen",
          "label": "Technik-Meldungen-Liste",
          "blurb": "",
          "setup": true
        },
        {
          "id": "raumplan",
          "label": "Raumplan",
          "blurb": "Räume suchen, Plan ansehen, CSV importieren.",
          "setup": false
        },
        {
          "id": "sharepoint-liste-raumplan",
          "label": "Raumplan-Liste",
          "blurb": "",
          "setup": true
        }
      ],
      "searchText": "betrieb & portal kalender und termine, aufgaben, technik-meldungen und raumplan für den schulalltag. kalender und termine aufgaben aufgaben-liste technik-meldungen technik-meldungen-liste raumplan raumplan-liste termine hub aufgaben sharepoint liste aufgaben technik meldungen sharepoint liste technik meldungen raumplan sharepoint liste raumplan technik raum"
    }
  ],
  "addons": [
    {
      "id": "addon-setup",
      "label": "Setup / Onboarding",
      "priceEuro": null,
      "priceLabel": "590–790 € einmalig",
      "blurb": "Gemeinsam einrichten: Intranet, Listen, Pakete, Audience.",
      "once": true,
      "packageIds": []
    }
  ],
  "searchAliases": {
    "schulwahl": [
      "schulsprecherwahl"
    ],
    "schulwahlen": [
      "schulsprecherwahl"
    ],
    "wahl": [
      "schulsprecherwahl",
      "wahlfach-zuteilung",
      "wahlfach-anmeldung"
    ],
    "schulsprecher": [
      "schulsprecherwahl"
    ],
    "matura": [
      "srdp-anmeldung",
      "diplomarbeiten"
    ],
    "srdp": [
      "srdp-anmeldung"
    ],
    "eltern": [
      "elternbrief",
      "einsammler",
      "krankmeldung",
      "elternsprechtag-bookings"
    ],
    "freistellung": [
      "freistellung-planer",
      "lehrer-freistellung-planer"
    ],
    "kursteam": [
      "kursteams"
    ],
    "unterrichtsteam": [
      "kursteams"
    ],
    "onenote": [
      "onenote-verteilung"
    ],
    "ausgabe": [
      "ausgaben"
    ],
    "schlüssel": [
      "ausgaben"
    ],
    "technik": [
      "technik-meldungen"
    ],
    "raum": [
      "raumplan",
      "raeume-ressourcen"
    ]
  }
};
