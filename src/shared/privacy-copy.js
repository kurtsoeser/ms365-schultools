/**
 * Kanonische Datenschutz-/Datenlage-Texte (Single Source of Truth).
 *
 * Genutzt von App-Hilfe und Landing über:
 *   node scripts/sync-privacy-copy.mjs
 *
 * Nur belegte Verarbeitungen. Optionale Module, die Schülerdaten an
 * Betreiber-Dienste senden können, sind ausdrücklich genannt.
 */

/** @typedef {{ title: string, paragraphs: string[], list?: string[] }} PrivacyBlock */

export const PRIVACY_COPY = Object.freeze({
    /** Sichtbarer Stand auf Landing/Hilfe */
    standLabel: '24. November 2026',

    metaDescription:
        'Datenschutz MS365-Schul-Tools: Schuldaten im Browser und im Microsoft-365-Tenant der Schule. Betreiber speichert Lizenz- und Kontaktdaten; optionale Module können weitere Daten an Azure-Dienste senden.',

    /**
     * Kurzfassung Landing-Sektion #datenschutz
     * (HTML in lead erlaubt – Sync escaped nichts zusätzlich).
     */
    landingTeaser: Object.freeze({
        eyebrow: 'Datenschutz &amp; Technik',
        title: 'Sensible Schuldaten bleiben in Ihrem Tenant.',
        lead:
            'Klassen-, Schüler- und Fachdaten liegen im <strong>Microsoft-365-Tenant Ihrer Schule</strong> (SharePoint, Teams, Graph) bzw. lokal im Browser der bedienenden Person – nicht als Produkt-Datenbank beim Betreiber. Die Lizenzprüfung speichert nur Schul- und Kontaktangaben. <strong>Optionale Module</strong> (Schulsprecherwahl, Einsammler, Krankmeldung, Schüler-Stammdaten-Portal, Kursteams) können weitere Daten an Dienste des Betreibers senden, wenn die Schule sie einschaltet.',
        points: Object.freeze([
            'Sensible Schul- und Schülerdaten: Schul-Tenant (SharePoint/Teams) bzw. lokaler Browser',
            'Kein Anwendungsserver für Ihre Klassen-/Schülerlisten als Betreiber-Datenbank',
            'Lizenzdienst: Schulname, Kontakte, Site-URL-Zeiger – keine Klassenlisten',
            'Anmeldung über die Entra-ID Ihrer Schule. Optionale Module getrennt in der Datenschutzerklärung'
        ])
    }),

    /** Hilfe: Schnellstart-Lead (ein Absatz, HTML erlaubt) */
    helpSchnellstartLead:
        'Die Oberfläche läuft im Browser: Stammdaten lokal, Fachdaten und Listen in <strong>SharePoint Ihres Schul-Tenants</strong>. Online-Aktionen gehen über Microsoft Graph mit Ihrem Schulkonto. Der Betreiber führt <strong>keinen</strong> Anwendungsserver, der Ihre Klassen- oder Schülerlisten als Produkt-Datenbank speichert. Lizenz, Hosting und – nur wenn Sie sie einschalten – einzelne Module (Wahl, Einsammler, Krankmeldung, Schülerportal, Kursteams) sind davon getrennt und in der Datenschutzerklärung beschrieben. Optional sichern Sie den lokalen Stand in die IT-Bibliothek auf Ihrer schulischen SharePoint-Site.',

    /** Hilfe: Hinweisblock #datenspeicher (HTML erlaubt) */
    helpDatenhaltungHint:
        '<strong>Datenhaltung:</strong> Sensible Schul- und Schülerdaten liegen im Browser (<code>localStorage</code> / <code>sessionStorage</code>, Anhänge zusätzlich in <code>IndexedDB</code>) und – sobald Sie Online-Werkzeuge nutzen – in <strong>SharePoint / Microsoft&nbsp;365 Ihrer Schule</strong>. Es gibt <strong>keinen</strong> Betreiber-Anwendungsserver für Ihre Stammlisten. Die Lizenzprüfung und optionale Module sind eigene Dienste; dort liegen keine Klassenlisten, einzelne Module können aber Personen- oder Schülerdaten entgegennehmen, wenn die Schule sie einschaltet. Über <strong>Browser-Backup</strong> (Dashboard oder Stammdaten) übertragen Sie lokale App-Daten als JSON. Zusätzlich können Sie denselben Stand in eine <strong>IT-Sicherungsbibliothek auf der schulischen SharePoint-Site</strong> schreiben (<strong>Nach SharePoint</strong> / optionaler Auto-Sync; siehe <a href="#tool-stammdaten-uebergabe">Stammdaten-Übergabe</a>). Die Backup-Datei liegt im Mandanten Ihrer Schule, nicht beim Betreiber. Die Microsoft-Anmeldung ist in beiden Backup-Wegen nicht enthalten und muss neu erfolgen.',

    /** Hilfe: Glossar „Stammdaten“ */
    helpGlossaryStammdaten:
        'Die Listen Ihrer Schule in dieser App: Domain, Fächer, Lehrkräfte, Klassen, Schülerinnen. Arbeitsstand im Browser; Fach- und Planerdaten in SharePoint Ihres Schul-Tenants – nicht als Produkt-Datenbank beim Betreiber.',

    /** README-Kurzabsatz (Markdown) */
    readmeBlurb:
        'Sensible Schul- und Schülerdaten liegen im **Microsoft-365-Tenant der Schule** (SharePoint/Teams/Graph) bzw. lokal im Browser (`localStorage`, Anhänge in IndexedDB) – **nicht** als Klassenlisten-Datenbank beim Betreiber. Die Lizenz-API speichert Schul- und Kontaktangaben. Optionale Module (Wahl, Einsammler, Krankmeldung, Schülerportal, Kursteams) können weitere Daten an Azure-Dienste des Betreibers senden, wenn die Schule sie einschaltet. Details: [hilfe.html#datenschutz](hilfe.html#datenschutz), ausführlich: [landing/datenschutz.html](landing/datenschutz.html).',

    /**
     * Gemeinsamer Kern (Hilfe + Landing).
     * `paragraphs` / `list` enthalten fertiges HTML (Links, <strong>, <code>).
     */
    core: Object.freeze({
        intro: Object.freeze({
            paragraphs: Object.freeze([
                'Diese Erklärung beschreibt die <strong>Browser-App MS365-Schul-Tools</strong> (statische HTML/JS-Oberfläche unter app.ms365.schule) und die Website ms365.schule. Sie ergänzt die Datenschutzhinweise Ihrer Schule und von Microsoft&nbsp;365, ersetzt sie aber nicht.'
            ])
        }),

        schoolResponsibility: Object.freeze({
            titleHelp: 'Verantwortlichkeit',
            titleLanding: 'Verantwortlichkeit für Schuldaten in Microsoft&nbsp;365',
            paragraphs: Object.freeze([
                'Für personenbezogene Daten im Microsoft-365-Mandanten der Schule (Entra&nbsp;ID, Teams, SharePoint, Exchange usw.) und für die Stammdaten im Browser der Schule ist die <strong>Schule bzw. der Schulerhalter</strong> Verantwortliche im Sinne der DSGVO. Microsoft ist für Microsoft&nbsp;365 Auftragsverarbeiter der Schule, nicht des Betreibers dieser Website.',
                'Die Schule bleibt zuständig für Rechtsgrundlage, Information der betroffenen Personen und – wo sie nötig ist – Einwilligung, auch bei Schülerinnen, Schülern und anderen minderjährigen Personen. Diese Website holt keine Einwilligung für Schuldaten ein.',
                'Der Betreiber ist Verantwortlicher für den Betrieb der Website, für E-Mail-Anfragen an ihn und für die Lizenzliste. Schülerdaten speichert er nicht als eigene Datenbank. Das Schülerportal schreibt in eine SharePoint-Liste der Schule. Eine Auftragsverarbeitung wird über diese Website nicht abgeschlossen. Der Anbieter klärt das mit der Schule individuell und bietet es dort an.'
            ])
        }),

        noProductDb: Object.freeze({
            titleHelp: 'Kein Anwendungsserver für Schul-Listen, keine App-Analytik',
            titleLanding: 'Kein Anwendungsserver für Schul-Listen, keine App-Analytik',
            paragraphs: Object.freeze([
                'Importierte Klassen-, Schüler- und Lehrerlisten werden <strong>in diesem Browser</strong> verarbeitet und – bei Online-Aktionen – in Ihrem Microsoft-365-Tenant (Graph / SharePoint) gespeichert. Es gibt <strong>keinen Anwendungsserver des Betreibers für Ihre Schul-Listen</strong>: keinen zentralen „Schuldaten-Server“, der Ihren gesamten Stamm dauerhaft als Produkt-Datenbank führt.',
                'Website und App-Dateien werden als statische Dateien ausgeliefert. Das Hosting kann übliche Zugriffsprotokolle führen (IP-Adresse, Zeitpunkt, angeforderte Datei) – nicht jedoch Ihre Stammdaten-Listen. Von dieser Website und von der App werden keine Tracking-Pixel, Werbenetze oder Analyse-Dienste eingebunden. Ein Web-Analytics-Skript wird nicht geladen.'
            ])
        }),

        localData: Object.freeze({
            titleHelp: 'Welche Daten lokal liegen',
            titleLanding: 'Welche Daten lokal in der App liegen',
            lead:
                'Im Speicher dieses Browser-Ursprungs (<code>localStorage</code> / <code>sessionStorage</code>, Keys mit Präfix <code>ms365-</code> bzw. <code>webuntis-</code>; Anhänge zusätzlich in <code>IndexedDB</code>) können u. a. liegen:',
            list: Object.freeze([
                'Schul-Domain, Schuljahr, Fächer, ARGE, Lehrkräfte, Verwaltung, Schüler:innen, Klassen',
                'Zuordnungen zu Microsoft-365-Gruppen (IDs, Anzeigenamen)',
                'Zwischenstände einzelner Werkzeuge, Dashboard-Favoriten und Darstellungseinstellungen',
                'optional Erziehungsberechtigte (Name/E-Mail) für den Eltern-Verteiler',
                'Anhänge einzelner Werkzeuge in IndexedDB'
            ]),
            after:
                'Die Microsoft-Anmeldung (MSAL) legt Zugriffstoken in <code>localStorage</code> ab und kann zusätzlich ein technisch nötiges Cookie für den Anmeldezustand setzen (<code>storeAuthStateInCookie</code>). Token und dieses Cookie sind <strong>kein</strong> Analyse-Cookie, sie sind <strong>nicht</strong> Teil des Browser-Backups und auch nicht Teil der SharePoint-Sicherungsdatei. Ein optionaler PIN-/Zugangsschutz speichert nur den Session-Status „Zugang gewährt“, nicht die Stammdaten verschlüsselt. Die Marketing-Seiten setzen in ihrem eigenen Skript kein Cookie.'
        }),

        backup: Object.freeze({
            titleHelp: 'Sicherung: Browser-Backup und schuleigene SharePoint-Bibliothek',
            titleLanding: 'Sicherung: Browser-Backup und schuleigene SharePoint-Bibliothek',
            paragraphs: Object.freeze([
                '<strong>Browser-Backup:</strong> Export als JSON-Datei auf dem Gerät (Dashboard, Stammdaten, Einrichtung). Damit übertragen Sie den lokalen Stand zwischen Browsern oder PCs. Die Datei liegt in Ihrer Verantwortung und kann Schüler- und Elterndaten enthalten.',
                '<strong>SharePoint (optional):</strong> Nach Einrichtung der IT-Sicherungsbibliothek schreibt die App denselben vollständigen Stand als JSON in eine Dokumentbibliothek auf der <strong>Intranet-/SharePoint-Site Ihrer Schule</strong> (typisch <code>MS365-IT-Stammdaten</code>, Datei unter <code>Backups/…</code>). Rechte werden auf IT/Verwaltung beschränkt. Die Datei liegt im Mandanten der Schule – nicht beim App-Betreiber. Schüler- und Elternstammdaten werden bewusst <strong>nicht</strong> als SharePoint-Listenzeilen abgelegt, sondern nur in dieser geschützten Backup-Datei (wie im lokalen Speicher).'
            ])
        }),

        microsoft: Object.freeze({
            titleHelp: 'Welche Daten zu Microsoft gehen',
            titleLanding: 'Welche Daten zu Microsoft gehen',
            paragraphs: Object.freeze([
                'Nur wenn Sie sich anmelden und eine Online-Aktion auslösen (einlesen, anlegen, ändern, archivieren, nach SharePoint sichern bzw. von dort laden), sendet der Browser Anfragen an Microsoft Login (<code>login.microsoftonline.com</code>) und Microsoft Graph (<code>graph.microsoft.com</code>). Die Anmeldung läuft mit dem Konto Ihrer Schule in <strong>deren</strong> Mandanten. Inhalt sind die von Ihnen angestoßenen Vorgänge (z. B. Gruppennamen, Mitglieds-UPNs, Lizenz-SKUs, Backup-Datei in Ihrer Site). Ohne solche Aktionen bleiben die Listen im Browser.',
                'PowerShell-Skripte und CSV-/JSON-Downloads verlassen den Browser als lokale Dateien. Sobald Sie sie ausführen oder weitergeben, gelten die Regeln Ihrer IT und der DSGVO für Dateien mit Personenbezug.'
            ])
        }),

        operatorServices: Object.freeze({
            titleHelp: 'Betreiber-Dienste (Lizenz und optionale Module)',
            titleLanding: 'Betreiber-Dienste (Lizenz und optionale Module)',
            lead:
                'Die App kann nach der Microsoft-Anmeldung mit Azure Functions des Betreibers sprechen. Das folgt dem <strong>Minimalprinzip</strong> nur dort, wo es unten so steht. Schülerdaten sind bei den Modulen genannt, die sie entgegennehmen können. Die Rechenzentrumsregion dieser Functions ist im Quellcode nicht festgelegt. Liegt sie außerhalb der EU, ist das ein Drittlandtransfer; eine konkrete Garantie dafür ist in diesem Repository nicht hinterlegt.',
            list: Object.freeze([
                '<strong>Lizenz-API:</strong> Tenant-ID, Freischaltung (Module/Pakete), Schulname, Domains, Adresse der Schule, Namen und E-Mail-Adressen von Direktion, Administration und IT, URL-Zeiger auf die Schul-SharePoint-Site (<code>ToolSiteUrl</code>). Ablage in einer SharePoint-Liste im Microsoft-365-Mandanten des Betreibers. <strong>Keine</strong> Klassen-, Schüler- oder Planerlisten. Rechtsgrundlage: Vertrag bzw. vorvertragliche Schritte und berechtigtes Interesse am Betrieb der Freischaltung, Art.&nbsp;6 Abs.&nbsp;1 lit.&nbsp;b und lit.&nbsp;f DSGVO.',
                '<strong>Feedback über die Lizenz-API:</strong> Titel, Text, Kategorie, Seitenadresse, User-Agent, Name und E-Mail der meldenden Person aus dem Anmeldetoken, Tenant-ID, Schulname, Domain, optional Bildanhänge. Dieselben Anhänge können enthalten, was die meldende Person hochlädt; eine Klassenliste ist dafür nicht vorgesehen und soll nicht angehängt werden. Ablage in der Feedback-Liste im Mandanten des Betreibers. Rechtsgrundlage: Art.&nbsp;6 Abs.&nbsp;1 lit.&nbsp;b und lit.&nbsp;f DSGVO.',
                '<strong>Schulsprecherwahl (nur mit gesetzter API-Adresse):</strong> Azure Table Storage. Die Wahl enthält Kandidatennamen, Klasse, E-Mail und Entra-Objekt-ID. Die Teilnahme wird als Hash gespeichert, der Stimmzettel getrennt davon mit Kandidaten-IDs. Das können Daten von Schülerinnen und Schülern sein, also von Minderjährigen. Zweck: die Stimmabgabe nicht als SharePoint-Zeile mit Graph-Autorenfeld zu speichern. Ohne gesetzte API-Adresse bleibt die Wahl im Browser. Rechtsgrundlage und Information der Betroffenen liegen bei der Schule; der Betreiber verarbeitet diese Daten nur im Auftrag und nur, wenn das Modul eingeschaltet ist.',
                '<strong>Einsammler und Krankmeldung (Remote standardmäßig aus):</strong> Nur wenn die Schule eine API-Adresse setzt (beim Einsammler zusätzlich ein Link-Geheimnis), nimmt eine Azure Function die Antwort entgegen. Gespeichert wird im Arbeitsspeicher der Function, nicht in einer eigenen Datenbank: Schülername, Klasse, Status, Kommentar, optional ein Dateianhang; bei der Krankmeldung zusätzlich Zeitraum, Kategorie, Name und E-Mail der meldenden Person. Krankmeldungen sind Gesundheitsdaten. Ohne diese Einstellung bleiben die Vorgänge in SharePoint Ihrer Schule bzw. im Browser. Die Schule bleibt für Rechtsgrundlage und Information verantwortlich.',
                '<strong>Schüler-Stammdaten-Portal (optional):</strong> Das Formular kann Name, Geburtsdatum, Sozialversicherungsnummer, Adresse, Erziehungsberechtigte und – wenn das Feld aktiv ist – Religionsbekenntnis enthalten. Die Function schreibt diese Angaben in eine SharePoint-Liste im Mandanten der Schule (Site aus der Einrichtung der Schule). Der Betreiber behält davon keine eigene Schülerdatenbank. Ein öffentlicher Statusaufruf meldet nur, ob der Dienst erreichbar ist und ob die Ablage auf SharePoint steht. Er enthält keine Schülerdaten, keine Namen und keine Tenant-ID. Die Schule bleibt für Rechtsgrundlage und Information verantwortlich.',
                '<strong>Kursteams (optional):</strong> Ein Auftragsdatensatz in Azure Blob Storage enthält Tenant-ID, Objekt-ID der anlegenden Person, Teamnamen, Gruppenadressen und Besitzer-UPNs. Die Function legt die Teams danach per Graph im angegebenen Mandanten an. Das ist keine Schülerliste, kann aber personenbezogene Daten von Bediensteten enthalten.'
            ]),
            after:
                'Ohne Anmeldung und ohne Einschalten des jeweiligen Moduls findet diese Übertragung nicht statt. Öffentliche Statusadressen antworten mit einem Betriebsstatus, nicht mit Schülerdaten, Tenant-IDs oder internen Zählern. Senden Sie keine Klassen- oder Schülerlisten an den Betreiber, auch nicht als Feedback-Anhang oder per E-Mail. Eine Auftragsverarbeitung wird über diese Website nicht abgeschlossen. Der Anbieter klärt das mit der Schule individuell und bietet es dort an.'
        }),

        retention: Object.freeze({
            titleHelp: 'Speicherdauer',
            titleLanding: 'Speicherdauer',
            list: Object.freeze([
                '<strong>Lokal im Browser:</strong> bis Sie Daten löschen, ein Backup überschreibt oder der Browser den Speicher leert',
                '<strong>SharePoint-Backup / Fachlisten im Schul-Tenant:</strong> bis die Schule löscht oder überschreibt',
                '<strong>Zugriffsprotokolle des Website-Hosters und E-Mail-Anfragen:</strong> eine feste Frist steht nicht im Quellcode. E-Mails liegen im Postfach, bis sie gelöscht werden. Kommt später ein Vertrag außerhalb dieser Website zustande, können gesetzliche Aufbewahrungsfristen dazukommen',
                '<strong>Lizenz und Feedback:</strong> für die Dauer der Freischaltung und der Bearbeitung, danach bis zur Löschung in der Betreiber-Liste',
                '<strong>Wahl in Azure Tables:</strong> bis die Schule die Wahl löscht. Eine automatische Frist nach dem Wahlgang ist im Code nicht hinterlegt',
                '<strong>Einsammler-/Krankmeldungs-API im Arbeitsspeicher:</strong> nur solange die Function-Instanz den Speicher hält; das ist keine zugesagte Löschfrist',
                '<strong>Schüler-Stammdaten-Portal:</strong> solange der Datensatz im konfigurierten Speicher (Arbeitsspeicher oder SharePoint-Liste) liegt'
            ])
        }),

        rightsShared: Object.freeze({
            titleHelp: 'Betroffenenrechte',
            titleLanding: 'Betroffenenrechte',
            paragraphs: Object.freeze([
                'Auskunft, Berichtigung, Löschung, Einschränkung und – bei einer Einwilligung – Widerruf zu Konten, Listen und Gruppen im Schul-Mandanten richten sich an die Schule (IT / Datenschutzbeauftragte). Die Schule informiert auch Schülerinnen, Schüler und Erziehungsberechtigte.',
                'Gegenüber dem Betreiber haben Sie diese Rechte für die Verarbeitungen, für die er verantwortlich ist: Auskunft, Berichtigung, Löschung, Einschränkung, Datenübertragbarkeit, soweit sie zusteht, und Widerspruch gegen eine Verarbeitung auf berechtigtes Interesse (Art.&nbsp;21 DSGVO). Das lokale Browser-Backup, die SharePoint-Sicherungsdatei und Exports (CSV/JSON) sind Kopien der Schule: nach einer Korrektur veraltete Dateien löschen oder neu erzeugen.'
            ])
        }),

        tips: Object.freeze({
            titleHelp: 'Empfehlungen für den Schulbetrieb',
            titleLanding: 'Empfehlungen für den Schulbetrieb',
            list: Object.freeze([
                'App möglichst nur auf dienstlichen Geräten bzw. dienstlichen Browserprofilen nutzen',
                'Browser-Backups, SharePoint-Sicherungsdatei und Skript-Downloads wie andere Dateien mit Schüler- und Elterndaten behandeln',
                'IT-Sicherungsbibliothek nur für berechtigte IT-/Verwaltungsgruppen freigeben',
                'Keine Klassen- oder Schülerlisten an den Betreiber mailen oder als Feedback-Bild anhängen',
                'Wahl, Einsammler und Schülerportal nur einschalten, wenn die Schule das so will',
                'Eltern nicht als Gäste anlegen, wenn ein Verteiler ausreicht',
                'Konto-Löschen nur nach dem schulischen Offboarding und mit eindeutiger Bestätigung'
            ])
        }),

        summary: Object.freeze({
            titleHelp: null,
            titleLanding: 'Kurzfassung',
            paragraphs: Object.freeze([
                'Sensible Schul- und Schülerdaten bleiben im Regelfall in Ihrem Schul-Tenant bzw. lokal im Browser. Der Betreiber führt keine Klassenlisten-Datenbank. Die Lizenzliste enthält Schul- und Kontaktangaben. Wahl, Einsammler, Krankmeldung, Schülerportal und Kursteams können weitere Daten an Azure-Dienste des Betreibers senden, aber nur wenn die Schule das Modul einschaltet. Was in Microsoft&nbsp;365 Ihrer Schule landet, landet dort, weil jemand mit einem Schulkonto eine Aktion ausgelöst hat.'
            ])
        })
    })
});

/** Phrasen, die in Hilfe/Landing nicht mehr vorkommen dürfen (Widerspruch). */
export const PRIVACY_FORBIDDEN_PHRASES = Object.freeze([
    'Es gibt keinen Anwendungsserver dieser Software, der Listen entgegennimmt',
    'Es gibt keinen App-Server dieser Software',
    'Es gibt keinen App-Server hinter dieser Oberfläche',
    'ohne eigenen Backend-Server für Ihre Stammdaten-Listen'
]);
