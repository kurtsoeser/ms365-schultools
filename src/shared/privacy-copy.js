/**
 * Kanonische Datenschutz-/Datenlage-Texte (Single Source of Truth).
 *
 * Genutzt von App-Hilfe und Landing über:
 *   node scripts/sync-privacy-copy.mjs
 *
 * Inhaltlich absichtlich ehrlich: sensible Schul-/Schülerdaten liegen im
 * Schul-Tenant bzw. lokal im Browser; der Betreiber betreibt nur schlanke
 * Minimaldienste (Lizenz, Hosting, optionale Portale) – keine Klassenlisten-DB.
 */

/** @typedef {{ title: string, paragraphs: string[], list?: string[] }} PrivacyBlock */

export const PRIVACY_COPY = Object.freeze({
    /** Sichtbarer Stand auf Landing/Hilfe */
    standLabel: 'November 2026',

    metaDescription:
        'Datenschutz MS365-Schul-Tools: sensible Schul- und Schülerdaten im Microsoft-365-Tenant Ihrer Schule; Betreiber speichert nur Minimaldaten für Lizenz und optionale Portale.',

    /**
     * Kurzfassung Landing-Sektion #datenschutz
     * (HTML in lead erlaubt – Sync escaped nichts zusätzlich).
     */
    landingTeaser: Object.freeze({
        eyebrow: 'Datenschutz &amp; Technik',
        title: 'Sensible Schuldaten bleiben in Ihrem Tenant.',
        lead:
            'Klassen-, Schüler- und Fachdaten liegen im <strong>Microsoft-365-Tenant Ihrer Schule</strong> (SharePoint, Teams, Graph) bzw. lokal im Browser der bedienenden Person – nicht als Produkt-Datenbank beim Betreiber. Schlanke Betreiber-Dienste (Lizenz, statisches App-Hosting, optionale Portale) speichern nur das Minimum, das Freischaltung und Betrieb brauchen.',
        points: Object.freeze([
            'Sensible Schul- und Schülerdaten: Schul-Tenant (SharePoint/Teams) bzw. lokaler Browser',
            'Kein Anwendungsserver für Ihre Klassen-/Schülerlisten als Betreiber-Datenbank',
            'Betreiber-Minimaldaten: Lizenz, Kontakte, Site-URL-Zeiger – keine Stammlisten',
            'Anmeldung über Ihre Entra-ID / Microsoft&nbsp;365'
        ])
    }),

    /** Hilfe: Schnellstart-Lead (ein Absatz, HTML erlaubt) */
    helpSchnellstartLead:
        'Die Oberfläche läuft im Browser: Stammdaten lokal, Fachdaten und Listen in <strong>SharePoint Ihres Schul-Tenants</strong>. Online-Aktionen gehen über Microsoft Graph mit Ihrem Schulkonto. Der Betreiber betreibt schlanke Minimaldienste (Lizenz, Hosting, optionale Portale) – <strong>keinen</strong> Anwendungsserver, der Ihre Klassen- oder Schülerlisten als Produkt-Datenbank speichert. Optional sichern Sie den lokalen Stand in die IT-Bibliothek auf Ihrer schulischen SharePoint-Site.',

    /** Hilfe: Hinweisblock #datenspeicher (HTML erlaubt) */
    helpDatenhaltungHint:
        '<strong>Datenhaltung:</strong> Sensible Schul- und Schülerdaten liegen im Browser (<code>localStorage</code> / <code>sessionStorage</code>) und – sobald Sie Online-Werkzeuge nutzen – in <strong>SharePoint / Microsoft&nbsp;365 Ihrer Schule</strong>. Es gibt <strong>keinen</strong> Betreiber-Anwendungsserver für Ihre Stammlisten. Schlanke Betreiber-Dienste prüfen Lizenzen bzw. bedienen optionale Portale; dort landen nur Minimaldaten (Tenant-ID, Freischaltung, Kontakte, Site-URL-Zeiger – keine Klassenlisten). Über <strong>Browser-Backup</strong> (Dashboard oder Stammdaten) übertragen Sie lokale App-Daten als JSON. Zusätzlich können Sie denselben Stand in eine <strong>IT-Sicherungsbibliothek auf der schulischen SharePoint-Site</strong> schreiben (<strong>Nach SharePoint</strong> / optionaler Auto-Sync; siehe <a href="#tool-stammdaten-uebergabe">Stammdaten-Übergabe</a>). Die Backup-Datei liegt im Mandanten Ihrer Schule, nicht beim Betreiber. Die Microsoft-Anmeldung ist in beiden Backup-Wegen nicht enthalten und muss neu erfolgen. Einzelne Werkzeuge haben zusätzlich eigenen Export/Import (JSON/CSV).',

    /** Hilfe: Glossar „Stammdaten“ */
    helpGlossaryStammdaten:
        'Die Listen Ihrer Schule in dieser App: Domain, Fächer, Lehrkräfte, Klassen, Schülerinnen. Arbeitsstand im Browser; Fach- und Planerdaten in SharePoint Ihres Schul-Tenants – nicht als Produkt-Datenbank beim Betreiber.',

    /** README-Kurzabsatz (Markdown) */
    readmeBlurb:
        'Sensible Schul- und Schülerdaten liegen im **Microsoft-365-Tenant der Schule** (SharePoint/Teams/Graph) bzw. lokal im Browser (`localStorage`) – **nicht** als Klassenlisten-Datenbank beim Betreiber. Schlanke Betreiber-Dienste (Lizenz-API, statisches Hosting, optionale Portale) speichern nur Minimaldaten für Freischaltung und Betrieb. Über **Browser-Backup** lassen sich lokale App-Daten als JSON übertragen; optional in die **IT-Sicherungsbibliothek auf der schulischen SharePoint-Site**. Online-Aktionen gehen – nur wenn Sie sie auslösen – an **Microsoft Graph**. Details: [hilfe.html#datenschutz](hilfe.html#datenschutz), ausführlich: [landing/datenschutz.html](landing/datenschutz.html).',

    /**
     * Gemeinsamer Kern (Hilfe + Landing).
     * `paragraphs` / `list` enthalten fertiges HTML (Links, <strong>, <code>).
     */
    core: Object.freeze({
        intro: Object.freeze({
            paragraphs: Object.freeze([
                'Diese Erklärung beschreibt die <strong>Browser-App MS365-Schul-Tools</strong> (statische HTML/JS-Oberfläche) und – auf ms365.schule – die Marketing-Website. Sie ergänzt die Datenschutzhinweise Ihrer Schule und von Microsoft&nbsp;365, ersetzt sie aber nicht.'
            ])
        }),

        schoolResponsibility: Object.freeze({
            titleHelp: 'Verantwortlichkeit',
            titleLanding: 'Verantwortlichkeit für Schuldaten in Microsoft&nbsp;365',
            paragraphs: Object.freeze([
                'Für personenbezogene Daten im Microsoft-365-Mandanten der Schule (Entra&nbsp;ID, Teams, SharePoint, Exchange usw.) ist die <strong>Schule bzw. der Schulerhalter</strong> Verantwortliche im Sinne der DSGVO. Microsoft ist typischerweise Auftragsverarbeiter für Microsoft&nbsp;365.',
                'Sensible Schul- und Schülerdaten werden <strong>nicht</strong> als Produkt-Datenbank beim Software-Betreiber geführt. Arbeitsstand und Stammdaten-Listen liegen im Browser der bedienenden Person; Fachdaten, Planer, Freistellungen, Elternkanal und Betrieb speichern in <strong>SharePoint bzw. Microsoft&nbsp;365 Ihrer Schule</strong>, sobald Sie die entsprechenden Werkzeuge nutzen.',
                'Der Betreiber betreibt bewusst nur <strong>schlanke Minimaldienste</strong> (Lizenzfreischaltung, Auslieferung der App-Dateien, ausgewählte optionale Portale). Dort landet nur das, was für Freischaltung und Betrieb nötig ist – nicht Ihr gesamter Schülerstamm.'
            ])
        }),

        noProductDb: Object.freeze({
            titleHelp: 'Kein Anwendungsserver für Schul-Listen, keine App-Analytik',
            titleLanding: 'Kein Anwendungsserver für Schul-Listen, keine App-Analytik',
            paragraphs: Object.freeze([
                'Importierte Klassen-, Schüler- und Lehrerlisten werden <strong>in diesem Browser</strong> verarbeitet und – bei Online-Aktionen – in Ihrem Microsoft-365-Tenant (Graph / SharePoint) gespeichert. Es gibt <strong>keinen Anwendungsserver des Betreibers für Ihre Schul-Listen</strong>: keinen zentralen „Schuldaten-Server“, der Ihren gesamten Stamm dauerhaft als Produkt-Datenbank führt.',
                'Website und App-Dateien werden typischerweise von einem <strong>statischen Host</strong> ausgeliefert (Webspace / CDN). Das Hosting kann übliche Zugriffsprotokolle führen (IP-Adresse, Zeitpunkt, angeforderte Datei) – nicht jedoch Ihre Stammdaten-Listen. Von dieser App werden keine Tracking-Pixel, Werbenetze oder Analyse-Dienste eingebunden.'
            ])
        }),

        localData: Object.freeze({
            titleHelp: 'Welche Daten lokal liegen',
            titleLanding: 'Welche Daten lokal in der App liegen',
            lead:
                'Im Speicher dieses Browser-Ursprungs (<code>localStorage</code> / <code>sessionStorage</code>, Keys mit Präfix <code>ms365-</code> bzw. <code>webuntis-</code>) können u. a. liegen:',
            list: Object.freeze([
                'Schul-Domain, Schuljahr, Fächer, ARGE, Lehrkräfte, Verwaltung, Schüler:innen, Klassen',
                'Zuordnungen zu Microsoft-365-Gruppen (IDs, Anzeigenamen)',
                'Zwischenstände einzelner Werkzeuge, Dashboard-Favoriten und Darstellungseinstellungen',
                'optional Erziehungsberechtigte (Name/E-Mail) für den Eltern-Verteiler'
            ]),
            after:
                'Die Microsoft-Anmeldung (MSAL) legt Zugriffstoken getrennt ab; sie sind <strong>nicht</strong> Teil des Browser-Backups und auch nicht Teil der SharePoint-Sicherungsdatei. Ein optionaler PIN-/Zugangsschutz speichert nur den Session-Status „Zugang gewährt“, nicht die Stammdaten verschlüsselt.'
        }),

        backup: Object.freeze({
            titleHelp: 'Sicherung: Browser-Backup und schuleigene SharePoint-Bibliothek',
            titleLanding: 'Sicherung: Browser-Backup und schuleigene SharePoint-Bibliothek',
            paragraphs: Object.freeze([
                '<strong>Browser-Backup:</strong> Export als JSON-Datei auf dem Gerät (Dashboard, Stammdaten, Einrichtung). Damit übertragen Sie den lokalen Stand zwischen Browsern oder PCs. Die Datei liegt in Ihrer Verantwortung.',
                '<strong>SharePoint (optional):</strong> Nach Einrichtung der IT-Sicherungsbibliothek schreibt die App denselben vollständigen Stand als JSON in eine Dokumentbibliothek auf der <strong>Intranet-/SharePoint-Site Ihrer Schule</strong> (typisch <code>MS365-IT-Stammdaten</code>, Datei unter <code>Backups/…</code>). Rechte werden auf IT/Verwaltung beschränkt. Die Datei liegt im Mandanten der Schule – nicht beim App-Betreiber. Schüler- und Elternstammdaten werden bewusst <strong>nicht</strong> als SharePoint-Listenzeilen abgelegt, sondern nur in dieser geschützten Backup-Datei (wie im lokalen Speicher).'
            ])
        }),

        microsoft: Object.freeze({
            titleHelp: 'Welche Daten zu Microsoft gehen',
            titleLanding: 'Welche Daten zu Microsoft gehen',
            paragraphs: Object.freeze([
                'Nur wenn Sie sich anmelden und eine Online-Aktion auslösen (einlesen, anlegen, ändern, archivieren, nach SharePoint sichern bzw. von dort laden), sendet der Browser Anfragen an Microsoft Login und Microsoft Graph. Inhalt sind die von Ihnen angestoßenen Vorgänge (z. B. Gruppennamen, Mitglieds-UPNs, Lizenz-SKUs, Backup-Datei in Ihrer Site). Ohne solche Aktionen bleiben die Listen im Browser.',
                'PowerShell-Skripte und CSV-/JSON-Downloads verlassen den Browser als lokale Dateien. Sobald Sie sie ausführen oder weitergeben, gelten die Regeln Ihrer IT und der DSGVO für Dateien mit Personenbezug.'
            ])
        }),

        operatorServices: Object.freeze({
            titleHelp: 'Betreiber-Minimaldienste (Lizenz, Portale)',
            titleLanding: 'Betreiber-Dienste (Lizenz, Portale, optionale APIs)',
            lead:
                'Je nach freigeschaltetem Modul kann die App nach der Microsoft-Anmeldung mit vom Betreiber betriebenen Azure Functions sprechen. Zweck ist Betrieb und Freischaltung – <strong>Minimalprinzip</strong>: nur speichern, was dafür nötig ist.',
            list: Object.freeze([
                '<strong>Lizenz-API:</strong> Tenant-ID, technische Freischaltung (Module/Pakete), optional Schulname/Kontaktdaten/Rollen aus der Betreiber-Lizenzliste, optional URL-Zeiger auf die Schul-SharePoint-Site (<code>ToolSiteUrl</code>), optional Feedback-Texte. <strong>Keine</strong> Klassen-, Schüler- oder Planerlisten.',
                '<strong>Statisches App-/Website-Hosting:</strong> Auslieferung der HTML/JS/CSS-Dateien; höchstens übliche Zugriffsprotokolle des Hosts – keine Stammdaten-Listen.',
                '<strong>Schulsprecherwahl (optional):</strong> Wahldaten in Azure Tables – bewusst nicht als SharePoint-Listenzeilen, damit Graph-Autorenfelder die Anonymität der Stimmabgabe nicht unterlaufen.',
                '<strong>Eltern-/Ablauf-Links (optional, z. B. Einsammler):</strong> Remote-API für Magic-Link-Antworten; Fachdaten der Vorgänge liegen primär in SharePoint Ihrer Schule bzw. lokal im Absender-Browser.',
                '<strong>Schüler-Stammdaten-Portal (optional):</strong> Gateway für Einschreibe-/Elternformulare; Ablage typisch in SharePoint im Schul-Tenant. Zugang über Tenant-Schlüssel / Session, nicht über Microsoft-Login der Eltern.'
            ]),
            after:
                'Ohne Anmeldung bzw. ohne Nutzung dieser Module findet keine solche Übertragung statt. Details zu Listen und Sync-Pfaden stehen in der App-Hilfe und den Tool-Dokumenten.'
        }),

        retention: Object.freeze({
            titleHelp: 'Speicherdauer',
            titleLanding: 'Speicherdauer',
            list: Object.freeze([
                '<strong>Lokal im Browser:</strong> bis Sie Daten löschen, ein Backup überschreibt oder der Browser den Speicher leert',
                '<strong>SharePoint-Backup / Fachlisten im Schul-Tenant:</strong> bis IT löscht oder überschreibt – im Rahmen der Aufbewahrung Ihrer Schule',
                '<strong>Host-Zugriffsprotokolle / E-Mail-Anfragen:</strong> nach betrieblicher Notwendigkeit bzw. gesetzlichen Fristen',
                '<strong>Betreiber-Minimaldienste (Lizenz, Wahl, optionale Portale):</strong> so lange wie für Freischaltung, Support und den jeweiligen Schulprozess nötig; Wahldaten nach Ende des Wahlzyklus bzw. nach Löschung durch die Schule/IT'
            ])
        }),

        rightsShared: Object.freeze({
            titleHelp: 'Betroffenenrechte',
            titleLanding: 'Betroffenenrechte',
            paragraphs: Object.freeze([
                'Auskunft, Berichtigung und Löschung zu Konten und Gruppen richten sich an die Schule (IT / Datenschutzbeauftragte) und werden in Microsoft&nbsp;365 bzw. den Schulverwaltungsquellen umgesetzt.',
                'Das lokale Browser-Backup, die SharePoint-Sicherungsdatei und Exports (CSV/JSON) sind Kopien: nach einer Korrektur im Mandanten veraltete Dateien löschen, neu sichern oder erneut erzeugen.'
            ])
        }),

        tips: Object.freeze({
            titleHelp: 'Empfehlungen für den Schulbetrieb',
            titleLanding: 'Empfehlungen für den Schulbetrieb',
            list: Object.freeze([
                'App möglichst nur auf dienstlichen Geräten bzw. dienstlichen Browserprofilen nutzen',
                'Browser-Backups, SharePoint-Sicherungsdatei und Skript-Downloads wie andere Dateien mit Schüler- und Elterndaten behandeln',
                'IT-Sicherungsbibliothek nur für berechtigte IT-/Verwaltungsgruppen freigeben',
                'Eltern nicht als Gäste anlegen, wenn ein Verteiler ausreicht',
                'Gast-Exporte und Personen-CSVs nur speichern, wenn Ihre Organisation das erlaubt',
                'Konto-Löschen nur nach dem schulischen Offboarding und mit eindeutiger Bestätigung'
            ])
        }),

        summary: Object.freeze({
            titleHelp: null,
            titleLanding: 'Kurzfassung',
            paragraphs: Object.freeze([
                'Sensible Schul- und Schülerdaten bleiben in Ihrem Schul-Tenant bzw. lokal im Browser. Der Betreiber führt keine zentrale Produkt-Datenbank Ihrer Klassenlisten. Schlanke Betreiber-Dienste dienen Lizenz, Hosting und optionalen Portalen – nach Minimalprinzip, ohne Ihren gesamten Stamm zu spiegeln. Was in Microsoft&nbsp;365 landet, landet dort, weil Sie (oder freigeschaltete Automationen) eine Aktion ausgelöst haben.'
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
