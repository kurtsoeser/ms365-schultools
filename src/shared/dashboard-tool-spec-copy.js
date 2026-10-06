/**
 * UI-Spec Abschnitt 5 – Kurzbeschreibungen für Katalog- und Aufgaben-Kacheln.
 * @type {Record<string, string>}
 */
export const DASHBOARD_SPEC_TOOL_DESCRIPTIONS = {
    'slg-schueler':
        'Sammelgruppe für alle Schüler:innen matchen, anlegen und mit Stammliste abgleichen.',
    'slg-lehrer':
        'Sammelgruppe für alle Lehrkräfte matchen, anlegen und mit Stammliste abgleichen.',
    verwaltung:
        'Gruppe für Sekretariat, Direktion und Verwaltungsrollen anlegen und Mitglieder pflegen.',
    klassenvorstaende:
        'Sammelgruppe oder Verteiler aller Klassenvorstände automatisch befüllen.',
    klassenchats: 'Teams-Gruppenchats für alle Lehrkräfte einer Klasse – inkl. Klassenvorstand.',
    'arge-fachgruppen':
        'Eine Gruppe pro Fach oder Arbeitsgemeinschaft – aus den Listen in den Stammdaten.',
    'weitere-teams-gruppen': 'Sonderfälle und zusätzliche Teams oder Gruppen anlegen und pflegen.',
    jahrgang:
        'Klassen aus den Stammdaten mit Microsoft 365-Teams verknüpfen und synchronisieren.',
    kursteams:
        'Teams für Unterrichtsgegenstände aus dem Stundenplan anlegen und befüllen.',
    'kursteam-einzeln':
        'Einzelne Unterrichtsteams manuell anlegen, wenn der Stundenplan-Import nicht greift.',
    'unterrichtsteams-katalog':
        'Alle vorhandenen Unterrichtsteams ansehen, filtern und verwalten.',
    'webuntis-sync-monitor':
        'Stundenplan-Abgleich prüfen: fehlende Besitzer, Lehrer ohne Match, Klassen ohne Team.',
    'klassen-merge':
        'Mehrere Klassengruppen zu einer gemeinsamen Gruppe oder einem Team zusammenführen.',
    'onenote-verteilung':
        'Inhalte aus einem OneNote-Notizbuch an Klassen- oder Kursnotizbücher verteilen.',
    'personen-verwaltung':
        'Personen der Schule suchen, nach Lizenz und Sync filtern, Berichte erstellen.',
    'schueler-lifecycle':
        'Eintritt, Klassenwechsel und Austritt von Schüler:innen geführt durchführen.',
    'gaeste-verwalten':
        'Externe Personen einladen, Gäste in Teams prüfen und Einladungsrechte festlegen.',
    lizenzverwaltung:
        'Sicherheitsgruppen für Lizenz-Pakete verwalten und Zuweisungen prüfen.',
    'namenskonvention-audit':
        'Anzeigenamen und UPN prüfen, Cloud-User korrigieren, AD-synchronisierte exportieren.',
    'playbook-schuljahresstart':
        'Geführte Checkliste für Stammdaten, Klassen, SLG, Kursteams und Cleanup.',
    'organisations-assistent':
        'Schuljahr in den Stammdaten aktivieren, Anzeigenamen anpassen, Abschlussjahrgang prüfen.',
    'klassen-umbenennen':
        'Anzeigenamen der Klassenteams für den Schuljahreswechsel anpassen.',
    'webuntis-stammdaten-import':
        'WebUntis- oder andere Exporte ins Schulregister übernehmen (Review vor dem Speichern).',
    'cleanup-playbook':
        'Geführter Aufräum-Prozess: leere Gruppen, besitzlose Teams, Archivierungen.',
    'sharepoint-intranet-hub':
        'Hub-Site, Listen und Planer – das Intranet für den Schulalltag einrichten und nutzen.',
    'schularbeiten-planer':
        'Schularbeitstermine anlegen, koordinieren und als Kalendereinträge verteilen.',
    'freistellung-planer':
        'Freistellungsanträge stellen, genehmigen und den Status nachverfolgen.',
    'playbook-intranet':
        'Geführte Einrichtung des Schulintranets: Listen, Planer, Berechtigungen, Apps.',
    'schulaktivitaeten-planer':
        'Schulveranstaltungen, Exkursionen und außerschulische Aktivitäten verwalten.',
    projektwochen: 'Projektwochen planen, Gruppen anlegen und Lehrkräfte zuordnen.',
    'sharepoint-liste-stammdaten':
        'SharePoint-Listen mit Stammdaten einsehen und verwalten (IT-Bereich).',
    datenhygiene:
        'Stammlisten vs. Microsoft 365 – letzter Gruppenabgleich mit Konsistent/Abweichung.',
    'leere-gruppen-report':
        'Microsoft 365-Gruppen ohne Mitglieder identifizieren und bereinigen.',
    'teams-archiv': 'Mehrere veraltete Teams auf einmal archivieren – z. B. zum Schuljahreswechsel.',
    'schulstruktur-sync': 'Vollständige Übersicht aller Microsoft 365-Gruppen und Teams der Schule.',
    datenlandkarte:
        'Visualisierung aller Datenquellen, Verbindungen und Sync-Stände auf einen Blick.',
    'stammdaten-backup-abgleich':
        'Gesicherten Stand mit aktuellem Stand vergleichen – Änderungen identifizieren.',
    'playbook-daten-import-verknuepfen':
        'Schritt für Schritt von Export bis Verknüpfung mit Microsoft 365.',
    'stammdaten-quelle-waehlen': 'Datei, Bildungsportal oder anderes Format als Datenquelle festlegen.',
    'bildungsportal-stammdaten': 'Anbindung und Roadmap für Ihr Bundesland-Bildungsportal.',
    'pa-termine-sync':
        'Schultermine zwischen SharePoint-Listen und Outlook-Kalendern synchronisieren – in Entwicklung.',
    'pa-schularbeiten-mail':
        'Automatische E-Mail bei Schularbeitsterminen – in Entwicklung.',
    'pa-projektwochen-mail': 'Status-E-Mails für Projektwochen-Anmeldungen – in Entwicklung.',
    'pa-antraege': 'Microsoft-Forms-Anträge automatisch weiterverarbeiten – in Entwicklung.',
    'pa-seminar': 'Fortbildungs- und Seminar-Anmeldungen per Flow – in Entwicklung.'
};

/**
 * @param {string} toolId
 */
export function specToolDescription(toolId) {
    const id = String(toolId || '').trim();
    if (!id) return '';
    return DASHBOARD_SPEC_TOOL_DESCRIPTIONS[id] || '';
}
