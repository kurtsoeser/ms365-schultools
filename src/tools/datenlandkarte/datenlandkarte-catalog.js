/**
 * Katalog: Datenblöcke (Listen/Quellen) und logische Verbindungen.
 * Rein deklarativ – Zähler kommen aus datenlandkarte-metrics.js
 */

/** @typedef {{
 *   id: string,
 *   title: string,
 *   description: string,
 *   icon: string,
 *   layer: 'stamm'|'schuljahr'|'sharepoint'|'planer'|'m365',
 *   href?: string,
 *   countKey: string
 * }} DatenBlockDef */

/** @typedef {{
 *   id: string,
 *   from: string,
 *   to: string,
 *   label: string,
 *   detail?: string
 * }} DatenLinkDef */

/** @type {DatenBlockDef[]} */
export const DATEN_BLOECKE = [
    {
        id: 'stamm-hub',
        title: 'Stammdaten (Browser)',
        description: 'Zentrale Pflege in Einstellungen – Quelle für fast alle Tools.',
        icon: 'bi-database',
        layer: 'stamm',
        href: '../tenant.html',
        countKey: 'stammHub'
    },
    {
        id: 'stamm-classes',
        title: 'Klassen',
        description: 'Codes, Anzeigenamen, KV-E-Mail, Jahrgang.',
        icon: 'bi-people',
        layer: 'stamm',
        href: '../tenant.html#classes',
        countKey: 'classes'
    },
    {
        id: 'stamm-teachers',
        title: 'Lehrkräfte',
        description: 'Kürzel, Name, E-Mail – Zuordnung zu Unterricht & KV.',
        icon: 'bi-person-badge',
        layer: 'stamm',
        href: '../tenant.html#teachers',
        countKey: 'teachers'
    },
    {
        id: 'stamm-subjects',
        title: 'Fächer',
        description: 'Fachcodes und Bezeichnungen für Planer und Kursteams.',
        icon: 'bi-book',
        layer: 'stamm',
        href: '../tenant.html#subjects',
        countKey: 'subjects'
    },
    {
        id: 'stamm-arges',
        title: 'ARGE / Fachgruppen',
        description: 'Fachgruppen-Stammliste mit zugeordneten Fächern.',
        icon: 'bi-diagram-3',
        layer: 'stamm',
        href: '../tenant.html#subjects',
        countKey: 'arges'
    },
    {
        id: 'year-students',
        title: 'Schüler:innen',
        description: 'Schuljahr-Bucket inkl. Klasse und E-Mail.',
        icon: 'bi-person',
        layer: 'schuljahr',
        href: '../tenant.html#students',
        countKey: 'students'
    },
    {
        id: 'year-guardians',
        title: 'Erziehungsberechtigte',
        description: 'Eltern/EB verknüpft über guardianIds am Schüler.',
        icon: 'bi-person-hearts',
        layer: 'schuljahr',
        href: '../tenant.html#students',
        countKey: 'guardians'
    },
    {
        id: 'year-unterricht',
        title: 'Unterrichtsbelegung',
        description: 'Klasse × Lehrkraft × Fach (aus Kursteam-Endliste).',
        icon: 'bi-calendar2-week',
        layer: 'schuljahr',
        href: '../tools/kursteams.html',
        countKey: 'unterrichtRows'
    },
    {
        id: 'import-webuntis',
        title: 'WebUntis / SIS-Import',
        description: 'Import-Historie und Schüler-Stammdaten aus Untis.',
        icon: 'bi-cloud-upload',
        layer: 'schuljahr',
        href: '../tools/webuntis-stammdaten-import.html',
        countKey: 'sisImports'
    },
    {
        id: 'spo-sp-klassen',
        title: 'SP Liste Klassen',
        description: 'SharePoint „Klassen“ – Sync aus Stammdaten.',
        icon: 'bi-list-ul',
        layer: 'sharepoint',
        href: '../tools/sharepoint-liste-stammdaten.html',
        countKey: 'spoListKlassen'
    },
    {
        id: 'spo-sp-faecher',
        title: 'SP Liste Fächer',
        description: 'SharePoint „Fächer“ – Fachcodes.',
        icon: 'bi-list-ul',
        layer: 'sharepoint',
        href: '../tools/sharepoint-liste-stammdaten.html',
        countKey: 'spoListFaecher'
    },
    {
        id: 'spo-sp-schueler',
        title: 'SP Liste Schülerinnen',
        description: 'SharePoint „Schülerinnen“ – optional mit Personenfeld.',
        icon: 'bi-list-ul',
        layer: 'sharepoint',
        href: '../tools/sharepoint-liste-stammdaten.html',
        countKey: 'spoListSchuelerinnen'
    },
    {
        id: 'spo-lehrer',
        title: 'SP Lehrerliste',
        description: 'Öffentliche Lehrerliste „Lehrerinnen“ auf Intranet.',
        icon: 'bi-journal-text',
        layer: 'sharepoint',
        href: '../tools/sharepoint-liste-lehrer.html',
        countKey: 'spoListLehrerPublic'
    },
    {
        id: 'spo-sap-schularbeiten',
        title: 'SP SAP-Schularbeiten',
        description: 'Listeneinträge Schularbeiten-Planer (Anträge/Termine).',
        icon: 'bi-journal-check',
        layer: 'sharepoint',
        href: '../tools/sharepoint-liste-schularbeiten.html',
        countKey: 'spoListSapSchularbeiten'
    },
    {
        id: 'spo-pw-angebote',
        title: 'SP PW-Angebote',
        description: 'Projektwochen-Angebote auf SharePoint.',
        icon: 'bi-calendar2-heart',
        layer: 'sharepoint',
        href: '../tools/sharepoint-liste-projektwochen.html',
        countKey: 'spoListPwAngebote'
    },
    {
        id: 'spo-freistellungen',
        title: 'SP Freistellungen',
        description: 'Freistellungsanträge (Liste + Approvals-Flow).',
        icon: 'bi-calendar2-check',
        layer: 'sharepoint',
        href: '../tools/freistellung-setup.html',
        countKey: 'spoListFreistellungen'
    },
    {
        id: 'spo-aktivitaeten',
        title: 'SP Schulaktivitäten',
        description: 'Exkursionen / Aktivitäten-Anträge.',
        icon: 'bi-bus-front',
        layer: 'sharepoint',
        href: '../tools/sharepoint-liste-schulaktivitaeten.html',
        countKey: 'spoListAktivitaeten'
    },
    {
        id: 'plan-schularbeiten',
        title: 'Schularbeiten-Planer',
        description: 'App auf SAP-Listen – Join über Codes aus Stammdaten.',
        icon: 'bi-journal-check',
        layer: 'planer',
        href: '../tools/schularbeiten-planer.html',
        countKey: 'schularbeitenSite'
    },
    {
        id: 'plan-projektwochen',
        title: 'Projektwochen',
        description: 'App auf PW-Listen – Klassen & Lehrkräfte.',
        icon: 'bi-calendar2-heart',
        layer: 'planer',
        href: '../tools/projektwochen.html',
        countKey: 'projektwochenSite'
    },
    {
        id: 'plan-freistellungen',
        title: 'Freistellungen-Planer',
        description: 'Anträge lesen/genehmigen – SharePoint-Liste.',
        icon: 'bi-calendar2-check',
        layer: 'planer',
        href: '../tools/freistellung-planer.html',
        countKey: 'freistellungSite'
    },
    {
        id: 'plan-aktivitaeten',
        title: 'Schulaktivitäten-Planer',
        description: 'Exkursionen beantragen & freigeben.',
        icon: 'bi-bus-front',
        layer: 'planer',
        href: '../tools/schulaktivitaeten-planer.html',
        countKey: 'aktivitaetenSite'
    },
    {
        id: 'm365-catalog',
        title: 'M365-Gruppen (Einrichtung)',
        description: 'catalogLinks: Stammdaten ↔ Entra-Gruppen.',
        icon: 'bi-microsoft-teams',
        layer: 'm365',
        href: '../ersteinrichtung.html',
        countKey: 'catalogLinks'
    }
];

/** @type {DatenLinkDef[]} */
export const DATEN_LINKS = [
    {
        id: 'students-classes',
        from: 'year-students',
        to: 'stamm-classes',
        label: 'Klasse',
        detail: 'Feld klasse/code pro Schüler:in'
    },
    {
        id: 'guardians-students',
        from: 'year-guardians',
        to: 'year-students',
        label: 'Betreuung',
        detail: 'guardianIds am Schüler-Datensatz'
    },
    {
        id: 'teachers-classes-kv',
        from: 'stamm-teachers',
        to: 'stamm-classes',
        label: 'KV',
        detail: 'headEmail der Klasse → Lehrkraft'
    },
    {
        id: 'unterricht-teachers',
        from: 'year-unterricht',
        to: 'stamm-teachers',
        label: 'Lehrkraft-Code',
        detail: 'lehrerCode / E-Mail in Belegung'
    },
    {
        id: 'unterricht-classes',
        from: 'year-unterricht',
        to: 'stamm-classes',
        label: 'Klasse',
        detail: 'klasse in Belegungszeile'
    },
    {
        id: 'unterricht-subjects',
        from: 'year-unterricht',
        to: 'stamm-subjects',
        label: 'Fach',
        detail: 'fach-Code in Belegung'
    },
    {
        id: 'arges-subjects',
        from: 'stamm-arges',
        to: 'stamm-subjects',
        label: 'Fach in ARGE',
        detail: 'subjects[] an ARGE-Stammdaten'
    },
    {
        id: 'hub-classes',
        from: 'stamm-hub',
        to: 'stamm-classes',
        label: 'pflegt',
        detail: 'Tenant-Einstellungen'
    },
    {
        id: 'hub-teachers',
        from: 'stamm-hub',
        to: 'stamm-teachers',
        label: 'pflegt'
    },
    {
        id: 'hub-subjects',
        from: 'stamm-hub',
        to: 'stamm-subjects',
        label: 'pflegt'
    },
    {
        id: 'spo-sp-klassen-stamm',
        from: 'spo-sp-klassen',
        to: 'stamm-classes',
        label: 'Sync',
        detail: 'Export Zeilen aus Stammdaten'
    },
    {
        id: 'spo-sp-teachers-stamm',
        from: 'spo-sp-faecher',
        to: 'stamm-subjects',
        label: 'Sync'
    },
    {
        id: 'spo-sp-schueler-stamm',
        from: 'spo-sp-schueler',
        to: 'year-students',
        label: 'Sync',
        detail: 'Schüler:innen-Liste'
    },
    {
        id: 'spo-lehrer-teachers',
        from: 'spo-lehrer',
        to: 'stamm-teachers',
        label: 'Sync',
        detail: 'Lehrerliste aus Stammdaten'
    },
    {
        id: 'spo-sa-plan',
        from: 'spo-sap-schularbeiten',
        to: 'plan-schularbeiten',
        label: 'Daten',
        detail: 'System of Record SharePoint'
    },
    {
        id: 'spo-pw-plan',
        from: 'spo-pw-angebote',
        to: 'plan-projektwochen',
        label: 'Daten'
    },
    {
        id: 'spo-fr-plan',
        from: 'spo-freistellungen',
        to: 'plan-freistellungen',
        label: 'Daten'
    },
    {
        id: 'spo-akt-plan',
        from: 'spo-aktivitaeten',
        to: 'plan-aktivitaeten',
        label: 'Daten'
    },
    {
        id: 'webuntis-students',
        from: 'import-webuntis',
        to: 'year-students',
        label: 'Import',
        detail: 'WebUntis CSV / Stammdaten-Handoff'
    },
    {
        id: 'sa-classes',
        from: 'plan-schularbeiten',
        to: 'stamm-classes',
        label: 'Join',
        detail: 'klasseCode in Schularbeit'
    },
    {
        id: 'sa-teachers',
        from: 'plan-schularbeiten',
        to: 'stamm-teachers',
        label: 'Join',
        detail: 'lehrerCode'
    },
    {
        id: 'sa-subjects',
        from: 'plan-schularbeiten',
        to: 'stamm-subjects',
        label: 'Join',
        detail: 'fachCode'
    },
    {
        id: 'pw-classes',
        from: 'plan-projektwochen',
        to: 'stamm-classes',
        label: 'Join'
    },
    {
        id: 'pw-teachers',
        from: 'plan-projektwochen',
        to: 'stamm-teachers',
        label: 'Join'
    },
    {
        id: 'fr-students',
        from: 'plan-freistellungen',
        to: 'year-students',
        label: 'Join',
        detail: 'Schüler:in / Klasse im Antrag'
    },
    {
        id: 'akt-classes',
        from: 'plan-aktivitaeten',
        to: 'stamm-classes',
        label: 'Join'
    },
    {
        id: 'm365-classes',
        from: 'm365-catalog',
        to: 'stamm-classes',
        label: 'Gruppe',
        detail: 'catalogLinks kind=class'
    },
    {
        id: 'm365-subjects',
        from: 'm365-catalog',
        to: 'stamm-subjects',
        label: 'Gruppe'
    },
    {
        id: 'm365-arges',
        from: 'm365-catalog',
        to: 'stamm-arges',
        label: 'Gruppe'
    }
];

/** @type {Record<string, string>} */
export const LAYER_LABELS = {
    stamm: 'Stammdaten',
    schuljahr: 'Schuljahr & Belegung',
    sharepoint: 'SharePoint-Listen',
    planer: 'Planer & Apps',
    m365: 'Microsoft 365'
};

/**
 * @param {string} layer
 * @param {DatenBlockDef[]} blocks
 */
export function blocksInLayer(layer, blocks) {
    return (blocks || DATEN_BLOECKE).filter((b) => b.layer === layer);
}
