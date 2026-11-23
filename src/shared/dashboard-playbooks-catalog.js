/**
 * Playbook-Metadaten für das Dashboard (Fortschritt aus localStorage).
 *
 * Entscheidungen Saison 1 (Wave 0):
 * - Jährlicher Einstieg = Schuljahresstart-Playbook; Org-Assistent nur unterstützend
 * - Ready-Checks (Hybrid) gewünscht → playbook-ready.js / wirePlaybookPage
 * - Eltern + Elternsprechtag bleiben zwei getrennte Playbooks
 * - Schul-App-Playbooks: Wave 3 ausgearbeitet, Kategorie schulapps in DASHBOARD_PLAYBOOKS
 * - Deep-Links: playbook-deeplink.js
 */
import { SCHULJAHRSTART_PLAYBOOK_STORAGE_KEY, SCHULJAHRSTART_STEP_IDS } from './playbook-schuljahresstart-gates.js';

/** @typedef {{ id: string, label: string, icon: string }} PlaybookCategoryDef */
/** @typedef {{ id: string, title: string, icon: string, href: string, storageKey: string, stepIds: string[], blurb?: string, detail?: string, categoryId: string, planned?: boolean }} DashboardPlaybookDef */

/** @type {PlaybookCategoryDef[]} */
export const PLAYBOOK_CATEGORIES = [
    { id: 'grundlagen', label: 'Einstieg & Schuljahr', icon: 'bi-flag' },
    { id: 'unterricht', label: 'Unterricht & Teams', icon: 'bi-mortarboard' },
    { id: 'kommunikation', label: 'Eltern & Kanäle', icon: 'bi-chat-heart' },
    { id: 'intranet', label: 'Intranet & Planung', icon: 'bi-house-door' },
    { id: 'schulapps', label: 'Schul-Apps', icon: 'bi-window-stack' }
];

/** @deprecated Wave 3 – Apps sind live; Alias für ältere Imports */
export const PLANNED_PLAYBOOK_CATEGORY = {
    id: 'schulapps',
    label: 'Schul-Apps',
    icon: 'bi-window-stack'
};

/** @type {DashboardPlaybookDef[]} */
export const DASHBOARD_PLAYBOOKS = [
    {
        id: 'daten-import',
        categoryId: 'grundlagen',
        title: 'Erste Einrichtung (Blaupause)',
        icon: 'bi-flag',
        href: 'tools/playbook-daten-import-verknuepfen.html',
        storageKey: 'ms365-playbook-daten-import-verknuepfen-v2',
        stepIds: [
            'register',
            'tool-site',
            'quelle',
            'import-datei',
            'bildungsportal-info',
            'review',
            'slg',
            'klassen',
            'faecher',
            'hygiene',
            'backup'
        ],
        blurb: 'Blaupause für neue Schulen: Profil, Tool-Site, Import, Verknüpfen, Sicherung – Bildungsportal optional später.',
        detail: 'Import & Verknüpfen'
    },
    {
        id: 'schuljahresstart',
        categoryId: 'grundlagen',
        title: 'Schuljahresstart',
        icon: 'bi-rocket-takeoff',
        href: 'tools/playbook-schuljahresstart.html',
        storageKey: SCHULJAHRSTART_PLAYBOOK_STORAGE_KEY,
        stepIds: [...SCHULJAHRSTART_STEP_IDS],
        blurb:
            'Primärer Jahres-Einstieg: Stammdaten, Klassen, SLG, Kursteams, Cleanup – Org-Assistent nur unterstützend.',
        detail: 'Jahresstart (primär)'
    },
    {
        id: 'intranet',
        categoryId: 'intranet',
        title: 'Intranet aufsetzen',
        icon: 'bi-house-door',
        href: 'tools/playbook-intranet.html',
        storageKey: 'ms365-playbook-intranet-v1',
        stepIds: ['hub', 'lehrer', 'termine', 'sa', 'pw', 'vertretung', 'nutzen'],
        blurb: 'Hub und Basislisten nach der Blaupause; SA/PW optional mit eigenem Playbook.',
        detail: 'SharePoint & Hub'
    },
    {
        id: 'schularbeiten-planer',
        categoryId: 'intranet',
        title: 'Schularbeiten-Planer',
        icon: 'bi-journal-check',
        href: 'tools/playbook-schularbeiten-planer.html',
        storageKey: 'ms365-playbook-schularbeiten-v1',
        stepIds: ['stammdaten', 'listen', 'regelwerk', 'berechtigungen', 'test', 'intranet', 'schulung'],
        blurb: 'Listen, Regelwerk, Rechte und Testantrag bis zur Nutzung durch Lehrkräfte.',
        detail: 'Listen & Genehmigung'
    },
    {
        id: 'freistellungen',
        categoryId: 'unterricht',
        title: 'Freistellungen',
        icon: 'bi-person-check',
        href: 'tools/playbook-freistellungen.html',
        storageKey: 'ms365-playbook-freistellungen-v1',
        stepIds: ['konzept', 'pa-basis', 'vorbereitung', 'liste', 'flow', 'planer-config', 'live'],
        blurb: 'Konzept, Power Automate, SharePoint-Liste und erster Testantrag.',
        detail: 'Power Automate & Liste'
    },
    {
        id: 'kursteams',
        categoryId: 'unterricht',
        title: 'Unterrichtsteams',
        icon: 'bi-mortarboard',
        href: 'tools/playbook-kursteams.html',
        storageKey: 'ms365-playbook-kursteams-v1',
        stepIds: ['import', 'klassen', 'kursteams', 'templates', 'monitor', 'katalog'],
        blurb: 'Import, SLG, Kursteams, Vorlagen und Monitor – oft parallel zum Schuljahresstart.',
        detail: 'SLG, Vorlagen & Monitor'
    },
    {
        id: 'eltern',
        categoryId: 'kommunikation',
        title: 'Elternkommunikation',
        icon: 'bi-chat-heart',
        href: 'tools/playbook-eltern.html',
        storageKey: 'ms365-playbook-eltern-v1',
        stepIds: ['guardians', 'verteiler', 'bookings', 'intranet', 'pa'],
        blurb: 'Kanal-Überblick: Guardians und Verteiler; Sprechtag und Elternbrief separat.',
        detail: 'Verteiler & Kanäle'
    },
    {
        id: 'elternsprechtag',
        categoryId: 'kommunikation',
        title: 'Elternsprechtag',
        icon: 'bi-calendar2-heart',
        href: 'tools/playbook-elternsprechtag.html',
        storageKey: 'ms365-playbook-elternsprechtag-v1',
        stepIds: ['termin', 'guardians', 'verteiler', 'bookings', 'lehrer', 'kommunikation', 'nachbereitung'],
        blurb: 'Termin, Bookings, Verteiler und Elterninfo – nur der Sprechtag.',
        detail: 'Bookings & Kommunikation'
    },
    {
        id: 'cleanup',
        categoryId: 'grundlagen',
        title: 'Aufräumen',
        icon: 'bi-broom',
        href: 'tools/cleanup-playbook.html',
        storageKey: 'ms365-cleanup-playbook-v1',
        stepIds: ['owners', 'empty', 'archiv', 'hygiene', 'naming', 'gaeste', 'uebergabe'],
        blurb: 'Leere Gruppen, Besitzlose, Archiv und Hygiene.',
        detail: 'Gruppen & Besitzer'
    },
    {
        id: 'schulaktivitaeten',
        categoryId: 'schulapps',
        title: 'Schulaktivitäten',
        icon: 'bi-calendar2-check',
        href: 'tools/playbook-schulaktivitaeten.html',
        storageKey: 'ms365-playbook-schulaktivitaeten-v1',
        stepIds: ['site', 'listen', 'rechte', 'test', 'schulung'],
        blurb: 'Primär-Site, AKT-Listen, Rechte und Testantrag für den Schulaktivitäten-Planer.',
        detail: 'Listen & Genehmigung'
    },
    {
        id: 'projektwochen',
        categoryId: 'schulapps',
        title: 'Projektwochen',
        icon: 'bi-calendar2-week',
        href: 'tools/playbook-projektwochen.html',
        storageKey: 'ms365-playbook-projektwochen-v1',
        stepIds: ['site', 'listen', 'bookings', 'test', 'schulung'],
        blurb: 'Listen, Bookings/Angebote und Test – bis Projektwochen laufen.',
        detail: 'Listen & Bookings'
    },
    {
        id: 'einsammler',
        categoryId: 'schulapps',
        title: 'Einsammler',
        icon: 'bi-collection',
        href: 'tools/playbook-einsammler.html',
        storageKey: 'ms365-playbook-einsammler-v1',
        stepIds: ['listen', 'rechte', 'test', 'schulung'],
        blurb: 'SharePoint-Liste, Rechte und Test-Einsammlung.',
        detail: 'Liste & Alltag'
    },
    {
        id: 'elternbrief',
        categoryId: 'schulapps',
        title: 'Elternbrief',
        icon: 'bi-envelope-paper',
        href: 'tools/playbook-elternbrief.html',
        storageKey: 'ms365-playbook-elternbrief-v1',
        stepIds: ['listen', 'verteiler', 'test', 'schulung'],
        blurb: 'Liste, Verteiler und Probebrief – getrennt vom Sprechtag.',
        detail: 'Liste & Versand'
    },
    {
        id: 'krankmeldung',
        categoryId: 'schulapps',
        title: 'Krankmeldung',
        icon: 'bi-heart-pulse',
        href: 'tools/playbook-krankmeldung.html',
        storageKey: 'ms365-playbook-krankmeldung-v1',
        stepIds: ['listen', 'rechte', 'test', 'schulung'],
        blurb: 'Liste, Rechte und Testmeldung für den Krankmeldungs-Kanal.',
        detail: 'Liste & Formular'
    },
    {
        id: 'srdp',
        categoryId: 'schulapps',
        title: 'SRDP-Anmeldung',
        icon: 'bi-mortarboard',
        href: 'tools/playbook-srdp.html',
        storageKey: 'ms365-playbook-srdp-v1',
        stepIds: ['listen', 'rechte', 'test', 'mail', 'schulung'],
        blurb: 'Listen, Rechte, Testanmeldung und optionales Mail-Setup.',
        detail: 'Listen & Anmeldung'
    },
    {
        id: 'wahlen',
        categoryId: 'schulapps',
        title: 'Wahlen (Schulsprecher & Wahlfach)',
        icon: 'bi-ui-checks-grid',
        href: 'tools/playbook-wahlen.html',
        storageKey: 'ms365-playbook-wahlen-v1',
        stepIds: ['konzept', 'listen', 'rechte', 'test', 'schulung'],
        blurb: 'Schulsprecherwahl und Wahlfach-Zuteilung in einer Checkliste.',
        detail: 'Zwei Wahl-Tools'
    }
];

/**
 * @deprecated Wave 3 – leer; Apps sind in DASHBOARD_PLAYBOOKS.
 * @type {DashboardPlaybookDef[]}
 */
export const PLANNED_APP_PLAYBOOKS = [];

/**
 * Alle Playbook-Defs (Dashboard = vollständig nach Wave 3).
 * @returns {DashboardPlaybookDef[]}
 */
export function allPlaybookDefs() {
    return DASHBOARD_PLAYBOOKS.concat(PLANNED_APP_PLAYBOOKS);
}

/**
 * Bootstrap-Icons-Klasse (akzeptiert `bi-name` oder `bi bi-name`).
 * @param {string} [icon]
 */
export function playbookBiClass(icon) {
    const s = String(icon || 'bi-journal-bookmark').trim();
    if (!s) return 'bi bi-journal-bookmark';
    if (s.indexOf('bi ') === 0) return s;
    if (s.indexOf('bi-') === 0) return 'bi ' + s;
    return 'bi bi-' + s;
}

/**
 * @param {DashboardPlaybookDef} def
 */
export function formatPlaybookMetaLine(def) {
    const n = Array.isArray(def.stepIds) ? def.stepIds.length : 0;
    const steps = n === 1 ? '1 Schritt' : n + ' Schritte';
    const parts = [steps, 'Checkliste'];
    if (def.detail) parts.push(def.detail);
    return parts.join(' · ');
}

/**
 * @param {string} categoryId
 * @returns {DashboardPlaybookDef[]}
 */
export function playbooksInCategory(categoryId) {
    return DASHBOARD_PLAYBOOKS.filter(function (p) {
        return p.categoryId === categoryId;
    });
}

/**
 * @param {Record<string, boolean>|null|undefined} state
 * @param {string[]} stepIds
 */
export function computePlaybookProgress(state, stepIds) {
    const ids = Array.isArray(stepIds) ? stepIds : [];
    const st = state && typeof state === 'object' ? state : {};
    let done = 0;
    ids.forEach(function (id) {
        if (st[id]) done += 1;
    });
    const total = ids.length;
    const ratio = total > 0 ? done / total : 0;
    /** @type {'not-started'|'in-progress'|'complete'} */
    let status = 'not-started';
    if (total === 0) {
        status = 'complete';
    } else if (done === 0) {
        status = 'not-started';
    } else if (done >= total) {
        status = 'complete';
    } else {
        status = 'in-progress';
    }
    return { done, total, ratio, status };
}

/**
 * @param {{ done: number, total: number, status: string }} progress
 */
export function formatPlaybookProgressLabel(progress) {
    const p = progress || { done: 0, total: 0, status: 'not-started' };
    if (p.status === 'complete' && p.total > 0) {
        return 'Abgeschlossen';
    }
    if (p.status === 'not-started' || p.done === 0) {
        return 'Noch nicht gestartet';
    }
    if (p.total === 1) {
        return p.done >= 1 ? 'Abgeschlossen' : '1 Schritt offen';
    }
    return p.done + ' von ' + p.total + ' Schritte erledigt';
}

/**
 * @param {number} ratio
 * @param {number} segments
 */
export function playbookProgressBarSegments(ratio, segments) {
    const n = Math.max(4, Math.min(14, segments || 12));
    const filled = Math.round(Math.max(0, Math.min(1, ratio)) * n);
    let out = '';
    for (let i = 0; i < n; i += 1) {
        out += i < filled ? '█' : '░';
    }
    return out;
}
