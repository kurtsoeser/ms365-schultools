/**
 * Playbook-Metadaten für das Dashboard (Fortschritt aus localStorage).
 */
import { SCHULJAHRSTART_PLAYBOOK_STORAGE_KEY, SCHULJAHRSTART_STEP_IDS } from './playbook-schuljahresstart-gates.js';

/** @typedef {{ id: string, title: string, icon: string, href: string, storageKey: string, stepIds: string[], blurb?: string }} DashboardPlaybookDef */

/** @type {DashboardPlaybookDef[]} */
export const DASHBOARD_PLAYBOOKS = [
    {
        id: 'daten-import',
        title: 'Daten importieren & verknüpfen',
        icon: 'bi-box-arrow-in-down',
        href: 'tools/playbook-daten-import-verknuepfen.html',
        storageKey: 'ms365-playbook-daten-import-verknuepfen-v2',
        stepIds: [
            'register',
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
        blurb: 'Einstieg für neue Mandanten: Quelle, Import, Review, SLG und Klassen.'
    },
    {
        id: 'schuljahresstart',
        title: 'Schuljahresstart',
        icon: 'bi-rocket-takeoff',
        href: 'tools/playbook-schuljahresstart.html',
        storageKey: SCHULJAHRSTART_PLAYBOOK_STORAGE_KEY,
        stepIds: [...SCHULJAHRSTART_STEP_IDS],
        blurb: 'Stammdaten, Klassen, SLG, Kursteams und Cleanup in einer Checkliste.'
    },
    {
        id: 'intranet',
        title: 'Intranet aufsetzen',
        icon: 'bi-house-door',
        href: 'tools/playbook-intranet.html',
        storageKey: 'ms365-playbook-intranet-v1',
        stepIds: ['hub', 'lehrer', 'termine', 'sa', 'pw', 'vertretung', 'nutzen'],
        blurb: 'Hub, Listen und Planer – einmal einrichten, dann nutzen.'
    },
    {
        id: 'schularbeiten-planer',
        title: 'Schularbeiten-Planer',
        icon: 'bi-journal-check',
        href: 'tools/playbook-schularbeiten-planer.html',
        storageKey: 'ms365-playbook-schularbeiten-v1',
        stepIds: ['stammdaten', 'listen', 'regelwerk', 'berechtigungen', 'test', 'intranet', 'schulung'],
        blurb: 'Listen, Regelwerk, Rechte und Testantrag bis zur Nutzung durch Lehrkräfte.'
    },
    {
        id: 'freistellungen',
        title: 'Freistellungen',
        icon: 'bi-person-check',
        href: 'tools/playbook-freistellungen.html',
        storageKey: 'ms365-playbook-freistellungen-v1',
        stepIds: ['konzept', 'pa-basis', 'vorbereitung', 'liste', 'flow', 'planer-config', 'live'],
        blurb: 'Konzept, Power Automate, SharePoint-Liste und erster Testantrag.'
    },
    {
        id: 'kursteams',
        title: 'Unterrichtsteams',
        icon: 'bi-mortarboard',
        href: 'tools/playbook-kursteams.html',
        storageKey: 'ms365-playbook-kursteams-v1',
        stepIds: ['import', 'klassen', 'kursteams', 'templates', 'monitor', 'katalog'],
        blurb: 'Import, SLG, Kursteams, Vorlagen und Qualitätscheck im Monitor.'
    },
    {
        id: 'eltern',
        title: 'Elternkommunikation',
        icon: 'bi-chat-heart',
        href: 'tools/playbook-eltern.html',
        storageKey: 'ms365-playbook-eltern-v1',
        stepIds: ['guardians', 'verteiler', 'bookings', 'intranet', 'pa'],
        blurb: 'Verteiler, Sprechtag und Intranet-Hinweise als Kanal.'
    },
    {
        id: 'elternsprechtag',
        title: 'Elternsprechtag',
        icon: 'bi-calendar2-heart',
        href: 'tools/playbook-elternsprechtag.html',
        storageKey: 'ms365-playbook-elternsprechtag-v1',
        stepIds: ['termin', 'guardians', 'verteiler', 'bookings', 'lehrer', 'kommunikation', 'nachbereitung'],
        blurb: 'Termin, Bookings, Verteiler und Elterninfo – fokussiert auf den Sprechtag.'
    },
    {
        id: 'cleanup',
        title: 'Aufräumen',
        icon: 'bi-broom',
        href: 'tools/cleanup-playbook.html',
        storageKey: 'ms365-cleanup-playbook-v1',
        stepIds: ['owners', 'empty', 'archiv', 'hygiene', 'naming', 'gaeste', 'uebergabe'],
        blurb: 'Leere Gruppen, Besitzlose, Archiv und Hygiene.'
    }
];

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
