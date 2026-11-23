/**
 * Katalog-Stopp (a11): jedes Dashboard-Werkzeug ist Kern, Pflege oder Archiv.
 * Kern = Alltags-/Verkaufsfokus · Pflege = nötig für Betrieb, kein Neuentwurf ·
 * Archiv = keine Weiterentwicklung, ggf. später ausblenden.
 */

import {
    DASHBOARD_LEHRER_SCHUELER_APPS,
    DASHBOARD_TOOL_LABELS
} from './dashboard-audience-catalog.js';

/** @typedef {'kern'|'pflege'|'archiv'} ToolLifecycle */

/** Explizite Overrides (sonst Heuristik aus Audience/Namen). */
/** @type {Record<string, ToolLifecycle>} */
export const TOOL_LIFECYCLE_OVERRIDES = {
    // Kern – Lehrer-/Schüler-Portal und Jahreswechsel
    'schularbeiten-planer': 'kern',
    'freistellung-planer': 'kern',
    'lehrer-freistellung-planer': 'kern',
    'schulaktivitaeten-planer': 'kern',
    projektwochen: 'kern',
    'termine-hub': 'kern',
    'srdp-anmeldung': 'kern',
    schulsprecherwahl: 'kern',
    einsammler: 'kern',
    elternbrief: 'kern',
    krankmeldung: 'kern',
    kursteams: 'kern',
    'organisations-assistent': 'kern',
    'webuntis-stammdaten-import': 'kern',
    'schul-organigramm': 'kern',
    'playbook-schuljahresstart': 'kern',
    'sharepoint-intranet-hub': 'kern',

    // Pflege – IT-Setup und Listen, weiter nutzen
    'freistellung-setup': 'pflege',
    'lehrer-freistellung-setup': 'pflege',
    'schulaktivitaeten-setup': 'pflege',
    'sharepoint-liste-schularbeiten': 'pflege',
    'sharepoint-liste-schulaktivitaeten': 'pflege',
    'sharepoint-liste-srdp': 'pflege',
    'sharepoint-liste-einsammler': 'pflege',
    'sharepoint-liste-elternbrief': 'pflege',
    'sharepoint-liste-krankmeldung': 'pflege',
    'stammdaten-uebergabe': 'pflege',
    'personen-verwaltung': 'pflege',
    lizenzverwaltung: 'pflege',
    'gaeste-verwalten': 'pflege',
    'schueler-lifecycle': 'pflege',
    'klassen-umbenennen': 'pflege',
    'teams-archiv': 'pflege',
    datenhygiene: 'pflege',
    'it-berechtigungen': 'pflege',

    // Archiv – kein Fokus bis Weihnachten
    spielwiesen: 'archiv',
    'teams-kanaele-bulk': 'archiv',
    'bildungsportal-stammdaten': 'archiv',
    'nextcloud-teams-migrate': 'archiv',
    'datei-migration': 'archiv',
    'pa-seminar': 'archiv',
    'pa-schilf': 'archiv',
    schulgraph: 'archiv',
    'webuntis-sync-monitor': 'archiv'
};

/**
 * @param {string} toolId
 * @returns {ToolLifecycle}
 */
export function lifecycleForTool(toolId) {
    const id = String(toolId || '').trim();
    if (!id) return 'archiv';
    if (TOOL_LIFECYCLE_OVERRIDES[id]) return TOOL_LIFECYCLE_OVERRIDES[id];
    if (DASHBOARD_LEHRER_SCHUELER_APPS.includes(id)) return 'kern';
    if (id.startsWith('pa-') || id.startsWith('playbook-')) return 'pflege';
    if (id.startsWith('sharepoint-liste-') || id.endsWith('-setup')) return 'pflege';
    return 'pflege';
}

/**
 * @returns {Array<{ id: string, label: string, lifecycle: ToolLifecycle }>}
 */
export function listToolLifecycles() {
    return Object.keys(DASHBOARD_TOOL_LABELS)
        .sort()
        .map((id) => ({
            id,
            label: DASHBOARD_TOOL_LABELS[id] || id,
            lifecycle: lifecycleForTool(id)
        }));
}

/**
 * @param {ToolLifecycle} lifecycle
 * @returns {Array<{ id: string, label: string, lifecycle: ToolLifecycle }>}
 */
export function toolsByLifecycle(lifecycle) {
    return listToolLifecycles().filter((t) => t.lifecycle === lifecycle);
}

/**
 * @returns {{ kern: number, pflege: number, archiv: number, total: number }}
 */
export function lifecycleCounts() {
    const all = listToolLifecycles();
    const counts = { kern: 0, pflege: 0, archiv: 0, total: all.length };
    for (const t of all) counts[t.lifecycle] += 1;
    return counts;
}
