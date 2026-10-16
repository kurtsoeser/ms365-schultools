/**
 * Tool-Nebenstates: UI-/Arbeitsstände neben dem kanonischen v2-Container.
 *
 * Während Lokal-First und danach:
 * - Stammdaten-Wahrheit = nur ms365-schooltool-data-v2 (Listen + schulOrganigramm)
 * - Diese Keys sind Cache / Wizard-Entwurf / Graph-Arbeitsstand – nicht die Schulwahrheit
 *
 * Neue Features sollen Stammdaten (Lehrer, ARGEs, Klassen, Belegung) nur über
 * writeSchoolData / ms365AppDataV2 schreiben, nicht in eigene Keys spiegeln.
 */

/** Kanonischer Speicher (Wahrheit). */
export const STAMMDATEN_CANONICAL_KEY = 'ms365-schooltool-data-v2';

/**
 * Keys, die Domänendaten duplizieren können – nur als flüchtiger Cache behandeln.
 * @type {ReadonlyArray<{ key: string, role: 'ui-cache'|'graph-work'|'wizard'|'legacy', domain: string }>}
 */
export const STAMMDATEN_SIDE_STATE_KEYS = Object.freeze([
    {
        key: 'webuntis-teams-creator-state-v1',
        role: 'ui-cache',
        domain: 'kursteams / unterrichtsbelegung'
    },
    { key: 'ms365-arge-state-v2', role: 'graph-work', domain: 'ARGE-Gruppen Entra' },
    { key: 'ms365-jahrgang-state-v1', role: 'graph-work', domain: 'Jahrgangsgruppen Entra' },
    { key: 'ms365-wtg-state-v1', role: 'graph-work', domain: 'weitere Teams/Gruppen' },
    {
        key: 'ms365-schulstruktur-sync-v1',
        role: 'legacy',
        domain: 'Schulstruktur-Sync (Cache; Spiegel auch in v2.structure)'
    },
    {
        key: 'ms365-schulstruktur-sync-match-v1',
        role: 'legacy',
        domain: 'Struktur-Match (Cache; Spiegel in v2.match)'
    },
    {
        key: 'ms365-schulstruktur-sync-tenant-cache-v1',
        role: 'ui-cache',
        domain: 'Graph-Tenant-Cache Struktur'
    },
    { key: 'ms365-schul-org-view-mode-v1', role: 'ui-cache', domain: 'Organigramm Ansicht' },
    { key: 'ms365-datenlandkarte-layout-v1', role: 'ui-cache', domain: 'Datenlandkarte Layout' },
    { key: 'ms365-schulgraph-options-v2', role: 'ui-cache', domain: 'Schulgraph Optionen' }
]);

/**
 * @param {string} key
 * @returns {boolean}
 */
export function isStammdatenSideStateKey(key) {
    const k = String(key || '');
    return STAMMDATEN_SIDE_STATE_KEYS.some(function (row) {
        return row.key === k;
    });
}

/**
 * @returns {string}
 */
export function sideStatePolicyNote() {
    return (
        'Nebenstates (' +
        STAMMDATEN_SIDE_STATE_KEYS.length +
        ' Keys) sind kein Stammdaten-Kanon. Wahrheit: ' +
        STAMMDATEN_CANONICAL_KEY +
        '.'
    );
}

if (typeof window !== 'undefined') {
    window.ms365StammdatenSideStateKeys = {
        STAMMDATEN_CANONICAL_KEY,
        STAMMDATEN_SIDE_STATE_KEYS,
        isStammdatenSideStateKey,
        sideStatePolicyNote
    };
}

export default {
    STAMMDATEN_CANONICAL_KEY,
    STAMMDATEN_SIDE_STATE_KEYS,
    isStammdatenSideStateKey,
    sideStatePolicyNote
};
