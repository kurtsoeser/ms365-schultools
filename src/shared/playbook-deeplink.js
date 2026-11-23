/**
 * Playbook Deep-Link-Kontrakt (Wave 0).
 *
 * Konvention für CTAs aus Playbook-Schritten:
 * - Stammdaten-Tabs: Hash an tenant.html (#lehrer, #klassen, …)
 * - Stammdaten-Setup-Panes: #profile, #spo-planer (Planer & Apps), …
 * - Schularbeiten-Planer: ?view=<VIEWS.id> (regeln, admin, …)
 * - Freistellung-Setup: ?step=1…7 (Wizard-Pane)
 * - Sonst: direkte Tool-URL; Deep-Link nachrüsten, sobald das Ziel es kann
 *
 * Entscheidungen (verbindlich, Saison 1):
 * - Jährlicher Pfad = Schuljahresstart-Playbook; Org-Assistent = Verweis/Unterstützung
 * - Ready-Checks gewünscht (Hybrid mit Checkbox); siehe playbook-ready.js
 * - Eltern + Elternsprechtag bleiben zwei Playbooks
 * - Neue App-Playbooks: PLANNED_APP_PLAYBOOKS in dashboard-playbooks-catalog.js
 */

/** @typedef {{ href: string, label?: string }} PlaybookDeepLink */

/** Stammdaten-Hashes (tenant.html#…). spo-planer öffnet Tab + Setup-Pane Planer & Apps. */
export const PLAYBOOK_STAMMDATEN_HASH = {
    profil: '#stammdaten',
    planerApps: '#spo-planer',
    lehrer: '#lehrer',
    schueler: '#schueler',
    klassen: '#klassen',
    faecher: '#faecher'
};

/**
 * @param {string} [hash] z. B. '#lehrer' oder 'lehrer'
 * @returns {string} relativer Pfad von tools/ aus
 */
export function stammdatenDeepLink(hash) {
    const raw = String(hash || 'stammdaten').replace(/^#/, '').trim() || 'stammdaten';
    return '../tenant.html#' + raw;
}

/**
 * Schularbeiten-Planer-Views (muss zu VIEWS in schularbeiten-planer-state.js passen).
 * @param {'dashboard'|'liste'|'kalender'|'neu'|'meine'|'regeln'|'admin'|string} view
 * @returns {string}
 */
export function schularbeitenPlanerDeepLink(view) {
    const v = String(view || '').trim();
    if (!v) return 'schularbeiten-planer.html';
    return 'schularbeiten-planer.html?view=' + encodeURIComponent(v);
}

/**
 * Freistellung-Setup-Wizard-Schritte 1–7.
 * @param {number} step
 * @returns {string}
 */
export function freistellungSetupDeepLink(step) {
    const n = Math.max(1, Math.min(7, parseInt(String(step), 10) || 1));
    return 'freistellung-setup.html?step=' + n;
}

if (typeof window !== 'undefined') {
    window.ms365PlaybookDeeplink = {
        PLAYBOOK_STAMMDATEN_HASH,
        stammdatenDeepLink,
        schularbeitenPlanerDeepLink,
        freistellungSetupDeepLink
    };
}
