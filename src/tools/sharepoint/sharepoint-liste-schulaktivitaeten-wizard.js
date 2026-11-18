/**
 * Mehrschritt-Assistent: Schulaktivitäten-Listen anlegen.
 */

export const SPAKT_WIZARD_STEP_COUNT = 4;
export const SPAKT_WIZARD_STEP_STORAGE_KEY = 'ms365-spakt-wizard-step-v1';

export const SPAKT_WIZARD_STEP_LABELS = [
    'Website',
    'Genehmigung',
    'Einreicher',
    'Paket starten'
];

/**
 * @param {number} n
 */
export function clampWizardStep(n) {
    const step = parseInt(String(n), 10);
    if (!Number.isFinite(step)) return 1;
    return Math.max(1, Math.min(SPAKT_WIZARD_STEP_COUNT, step));
}

export function loadWizardStep() {
    try {
        const n = parseInt(localStorage.getItem(SPAKT_WIZARD_STEP_STORAGE_KEY) || '1', 10);
        return clampWizardStep(n);
    } catch {
        return 1;
    }
}

/**
 * @param {number} n
 */
export function saveWizardStep(n) {
    try {
        localStorage.setItem(SPAKT_WIZARD_STEP_STORAGE_KEY, String(clampWizardStep(n)));
    } catch {
        /* ignore */
    }
}

/**
 * @param {number} step
 * @param {{
 *   siteUrl?: string,
 *   skipPerms?: boolean,
 *   hasAdmin?: boolean,
 *   hasEinreicher?: boolean
 * }} ctx
 * @returns {string|null}
 */
export function validateWizardStep(step, ctx) {
    const s = clampWizardStep(step);
    const site = String((ctx && ctx.siteUrl) || '').trim();
    if (s === 1 && !site) {
        return 'Bitte die SharePoint-Website eintragen (Schritt 1).';
    }
    // Schritte 2–3: Genehmigung/Einreicher empfohlen, aber erst beim Paket-Lauf verbindlich
    void ctx;
    return null;
}

/**
 * @param {{
 *   siteUrl?: string,
 *   skipPerms?: boolean,
 *   adminLabel?: string,
 *   einreicherLabel?: string,
 *   adminUsersCount?: number,
 *   einreicherUsersCount?: number
 * }} ctx
 */
export function buildWizardSummary(ctx) {
    const site = String((ctx && ctx.siteUrl) || '').trim() || '(keine URL)';
    const skip = !!(ctx && ctx.skipPerms);
    const admin =
        String((ctx && ctx.adminLabel) || '').trim() ||
        ((ctx && ctx.adminUsersCount) ? ctx.adminUsersCount + ' Person(en)' : '— nicht gesetzt —');
    const einreicher =
        String((ctx && ctx.einreicherLabel) || '').trim() ||
        ((ctx && ctx.einreicherUsersCount)
            ? ctx.einreicherUsersCount + ' Person(en)'
            : '— nicht gesetzt —');
    return {
        site,
        skipPerms: skip,
        admin,
        einreicher,
        listsText:
            'Lehrerinnen + Klassen (aus Stammdaten) · Aktivitaet-Regelwerk · Schulaktivitaeten'
    };
}
