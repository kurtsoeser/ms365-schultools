/**
 * Mehrschritt-Assistent für sharepoint-liste-stammdaten.html
 */

export const SPS_WIZARD_STEP_COUNT = 5;
export const SPS_WIZARD_STEP_STORAGE_KEY = 'ms365-sps-wizard-step-v1';

export const SPS_WIZARD_STEP_LABELS = [
    'Website',
    'Listen wählen',
    'Abgleich-Modus',
    'Berechtigungen',
    'Starten'
];

/**
 * @param {number} n
 */
export function clampWizardStep(n) {
    const step = parseInt(String(n), 10);
    if (!isFinite(step)) return 1;
    return Math.max(1, Math.min(SPS_WIZARD_STEP_COUNT, step));
}

export function loadWizardStep() {
    try {
        const n = parseInt(localStorage.getItem(SPS_WIZARD_STEP_STORAGE_KEY) || '1', 10);
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
        localStorage.setItem(SPS_WIZARD_STEP_STORAGE_KEY, String(clampWizardStep(n)));
    } catch {
        /* ignore */
    }
}

/**
 * @param {Record<string, boolean>} lists
 * @returns {string|null} Fehlermeldung oder null
 */
export function validateListSelection(lists) {
    const o = lists || {};
    const any =
        o.schueler || o.faecher || o.fachgruppen || o.arges || o.klassen || o.lehrer;
    if (!any) return 'Bitte mindestens eine Liste auswählen.';
    return null;
}

/**
 * @param {number} step
 * @param {{ siteUrl?: string, lists?: Record<string, boolean> }} ctx
 * @returns {string|null}
 */
export function validateWizardStep(step, ctx) {
    const s = clampWizardStep(step);
    const site = String(ctx && ctx.siteUrl ? ctx.siteUrl : '').trim();
    if (s === 1 && !site) {
        return 'Bitte die SharePoint-Website eintragen (Schritt 1).';
    }
    if (s === 2) {
        const listErr = validateListSelection(ctx && ctx.lists);
        if (listErr) return listErr;
        const titles = ctx && ctx.listTitles ? ctx.listTitles : {};
        const lists = ctx && ctx.lists ? ctx.lists : {};
        const keys = [
            ['schueler', 'Schülerinnenliste'],
            ['faecher', 'Fächerliste'],
            ['fachgruppen', 'Fachgruppenliste'],
            ['arges', 'ARGE-Liste'],
            ['klassen', 'Klassenliste']
        ];
        for (let i = 0; i < keys.length; i++) {
            const k = keys[i][0];
            if (!lists[k]) continue;
            if (!String(titles[k] || '').trim()) {
                return 'Bitte für „' + keys[i][1] + '“ einen Listenname eintragen.';
            }
        }
    }
    return null;
}

/**
 * @param {{
 *   siteUrl?: string,
 *   syncMode?: boolean,
 *   removeOrphans?: boolean,
 *   lists?: Record<string, boolean>,
 *   listTitles?: Record<string, string>,
 *   skipPerms?: boolean
 * }} ctx
 */
export function buildWizardSummary(ctx) {
    const site = String(ctx.siteUrl || '').trim() || '(keine URL)';
    const mode = ctx.syncMode
        ? 'Abgleich vorhandener Listen (gleicher Name)'
        : 'Immer neue Listen anlegen';
    const orphans = ctx.removeOrphans ? 'Verwaiste Zeilen entfernen' : 'Alte SP-Zeilen behalten';
    const lists = ctx.lists || {};
    const titles = ctx.listTitles || {};
    const picked = [];
    if (lists.schueler) picked.push('Schülerinnen („' + (titles.schueler || 'Schülerinnen') + '“)');
    if (lists.faecher) picked.push('Fächer („' + (titles.faecher || 'Fächer') + '“)');
    if (lists.fachgruppen) picked.push('Fachgruppen („' + (titles.fachgruppen || 'Fachgruppen') + '“)');
    if (lists.arges) picked.push('ARGEs („' + (titles.arges || 'ARGEs') + '“)');
    if (lists.klassen) picked.push('Klassen („' + (titles.klassen || 'Klassen') + '“)');
    const perms = ctx.skipPerms ? 'Berechtigungen überspringen' : 'Berechtigungen nach Abgleich anwenden';
    return {
        site,
        mode,
        orphans,
        listsText: picked.length ? picked.join(' · ') : '— keine Liste gewählt —',
        perms
    };
}

/**
 * @param {number} step
 */
export function wizardPhaseHint(step) {
    const s = clampWizardStep(step);
    const labels = SPS_WIZARD_STEP_LABELS;
    return 'Schritt ' + s + ' von ' + SPS_WIZARD_STEP_COUNT + ': ' + (labels[s - 1] || '');
}
