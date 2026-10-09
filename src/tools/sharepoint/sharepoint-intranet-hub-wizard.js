/**
 * Wizard-Schritte für tools/sharepoint-intranet-hub.html
 */
export const IH_WIZARD_STEP_COUNT = 4;
export const IH_WIZARD_STEP_STORAGE_KEY = 'ms365-intranet-hub-wizard-step-v1';

export function clampWizardStep(n) {
    const x = parseInt(n, 10);
    if (!Number.isFinite(x) || x < 1) return 1;
    if (x > IH_WIZARD_STEP_COUNT) return IH_WIZARD_STEP_COUNT;
    return x;
}

export function loadWizardStep() {
    try {
        const q = new URL(window.location.href).searchParams.get('step');
        if (q) return clampWizardStep(q);
    } catch {
        /* ignore */
    }
    try {
        const stored = localStorage.getItem(IH_WIZARD_STEP_STORAGE_KEY);
        if (stored) return clampWizardStep(stored);
    } catch {
        /* ignore */
    }
    return 1;
}

export function saveWizardStep(step) {
    const s = clampWizardStep(step);
    try {
        localStorage.setItem(IH_WIZARD_STEP_STORAGE_KEY, String(s));
    } catch {
        /* ignore */
    }
    return s;
}

export function wizardPhaseHint(step) {
    const s = clampWizardStep(step);
    const labels = [
        'Schritt 1 von 4: Kommunikationssite anlegen',
        'Schritt 2 von 4: Als Hub registrieren',
        'Schritt 3 von 4: Listen & Startpaket',
        'Schritt 4 von 4: Intranet fertigstellen'
    ];
    return labels[s - 1] || labels[0];
}

/**
 * @param {number} fromStep
 * @param {{ siteUrl?: string }} ctx
 * @returns {string|null} Fehlermeldung oder null
 */
export function validateWizardStep(fromStep, ctx) {
    const step = clampWizardStep(fromStep);
    if (step === 2) {
        const url = String(ctx && ctx.siteUrl ? ctx.siteUrl : '').trim();
        if (!url) {
            return 'Site-URL fehlt: in Schritt 1 anlegen oder in Schritt 2 die Adresse eintragen.';
        }
    }
    return null;
}
