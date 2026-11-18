/**
 * Ein Modus für das Dashboard-Layout (eine Quelle der Wahrheit).
 *
 * full – Schul-IT: kompletter Werkzeugkatalog (bewusster Einstieg)
 * preview-lehrer – IT im Schulportal (Lehrer-Ansicht) – Standard für IT
 * preview-schueler – IT-Vorschau Schüler-Portal
 * apps-lehrer | apps-schueler – echte Lehrkraft / Schüler/in
 *
 * Speicher (IT): leer|portal|lehrer → Schulportal; full|it → IT-Bereich; schueler → Schüler-Vorschau
 */

const IT_PREVIEW_KEY = 'ms365-dashboard-it-preview-v1';

function itPreviewStorageKey() {
    let tid = '';
    try {
        if (typeof window !== 'undefined' && typeof window.ms365AuthGetAccountInfo === 'function') {
            const info = window.ms365AuthGetAccountInfo();
            tid = info && info.tenantId ? String(info.tenantId).trim() : '';
        }
    } catch {
        /* ignore */
    }
    return tid ? IT_PREVIEW_KEY + ':' + tid : IT_PREVIEW_KEY;
}

/** @typedef {'full'|'preview-lehrer'|'preview-schueler'|'apps-lehrer'|'apps-schueler'} DashboardLayoutMode */

/**
 * @returns {'portal'|'full'|'schueler'}
 */
export function readItShellChoice() {
    try {
        const v = String(localStorage.getItem(itPreviewStorageKey()) || '')
            .trim()
            .toLowerCase();
        if (v === 'full' || v === 'it') return 'full';
        if (v === 'schueler') return 'schueler';
        // leer, portal, lehrer → Schulportal zuerst
        return 'portal';
    } catch {
        return 'portal';
    }
}

/**
 * @deprecated Alias – nutzt readItShellChoice
 * @returns {'lehrer'|'schueler'|'full'}
 */
export function readItPreviewChoice() {
    const c = readItShellChoice();
    if (c === 'full') return 'full';
    if (c === 'schueler') return 'schueler';
    return 'lehrer';
}

/**
 * @param {import('./dashboard-audience-resolve.js').DashboardPersonas} personas
 * @returns {DashboardLayoutMode}
 */
export function resolveLayoutModeFromPersonas(personas) {
    // Unklar / ausgeloggt: Portal-Chrome, nie IT-Katalog flashen
    if (!personas || !personas.loggedIn) return 'preview-lehrer';

    if (personas.isIt) {
        const p = readItShellChoice();
        if (p === 'full') return 'full';
        if (p === 'schueler') return 'preview-schueler';
        return 'preview-lehrer';
    }

    if (personas.isLehrer) return 'apps-lehrer';
    return 'apps-schueler';
}

/**
 * @param {DashboardLayoutMode} mode
 */
export function isAppsOnlyLayout(mode) {
    return (
        mode === 'apps-lehrer' ||
        mode === 'apps-schueler' ||
        mode === 'preview-lehrer' ||
        mode === 'preview-schueler'
    );
}

/**
 * @param {DashboardLayoutMode} mode
 * @returns {import('./dashboard-audience-catalog.js').DashboardView}
 */
export function catalogViewForLayout(mode) {
    if (mode === 'full') return 'it';
    if (mode === 'preview-lehrer' || mode === 'apps-lehrer') return 'lehrer';
    return 'schueler';
}

/**
 * @param {'full'|'preview-lehrer'|'preview-schueler'} mode
 */
export function writeItPreviewChoice(mode) {
    try {
        const storageKey = itPreviewStorageKey();
        if (mode === 'full') localStorage.setItem(storageKey, 'full');
        else if (mode === 'preview-lehrer') localStorage.setItem(storageKey, 'portal');
        else if (mode === 'preview-schueler') localStorage.setItem(storageKey, 'schueler');
    } catch {
        /* ignore */
    }
}

/**
 * @param {DashboardLayoutMode} mode
 */
export function layoutModeLabel(mode) {
    if (mode === 'full') return 'IT-Backend';
    if (mode === 'preview-lehrer') return 'Portal';
    if (mode === 'preview-schueler') return 'Vorschau Schüler/in';
    if (mode === 'apps-lehrer') return 'Lehrkraft';
    return 'Schüler/in';
}
