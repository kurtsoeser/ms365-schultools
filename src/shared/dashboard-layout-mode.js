/**
 * Ein Modus für das Dashboard-Layout (eine Quelle der Wahrheit).
 *
 * full – Global Admin / Betreiber: kompletter Katalog
 * preview-lehrer | preview-schueler – IT-Vorschau (nur 2 Apps)
 * apps-lehrer | apps-schueler – echte Lehrkraft / Schüler/in (nur 2 Apps)
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
 * @param {import('./dashboard-audience-resolve.js').DashboardPersonas} personas
 * @returns {DashboardLayoutMode}
 */
export function resolveLayoutModeFromPersonas(personas) {
    if (!personas || !personas.loggedIn) return 'full';

    if (personas.isIt) {
        const p = readItPreviewChoice();
        if (p === 'lehrer') return 'preview-lehrer';
        if (p === 'schueler') return 'preview-schueler';
        return 'full';
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

export function readItPreviewChoice() {
    try {
        const v = String(localStorage.getItem(itPreviewStorageKey()) || '').toLowerCase();
        if (v === 'lehrer' || v === 'schueler') return v;
    } catch {
        /* ignore */
    }
    return '';
}

/**
 * @param {'full'|'preview-lehrer'|'preview-schueler'} mode
 */
export function writeItPreviewChoice(mode) {
    try {
        const storageKey = itPreviewStorageKey();
        if (mode === 'full') localStorage.removeItem(storageKey);
        else if (mode === 'preview-lehrer') localStorage.setItem(storageKey, 'lehrer');
        else if (mode === 'preview-schueler') localStorage.setItem(storageKey, 'schueler');
    } catch {
        /* ignore */
    }
}

/**
 * @param {DashboardLayoutMode} mode
 */
export function layoutModeLabel(mode) {
    if (mode === 'full') return 'Vollzugriff (Schul-IT)';
    if (mode === 'preview-lehrer') return 'Vorschau Lehrkraft';
    if (mode === 'preview-schueler') return 'Vorschau Schüler/in';
    if (mode === 'apps-lehrer') return 'Lehrkraft';
    return 'Schüler/in';
}
