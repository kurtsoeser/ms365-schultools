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

/**
 * Vor Auth ist oft nur der Tenant-Key gesetzt – Unscoped leer/portal.
 * Dann alle Keys mit Prefix scannen und full bevorzugen (FOUC-/Rail-Schutz).
 * @returns {'portal'|'full'|'schueler'|''}
 */
function scanItShellChoiceFromStorage() {
    try {
        let found = '';
        for (let i = 0; i < localStorage.length; i++) {
            const k = localStorage.key(i);
            if (!k || k.indexOf(IT_PREVIEW_KEY) !== 0) continue;
            const v = String(localStorage.getItem(k) || '')
                .trim()
                .toLowerCase();
            if (v === 'full' || v === 'it') return 'full';
            if (v === 'schueler' && !found) found = 'schueler';
            if ((v === 'portal' || v === 'lehrer') && !found) found = 'portal';
        }
        return found;
    } catch {
        return '';
    }
}

/** @typedef {'full'|'preview-lehrer'|'preview-schueler'|'apps-lehrer'|'apps-schueler'} DashboardLayoutMode */

/**
 * @returns {'portal'|'full'|'schueler'}
 */
export function readItShellChoice() {
    try {
        const scoped = String(localStorage.getItem(itPreviewStorageKey()) || '')
            .trim()
            .toLowerCase();
        if (scoped === 'full' || scoped === 'it') return 'full';
        if (scoped === 'schueler') return 'schueler';

        // Während aktiver IT-Arbeitsfläche: Tenant-full schlägt veraltetes Unscoped-portal
        // (sonst Flash zurück auf Lehrer-Portal beim Tool-Wechsel).
        let inItWorkspace = false;
        try {
            inItWorkspace = sessionStorage.getItem('ms365-it-workspace-v1') === '1';
        } catch {
            inItWorkspace = false;
        }
        if (inItWorkspace || scoped === '' || scoped == null) {
            const scanned = scanItShellChoiceFromStorage();
            if (scanned === 'full') return 'full';
            if (scanned === 'schueler') return 'schueler';
        }

        if (scoped === 'portal' || scoped === 'lehrer') return 'portal';

        const scanned = scanItShellChoiceFromStorage();
        if (scanned === 'full') return 'full';
        if (scanned === 'schueler') return 'schueler';
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
    const choice = readItShellChoice();
    // Bewusst gewählte IT-Arbeitsfläche bleibt – auch während Auth, pending-directory
    // oder wenn die Person zusätzlich als Lehrkraft geführt wird.
    // Sonst landet jeder Tool-Wechsel zuerst auf „Guten Abend“ (Lehrer-Portal).
    if (choice === 'full') return 'full';

    if (!personas || !personas.loggedIn) return 'preview-lehrer';

    if (personas.isIt || personas.globalAdmin || personas.designatedSchoolIt || personas.operator) {
        if (choice === 'schueler') return 'preview-schueler';
        return 'preview-lehrer';
    }

    if (personas.isLehrer) return 'apps-lehrer';
    if (personas.unknownUser) return 'apps-lehrer';
    if (personas.isSchueler) return 'apps-schueler';
    return 'apps-lehrer';
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
        let value = 'portal';
        if (mode === 'full') value = 'full';
        else if (mode === 'preview-schueler') value = 'schueler';
        else if (mode === 'preview-lehrer') value = 'portal';
        const storageKey = itPreviewStorageKey();
        localStorage.setItem(storageKey, value);
        // Unscoped Spiegel für theme-boot / FOUC vor Auth (Tenant-Key oft noch unbekannt)
        if (storageKey !== IT_PREVIEW_KEY) {
            localStorage.setItem(IT_PREVIEW_KEY, value);
        }
    } catch {
        /* ignore */
    }
}

/**
 * @param {DashboardLayoutMode} mode
 */
/**
 * @param {DashboardLayoutMode} mode
 * @param {import('./dashboard-audience-resolve.js').DashboardPersonas|null|undefined} [personas]
 */
export function layoutModeLabel(mode, personas) {
    if (personas && personas.unknownUser) return 'Noch nicht zugeordnet';
    if (mode === 'full') return 'IT-Backend';
    if (mode === 'preview-lehrer') return 'Portal';
    if (mode === 'preview-schueler') return 'Vorschau Schüler/in';
    if (mode === 'apps-lehrer') return 'Lehrkraft';
    return 'Schüler/in';
}
