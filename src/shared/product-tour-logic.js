/**
 * Persistenz und Schritt-Logik für First-Run / Product Tours (ohne DOM).
 */

export const ONBOARDING_WELCOME_KEY = 'ms365-onboarding-welcome-v1';
export const IT_FIRST_RUN_KEY = 'ms365-it-first-run-v1';
export const LEHRER_WELCOME_KEY = 'ms365-lehrer-welcome-v1';
export const LEHRER_FIRST_RUN_KEY = 'ms365-lehrer-first-run-v1';

/** @typedef {'idle'|'welcome'|'dashboard'|'stammdaten'|'done'|'skipped'} ItFirstRunPhase */
/** @typedef {'idle'|'welcome'|'dashboard'|'done'|'skipped'} LehrerFirstRunPhase */

/**
 * @param {unknown} raw
 * @returns {{ v: number, phase: ItFirstRunPhase, step: number, updatedAt: number }}
 */
export function normalizeItFirstRunState(raw) {
    const base = {
        v: 1,
        phase: /** @type {ItFirstRunPhase} */ ('idle'),
        step: 0,
        updatedAt: 0
    };
    if (!raw || typeof raw !== 'object') return base;
    const o = /** @type {Record<string, unknown>} */ (raw);
    const phase = String(o.phase || '').trim();
    const allowed = new Set(['idle', 'welcome', 'dashboard', 'stammdaten', 'done', 'skipped']);
    return {
        v: 1,
        phase: /** @type {ItFirstRunPhase} */ (allowed.has(phase) ? phase : 'idle'),
        step: Math.max(0, Number(o.step) || 0),
        updatedAt: Math.max(0, Number(o.updatedAt) || 0)
    };
}

/**
 * @param {Storage | null | undefined} store
 * @param {string} key
 */
export function readTourFlag(store, key) {
    try {
        if (!store || typeof store.getItem !== 'function') return '';
        return String(store.getItem(key) || '').trim();
    } catch {
        return '';
    }
}

/**
 * @param {Storage | null | undefined} store
 * @param {string} key
 * @param {string} value
 */
export function writeTourFlag(store, key, value) {
    try {
        if (!store || typeof store.setItem !== 'function') return false;
        store.setItem(key, String(value || ''));
        return true;
    } catch {
        return false;
    }
}

/**
 * @param {Storage | null | undefined} store
 */
export function isWelcomeTourDone(store) {
    const v = readTourFlag(store, ONBOARDING_WELCOME_KEY).toLowerCase();
    return v === 'seen' || v === 'skipped' || v === '1' || v === 'true';
}

/**
 * @param {Storage | null | undefined} store
 * @param {'seen'|'skipped'} [status]
 */
export function markWelcomeTourDone(store, status) {
    return writeTourFlag(store, ONBOARDING_WELCOME_KEY, status === 'skipped' ? 'skipped' : 'seen');
}

/**
 * @param {Storage | null | undefined} store
 */
export function readItFirstRunState(store) {
    try {
        if (!store || typeof store.getItem !== 'function') return normalizeItFirstRunState(null);
        const raw = store.getItem(IT_FIRST_RUN_KEY);
        if (!raw) return normalizeItFirstRunState(null);
        return normalizeItFirstRunState(JSON.parse(raw));
    } catch {
        return normalizeItFirstRunState(null);
    }
}

/**
 * @param {Storage | null | undefined} store
 * @param {{ phase?: ItFirstRunPhase, step?: number }} patch
 */
export function patchItFirstRunState(store, patch) {
    const cur = readItFirstRunState(store);
    const next = normalizeItFirstRunState({
        ...cur,
        ...(patch || {}),
        updatedAt: Date.now()
    });
    try {
        if (!store || typeof store.setItem !== 'function') return next;
        store.setItem(IT_FIRST_RUN_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

/**
 * Soll die Willkommens-Tour auf dem Dashboard starten?
 * @param {{ welcomeDone?: boolean, isIt?: boolean, loggedIn?: boolean, force?: boolean, phase?: ItFirstRunPhase }} opts
 */
export function shouldStartItWelcomeTour(opts) {
    const o = opts || {};
    if (o.force) return true;
    if (!o.loggedIn || !o.isIt) return false;
    if (o.welcomeDone) return false;
    const phase = o.phase || 'idle';
    if (phase === 'done' || phase === 'skipped' || phase === 'stammdaten') return false;
    return true;
}

/**
 * Soll die Stammdaten-Fortsetzung starten?
 * @param {{ isIt?: boolean, loggedIn?: boolean, force?: boolean, phase?: ItFirstRunPhase }} opts
 */
export function shouldStartItStammdatenTour(opts) {
    const o = opts || {};
    if (o.force) return true;
    if (!o.loggedIn || !o.isIt) return false;
    return (o.phase || 'idle') === 'stammdaten';
}

/**
 * Fortschritts-Chip: offene oder abgebrochene Einrichtung (nicht bei idle/done).
 * @param {{ isIt?: boolean, loggedIn?: boolean, phase?: ItFirstRunPhase, chipDismissed?: boolean }} opts
 */
export function shouldShowItFirstRunChip(opts) {
    const o = opts || {};
    if (o.chipDismissed) return false;
    if (!o.loggedIn || !o.isIt) return false;
    const phase = o.phase || 'idle';
    return phase === 'welcome' || phase === 'dashboard' || phase === 'stammdaten' || phase === 'skipped';
}

/**
 * @param {ItFirstRunPhase | string} phase
 * @returns {'stammdaten'|'dashboard-restart'}
 */
export function resumeTargetFromPhase(phase) {
    return String(phase || '') === 'stammdaten' ? 'stammdaten' : 'dashboard-restart';
}

/**
 * @param {ItFirstRunPhase | string} phase
 */
export function chipLabelFromPhase(phase) {
    const p = String(phase || '');
    if (p === 'stammdaten') return 'Stammdaten-Einrichtung fortsetzen';
    if (p === 'skipped') return 'IT-Tour starten';
    if (p === 'welcome' || p === 'dashboard') return 'IT-Tour fortsetzen';
    return 'IT-Tour starten';
}

export const IT_FIRST_RUN_CHIP_DISMISS_KEY = 'ms365-it-first-run-chip-dismiss-v1';
export const LEHRER_FIRST_RUN_CHIP_DISMISS_KEY = 'ms365-lehrer-first-run-chip-dismiss-v1';
/** Nach erstem manuellem „Speichern“: Hinweis lokal + Sicherungsbibliothek. */
export const IT_FIRST_SAVE_TIP_KEY = 'ms365-it-first-save-tip-v1';

/**
 * @param {unknown} raw
 * @returns {{ v: number, phase: LehrerFirstRunPhase, step: number, updatedAt: number }}
 */
export function normalizeLehrerFirstRunState(raw) {
    const base = {
        v: 1,
        phase: /** @type {LehrerFirstRunPhase} */ ('idle'),
        step: 0,
        updatedAt: 0
    };
    if (!raw || typeof raw !== 'object') return base;
    const o = /** @type {Record<string, unknown>} */ (raw);
    const phase = String(o.phase || '').trim();
    const allowed = new Set(['idle', 'welcome', 'dashboard', 'done', 'skipped']);
    return {
        v: 1,
        phase: /** @type {LehrerFirstRunPhase} */ (allowed.has(phase) ? phase : 'idle'),
        step: Math.max(0, Number(o.step) || 0),
        updatedAt: Math.max(0, Number(o.updatedAt) || 0)
    };
}

/**
 * @param {Storage | null | undefined} store
 */
export function isLehrerWelcomeTourDone(store) {
    const v = readTourFlag(store, LEHRER_WELCOME_KEY).toLowerCase();
    return v === 'seen' || v === 'skipped' || v === '1' || v === 'true';
}

/**
 * @param {Storage | null | undefined} store
 * @param {'seen'|'skipped'} [status]
 */
export function markLehrerWelcomeTourDone(store, status) {
    return writeTourFlag(store, LEHRER_WELCOME_KEY, status === 'skipped' ? 'skipped' : 'seen');
}

/**
 * @param {Storage | null | undefined} store
 */
export function readLehrerFirstRunState(store) {
    try {
        if (!store || typeof store.getItem !== 'function') return normalizeLehrerFirstRunState(null);
        const raw = store.getItem(LEHRER_FIRST_RUN_KEY);
        if (!raw) return normalizeLehrerFirstRunState(null);
        return normalizeLehrerFirstRunState(JSON.parse(raw));
    } catch {
        return normalizeLehrerFirstRunState(null);
    }
}

/**
 * @param {Storage | null | undefined} store
 * @param {{ phase?: LehrerFirstRunPhase, step?: number }} patch
 */
export function patchLehrerFirstRunState(store, patch) {
    const cur = readLehrerFirstRunState(store);
    const next = normalizeLehrerFirstRunState({
        ...cur,
        ...(patch || {}),
        updatedAt: Date.now()
    });
    try {
        if (!store || typeof store.setItem !== 'function') return next;
        store.setItem(LEHRER_FIRST_RUN_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

/**
 * Lehrer-Willkommens-Tour: angemeldete Lehrkraft (nicht IT-Backend).
 * Bei setup_incomplete (portalReady === false) kein Auto-Start – Banner zuerst.
 * @param {{ welcomeDone?: boolean, isLehrer?: boolean, isIt?: boolean, loggedIn?: boolean, force?: boolean, phase?: LehrerFirstRunPhase, portalReady?: boolean }} opts
 */
export function shouldStartLehrerWelcomeTour(opts) {
    const o = opts || {};
    if (o.force) return true;
    if (!o.loggedIn || !o.isLehrer) return false;
    // Schul-IT bekommt die IT-Tour; Lehrkräfte das Portal.
    if (o.isIt) return false;
    if (o.welcomeDone) return false;
    const phase = o.phase || 'idle';
    if (phase === 'done' || phase === 'skipped') return false;
    // R13: leerer Erststart – Tour wartet, bis Portal bereit ist (Menü-Start bleibt via force).
    if (o.portalReady === false) return false;
    return true;
}

/**
 * @param {{ isLehrer?: boolean, isIt?: boolean, loggedIn?: boolean, phase?: LehrerFirstRunPhase, chipDismissed?: boolean }} opts
 */
export function shouldShowLehrerFirstRunChip(opts) {
    const o = opts || {};
    if (o.chipDismissed) return false;
    if (!o.loggedIn || !o.isLehrer || o.isIt) return false;
    const phase = o.phase || 'idle';
    return phase === 'welcome' || phase === 'dashboard' || phase === 'skipped';
}

/**
 * @param {LehrerFirstRunPhase | string} phase
 */
export function lehrerChipLabelFromPhase(phase) {
    const p = String(phase || '');
    if (p === 'skipped') return 'Lehrer-Tour starten';
    if (p === 'welcome' || p === 'dashboard') return 'Lehrer-Tour fortsetzen';
    return 'Lehrer-Tour starten';
}

/**
 * Menüpunkt „Lehrer-Tour erneut“: echte Lehrkraft oder IT in Portal-Vorschau.
 * @param {{ loggedIn?: boolean, isLehrer?: boolean, isIt?: boolean, portalPreview?: boolean }} opts
 */
export function shouldShowLehrerTourMenuButton(opts) {
    const o = opts || {};
    if (!o.loggedIn) return false;
    if (o.isLehrer && !o.isIt) return true;
    if (o.isIt && o.portalPreview) return true;
    return false;
}

/**
 * @param {Storage | null | undefined} store
 */
export function isFirstSaveTipDone(store) {
    const v = readTourFlag(store, IT_FIRST_SAVE_TIP_KEY).toLowerCase();
    return v === 'seen' || v === 'skipped' || v === '1' || v === 'true';
}

/**
 * @param {Storage | null | undefined} store
 * @param {'seen'|'skipped'} [status]
 */
export function markFirstSaveTipDone(store, status) {
    return writeTourFlag(store, IT_FIRST_SAVE_TIP_KEY, status === 'skipped' ? 'skipped' : 'seen');
}

/**
 * Erster Speichern-/Datensicherungs-Klick → Kurz-Tour (lokal + Sicherungsbibliothek).
 * @param {{ tipDone?: boolean, reason?: string, force?: boolean, isTenantPage?: boolean }} opts
 */
export function shouldShowFirstSaveTip(opts) {
    const o = opts || {};
    if (o.force) return true;
    if (o.tipDone) return false;
    const reason = String(o.reason || '');
    // Floppy oben (Datensicherung): auf Dashboard und Stammdaten
    if (reason === 'backup-header') return true;
    if (o.isTenantPage === false) return false;
    return reason === 'manual-save';
}

/**
 * @param {number} index
 * @param {number} total
 */
export function clampTourStep(index, total) {
    const n = Math.max(0, Number(total) || 0);
    if (!n) return 0;
    const i = Math.max(0, Number(index) || 0);
    return Math.min(i, n - 1);
}

/**
 * @param {number} index
 * @param {number} total
 */
export function isLastTourStep(index, total) {
    const n = Math.max(0, Number(total) || 0);
    if (!n) return true;
    return clampTourStep(index, n) >= n - 1;
}
