/**
 * Offen/zu der IT-Bereiche. Eine Quelle für Startseite und Werkzeug-Leiste.
 * Ein Bereichswechsel schreibt nur den betroffenen Bereich, die anderen bleiben.
 */

/** Gleicher Schlüssel wie bisher auf der Startseite. */
export const IT_NAV_GROUP_OPEN_KEY = 'ms365-dash-tasks-nav-groups-v2';
/** „Mehr …“ je Bereich. */
export const IT_NAV_MORE_OPEN_KEY = 'ms365-dash-tasks-nav-more-v1';

/**
 * @param {string} key
 * @returns {Record<string, boolean>}
 */
function readMap(key) {
    try {
        const raw = JSON.parse(localStorage.getItem(key) || '{}');
        return raw && typeof raw === 'object' ? raw : {};
    } catch {
        return {};
    }
}

/**
 * @param {string} key
 * @param {Record<string, boolean>} state
 */
function writeMap(key, state) {
    try {
        localStorage.setItem(key, JSON.stringify(state || {}));
    } catch {
        /* ignore */
    }
}

/** @returns {Record<string, boolean>} */
export function readItNavGroupOpenState() {
    return readMap(IT_NAV_GROUP_OPEN_KEY);
}

/** @param {Record<string, boolean>} state */
export function writeItNavGroupOpenState(state) {
    writeMap(IT_NAV_GROUP_OPEN_KEY, state);
}

/** @returns {Record<string, boolean>} */
export function readItNavMoreOpenState() {
    return readMap(IT_NAV_MORE_OPEN_KEY);
}

/** @param {Record<string, boolean>} state */
export function writeItNavMoreOpenState(state) {
    writeMap(IT_NAV_MORE_OPEN_KEY, state);
}

/**
 * @param {string} taskId
 * @param {boolean} open
 */
export function rememberItNavGroupOpen(taskId, open) {
    const id = String(taskId || '').trim();
    if (!id) return readItNavGroupOpenState();
    const state = readItNavGroupOpenState();
    state[id] = !!open;
    writeItNavGroupOpenState(state);
    return state;
}

/**
 * @param {string} taskId
 * @param {boolean} open
 */
export function rememberItNavMoreOpen(taskId, open) {
    const id = String(taskId || '').trim();
    if (!id) return readItNavMoreOpenState();
    const state = readItNavMoreOpenState();
    state[id] = !!open;
    writeItNavMoreOpenState(state);
    return state;
}

/**
 * Gespeicherter Zustand gewinnt. Fehlt er, ist nur der aktive Bereich offen.
 * Andere Bereiche werden dabei nicht zugeklappt.
 * @param {Record<string, boolean>} state
 * @param {string} taskId
 * @param {boolean} isActive
 * @returns {boolean}
 */
export function itNavGroupShouldOpen(state, taskId, isActive) {
    const id = String(taskId || '');
    if (state && Object.prototype.hasOwnProperty.call(state, id)) {
        return state[id] === true;
    }
    return !!isActive;
}
