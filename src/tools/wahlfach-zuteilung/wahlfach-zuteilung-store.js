/**
 * localStorage-Store Wahlfach-Zuteilung.
 */

import { buildWahlfachDemoState } from './wahlfach-zuteilung-demo.js';
import {
    applyManuelleZuteilung,
    validateAnmeldung,
    validateAngebot,
    validateRunde
} from './wahlfach-zuteilung-logic.js';
import {
    emptyState,
    normalizeAnmeldung,
    normalizeAngebot,
    normalizeRunde,
    normalizeState,
    STORAGE_KEY
} from './wahlfach-zuteilung-model.js';
import { matchAnmeldungIndex } from './wahlfach-zuteilung-state.js';

function storage() {
    if (typeof localStorage === 'undefined') return null;
    return localStorage;
}

/**
 * @returns {ReturnType<typeof normalizeState>}
 */
export function loadState() {
    const ls = storage();
    if (!ls) return emptyState();
    try {
        const raw = ls.getItem(STORAGE_KEY);
        if (!raw) return emptyState();
        return normalizeState(JSON.parse(raw));
    } catch {
        return emptyState();
    }
}

/**
 * @param {object} state
 * @returns {ReturnType<typeof normalizeState>}
 */
export function saveState(state) {
    const next = normalizeState(state);
    const ls = storage();
    if (ls) ls.setItem(STORAGE_KEY, JSON.stringify(next));
    return next;
}

/**
 * @param {object} draft
 * @returns {{ ok: boolean, errors: string[], state: object }}
 */
export function upsertRunde(draft) {
    const v = validateRunde(draft);
    if (!v.ok) return { ok: false, errors: v.errors, state: loadState() };
    const state = loadState();
    const prev = state.runde || {};
    state.runde = normalizeRunde({ ...prev, ...draft, id: (draft && draft.id) || prev.id });
    return { ok: true, errors: [], state: saveState(state) };
}

/**
 * @param {object} draft
 * @returns {{ ok: boolean, errors: string[], angebot: object|null, state: object }}
 */
export function upsertAngebot(draft) {
    const v = validateAngebot(draft);
    if (!v.ok) return { ok: false, errors: v.errors, angebot: null, state: loadState() };
    const state = loadState();
    const item = normalizeAngebot(draft);
    const idx = state.angebote.findIndex((a) => a.id === item.id);
    if (idx >= 0) state.angebote[idx] = item;
    else state.angebote.push(item);
    return { ok: true, errors: [], angebot: item, state: saveState(state) };
}

/**
 * @param {string} angebotId
 * @returns {{ ok: boolean, state: object }}
 */
export function removeAngebot(angebotId) {
    const id = String(angebotId || '').trim();
    const state = loadState();
    state.angebote = state.angebote.filter((a) => a.id !== id);
    return { ok: true, state: saveState(state) };
}

/**
 * @param {object} draft
 * @returns {{ ok: boolean, errors: string[], anmeldung: object|null, state: object }}
 */
export function upsertAnmeldung(draft) {
    const state = loadState();
    const v = validateAnmeldung(draft, state.angebote);
    if (!v.ok) return { ok: false, errors: v.errors, anmeldung: null, state };
    const item = normalizeAnmeldung(draft);
    const idx = matchAnmeldungIndex(state.anmeldungen, item);
    if (idx >= 0) {
        const prev = state.anmeldungen[idx];
        item.id = prev.id;
        if (!item.spoItemId && prev.spoItemId) item.spoItemId = prev.spoItemId;
        if (prev.seedTag) item.seedTag = prev.seedTag;
        if (prev.eingereichtAm) item.eingereichtAm = prev.eingereichtAm;
        state.anmeldungen[idx] = item;
    } else state.anmeldungen.push(item);
    return { ok: true, errors: [], anmeldung: item, state: saveState(state) };
}

/**
 * @param {string} anmeldungId
 * @returns {{ ok: boolean, state: object }}
 */
export function removeAnmeldung(anmeldungId) {
    const id = String(anmeldungId || '').trim();
    const state = loadState();
    state.anmeldungen = state.anmeldungen.filter((a) => a.id !== id);
    state.zuteilungen = state.zuteilungen.filter((z) => z.anmeldungId !== id);
    return { ok: true, state: saveState(state) };
}

/**
 * @param {string} anmeldungId
 * @param {string|null} angebotId
 * @param {{ begruendung?: string, zugeteiltVon?: string, force?: boolean }} [meta]
 * @returns {{ ok: boolean, errors: string[], warnings: string[], state: object }}
 */
export function applyZuteilung(anmeldungId, angebotId, meta = {}) {
    const state = loadState();
    const result = applyManuelleZuteilung(
        state,
        anmeldungId,
        angebotId,
        meta,
        state.angebote
    );
    if (!result.ok) {
        return { ok: false, errors: result.errors, warnings: result.warnings, state };
    }
    state.anmeldungen = result.anmeldungen;
    state.zuteilungen = result.zuteilungen;
    return {
        ok: true,
        errors: [],
        warnings: result.warnings,
        state: saveState(state)
    };
}

/**
 * Demo-Daten laden (ersetzt lokalen Store) – nur nach Opt-in.
 * @returns {object}
 */
export function loadDemoData() {
    return saveState(buildWahlfachDemoState());
}

/**
 * Demo nur wenn Store leer (Opt-in-Pfad ?demo=1).
 * Überschreibt keinen vorhandenen produktiven Stand.
 * @returns {object}
 */
export function seedDemoIfEmpty() {
    const state = loadState();
    if (!state.runde && !state.angebote.length && !state.anmeldungen.length) {
        return loadDemoData();
    }
    return state;
}

/**
 * @param {object} state
 * @returns {boolean}
 */
export function hasLocalWzData(state) {
    const s = state || loadState();
    return !!(s.runde || (s.angebote && s.angebote.length) || (s.anmeldungen && s.anmeldungen.length));
}

export { STORAGE_KEY };
