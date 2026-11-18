/**
 * localStorage-Store Kalender und Termine.
 */

import { buildTermineHubDemoState } from './termine-hub-demo.js';
import { validateTermin } from './termine-hub-logic.js';
import {
    SYNC_STATUS,
    emptyState,
    normalizeState,
    normalizeTermin,
    STORAGE_KEY
} from './termine-hub-model.js';

function storage() {
    if (typeof localStorage === 'undefined') return null;
    return localStorage;
}

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

export function saveState(state) {
    const next = normalizeState(state);
    next.updatedAt = new Date().toISOString();
    const ls = storage();
    if (ls) ls.setItem(STORAGE_KEY, JSON.stringify(next));
    return next;
}

export function seedDemoIfEmpty() {
    const state = loadState();
    if ((state.termine || []).length) return state;
    return saveState(buildTermineHubDemoState());
}

export function loadDemoData() {
    return saveState(buildTermineHubDemoState());
}

export function clearAll() {
    return saveState(emptyState());
}

/**
 * @param {object} draft
 */
export function upsertTermin(draft) {
    const check = validateTermin(draft);
    if (!check.ok) return { ok: false, errors: check.errors, item: null, state: loadState() };
    const state = loadState();
    const item = normalizeTermin({
        ...check.item,
        updatedAt: new Date().toISOString()
    });
    const idx = state.termine.findIndex((t) => t.id === item.id);
    if (idx >= 0) state.termine[idx] = { ...state.termine[idx], ...item };
    else state.termine.unshift(item);
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 */
export function removeTermin(id) {
    const state = loadState();
    const tid = String(id || '').trim();
    if (!state.termine.some((t) => t.id === tid)) {
        return { ok: false, errors: ['Termin nicht gefunden.'], state };
    }
    state.termine = state.termine.filter((t) => t.id !== tid);
    return { ok: true, errors: [], state: saveState(state) };
}

/**
 * Für Outlook vormerken (Graph/PA folgt).
 * @param {string} id
 */
export function markOutlookPending(id) {
    const state = loadState();
    const item = state.termine.find((t) => t.id === String(id || '').trim());
    if (!item) return { ok: false, errors: ['Termin nicht gefunden.'], item: null, state };
    item.syncStatus = SYNC_STATUS.PENDING;
    item.updatedAt = new Date().toISOString();
    return {
        ok: true,
        errors: [],
        item,
        state: saveState(state),
        message:
            'Für Outlook vorgemerkt. Graph-Schreiben bzw. Power Automate folgt – noch kein Kalender-Eintrag erstellt.'
    };
}
