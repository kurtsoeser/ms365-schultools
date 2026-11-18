/**
 * localStorage-Store Technik-Meldungen (local-first).
 */

import { buildTechnikDemoState } from './technik-meldungen-demo.js';
import { validateMeldung } from './technik-meldungen-logic.js';
import {
    DEMO_SEED_TAG,
    STATUS,
    emptyState,
    isDemoSeedItem,
    normalizeMeldung,
    normalizeState,
    STORAGE_KEY,
    toIsoDateOnly
} from './technik-meldungen-model.js';

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
    if ((state.meldungen || []).length) return state;
    return saveState(buildTechnikDemoState());
}

export function loadDemoData() {
    return saveState(buildTechnikDemoState());
}

/**
 * Nur Seed-Zeilen entfernen (echte Meldungen bleiben).
 */
export function clearDemoSeed() {
    const state = loadState();
    state.meldungen = (state.meldungen || []).filter((m) => !isDemoSeedItem(m));
    if (!(state.meldungen || []).length) state.localDemoOnly = false;
    return saveState(state);
}

export function clearAll() {
    return saveState(emptyState());
}

/**
 * @param {object} draft
 * @param {{ actorName?: string, actorEmail?: string }} [opts]
 */
export function upsertMeldung(draft, opts = {}) {
    const v = validateMeldung(draft);
    if (!v.ok) return { ok: false, errors: v.errors, item: null, state: loadState() };
    const state = loadState();
    const item = normalizeMeldung(draft);
    if (!item.melderEmail && opts.actorEmail) {
        item.melderEmail = String(opts.actorEmail).trim().toLowerCase();
    }
    if (!item.melderName && opts.actorName) {
        item.melderName = String(opts.actorName).trim();
    }
    const idx = state.meldungen.findIndex((x) => x.id === item.id);
    if (idx >= 0) {
        state.meldungen[idx] = {
            ...state.meldungen[idx],
            ...item,
            gemeldetAm: state.meldungen[idx].gemeldetAm || item.gemeldetAm,
            seedTag: state.meldungen[idx].seedTag || item.seedTag
        };
    } else {
        state.meldungen.unshift(item);
    }
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 * @param {string} status
 * @param {{ kommentar?: string, actorName?: string, actorEmail?: string }} [opts]
 */
export function setMeldungStatus(id, status, opts = {}) {
    const state = loadState();
    const item = state.meldungen.find((x) => x.id === String(id || '').trim());
    if (!item) return { ok: false, errors: ['Meldung nicht gefunden.'], state };
    const nextStatus = String(status || '')
        .trim()
        .toLowerCase();
    if (!Object.values(STATUS).includes(nextStatus)) {
        return { ok: false, errors: ['Status ungültig.'], state };
    }
    item.status = nextStatus;
    if (nextStatus === STATUS.ERLEDIGT && !item.erledigtAm) {
        item.erledigtAm = toIsoDateOnly(new Date());
    }
    if (nextStatus === STATUS.OFFEN || nextStatus === STATUS.IN_ARBEIT) {
        item.erledigtAm = '';
    }
    if (opts.kommentar != null) item.kommentar = String(opts.kommentar).trim();
    const actorEmail = String(opts.actorEmail || '')
        .trim()
        .toLowerCase();
    const actorName = String(opts.actorName || '').trim();
    if (actorEmail || actorName) {
        item.bearbeitetVonEmail = actorEmail || item.bearbeitetVonEmail;
        item.bearbeitetVonName = actorName || item.bearbeitetVonName;
        item.bearbeitetAm = new Date().toISOString();
    }
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 * @param {string} kommentar
 * @param {{ actorName?: string, actorEmail?: string }} [opts]
 */
export function appendKommentar(id, kommentar, opts = {}) {
    const state = loadState();
    const item = state.meldungen.find((x) => x.id === String(id || '').trim());
    if (!item) return { ok: false, errors: ['Meldung nicht gefunden.'], state };
    const text = String(kommentar || '').trim();
    if (!text) return { ok: false, errors: ['Kommentar leer.'], state };
    const stamp = toIsoDateOnly(new Date());
    const who = String(opts.actorName || opts.actorEmail || '').trim();
    const line = who ? '[' + stamp + ' · ' + who + '] ' + text : '[' + stamp + '] ' + text;
    item.kommentar = item.kommentar ? item.kommentar + '\n' + line : line;
    const actorEmail = String(opts.actorEmail || '')
        .trim()
        .toLowerCase();
    if (actorEmail) {
        item.bearbeitetVonEmail = actorEmail;
        item.bearbeitetVonName = String(opts.actorName || '').trim() || item.bearbeitetVonName;
        item.bearbeitetAm = new Date().toISOString();
    }
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 */
export function removeMeldung(id) {
    const state = loadState();
    const mid = String(id || '').trim();
    state.meldungen = state.meldungen.filter((x) => x.id !== mid);
    return { ok: true, errors: [], state: saveState(state) };
}

export { DEMO_SEED_TAG, isDemoSeedItem };
