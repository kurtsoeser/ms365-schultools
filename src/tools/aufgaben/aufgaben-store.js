/**
 * localStorage-Store Aufgaben (local-first, Actor beim Schreiben).
 */

import { buildAufgabenDemoState } from './aufgaben-demo.js';
import { validateAufgabe } from './aufgaben-logic.js';
import {
    DEMO_SEED_TAG,
    STATUS,
    emptyState,
    isDemoSeedItem,
    normalizeAufgabe,
    normalizeState,
    STORAGE_KEY,
    toIsoDateOnly
} from './aufgaben-model.js';

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

/**
 * Nur bei leerem Store und explizitem Opt-in (Caller prüft ?demo=1).
 */
export function seedDemoIfEmpty() {
    const state = loadState();
    if ((state.aufgaben || []).length) return state;
    return saveState(buildAufgabenDemoState());
}

export function loadDemoData() {
    return saveState(buildAufgabenDemoState());
}

/**
 * Nur Seed-Zeilen entfernen (echte Aufgaben bleiben).
 */
export function clearDemoSeed() {
    const state = loadState();
    state.aufgaben = (state.aufgaben || []).filter((a) => !isDemoSeedItem(a));
    if (!(state.aufgaben || []).length) state.localDemoOnly = false;
    return saveState(state);
}

export function clearAll() {
    return saveState(emptyState());
}

/**
 * @param {object} draft
 * @param {{ actorName?: string, actorEmail?: string }} [opts]
 */
export function upsertAufgabe(draft, opts = {}) {
    const v = validateAufgabe(draft);
    if (!v.ok) return { ok: false, errors: v.errors, item: null, state: loadState() };
    const state = loadState();
    const item = normalizeAufgabe(draft);
    if (!item.erstelltVonEmail && opts.actorEmail) {
        item.erstelltVonEmail = String(opts.actorEmail).trim().toLowerCase();
    }
    if (!item.erstelltVon && opts.actorName) {
        item.erstelltVon = String(opts.actorName).trim();
    }
    const idx = state.aufgaben.findIndex((x) => x.id === item.id);
    if (idx >= 0) {
        state.aufgaben[idx] = {
            ...state.aufgaben[idx],
            ...item,
            erstelltAm: state.aufgaben[idx].erstelltAm || item.erstelltAm,
            erstelltVon: state.aufgaben[idx].erstelltVon || item.erstelltVon,
            erstelltVonEmail: state.aufgaben[idx].erstelltVonEmail || item.erstelltVonEmail,
            seedTag: state.aufgaben[idx].seedTag || item.seedTag
        };
    } else {
        state.aufgaben.unshift(item);
    }
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 * @param {string} status
 * @param {{ kommentar?: string, actorName?: string, actorEmail?: string }} [opts]
 */
export function setAufgabeStatus(id, status, opts = {}) {
    const state = loadState();
    const item = state.aufgaben.find((x) => x.id === String(id || '').trim());
    if (!item) return { ok: false, errors: ['Aufgabe nicht gefunden.'], state };
    const nextStatus = String(status || '')
        .trim()
        .toLowerCase();
    if (!Object.values(STATUS).includes(nextStatus)) {
        return { ok: false, errors: ['Status ungültig.'], state };
    }
    item.status = nextStatus;
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
    const item = state.aufgaben.find((x) => x.id === String(id || '').trim());
    if (!item) return { ok: false, errors: ['Aufgabe nicht gefunden.'], state };
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
export function removeAufgabe(id) {
    const state = loadState();
    const aid = String(id || '').trim();
    state.aufgaben = state.aufgaben.filter((x) => x.id !== aid);
    return { ok: true, errors: [], state: saveState(state) };
}

export { DEMO_SEED_TAG, isDemoSeedItem };
