/**
 * localStorage-Store Kollegiumsumfragen.
 */

import { buildKollegiumDemoState } from './kollegium-umfragen-demo.js';
import {
    parseLehrerListe,
    parseMails,
    progressStats,
    teacherInAudience,
    validateAntwortPayload,
    validateUmfrage
} from './kollegium-umfragen-logic.js';
import {
    ANTWORT_STATUS,
    STATUS,
    emptyState,
    isDemoSeedItem,
    normalizeAntwort,
    normalizeState,
    normalizeUmfrage,
    STORAGE_KEY
} from './kollegium-umfragen-model.js';

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
    if ((state.umfragen || []).length) return state;
    return saveState(buildKollegiumDemoState());
}

export function loadDemoData() {
    return saveState(buildKollegiumDemoState());
}

/**
 * Nur Seed-Zeilen entfernen (echte Umfragen/Antworten bleiben).
 */
export function clearDemoSeed() {
    const state = loadState();
    state.umfragen = (state.umfragen || []).filter((u) => !isDemoSeedItem(u));
    state.antworten = (state.antworten || []).filter((a) => !isDemoSeedItem(a));
    if (!(state.umfragen || []).length && !(state.antworten || []).length) {
        state.localDemoOnly = false;
    }
    return saveState(state);
}

export function clearAll() {
    return saveState(emptyState());
}

/**
 * @param {object} draft
 */
export function upsertUmfrage(draft) {
    const v = validateUmfrage(draft);
    if (!v.ok) return { ok: false, errors: v.errors, item: null, state: loadState() };
    const state = loadState();
    const item = normalizeUmfrage(draft);
    const idx = state.umfragen.findIndex((x) => x.id === item.id);
    if (idx >= 0) state.umfragen[idx] = { ...state.umfragen[idx], ...item };
    else state.umfragen.unshift(item);
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * Antwort ohne Validierung speichern (Sync / Empfänger-Seed).
 * @param {object} draft
 */
export function upsertAntwort(draft) {
    const state = loadState();
    const item = normalizeAntwort(draft);
    if (!item.umfrageId) {
        return { ok: false, errors: ['Umfrage-ID fehlt.'], item: null, state };
    }
    const idx = state.antworten.findIndex(
        (x) =>
            x.id === item.id ||
            (x.umfrageId === item.umfrageId && x.lehrerEmail === item.lehrerEmail)
    );
    if (idx >= 0) {
        state.antworten[idx] = {
            ...state.antworten[idx],
            ...item,
            id: state.antworten[idx].id || item.id
        };
        return { ok: true, errors: [], item: state.antworten[idx], state: saveState(state) };
    }
    state.antworten.unshift(item);
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 * @param {string} status
 */
export function setUmfrageStatus(id, status) {
    const state = loadState();
    const item = state.umfragen.find((x) => x.id === String(id || '').trim());
    if (!item) return { ok: false, errors: ['Umfrage nicht gefunden.'], state };
    item.status = String(status || '').toLowerCase();
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 */
export function removeUmfrage(id) {
    const state = loadState();
    const uid = String(id || '').trim();
    const removed = state.umfragen.find((x) => x.id === uid) || null;
    const antworten = state.antworten.filter((x) => x.umfrageId === uid);
    state.umfragen = state.umfragen.filter((x) => x.id !== uid);
    state.antworten = state.antworten.filter((x) => x.umfrageId !== uid);
    return { state: saveState(state), removed, antworten };
}

/**
 * Empfängerliste anlegen / ergänzen (offen).
 * @param {string} umfrageId
 * @param {string} text Name;E-Mail[;KV]
 */
export function seedEmpfaenger(umfrageId, text) {
    const state = loadState();
    const u = state.umfragen.find((x) => x.id === String(umfrageId || '').trim());
    if (!u) return { ok: false, errors: ['Umfrage nicht gefunden.'], state, addedItems: [] };
    const people = parseLehrerListe(text);
    const mails = parseMails(u.empfaengerMails);
    let added = 0;
    /** @type {ReturnType<typeof normalizeAntwort>[]} */
    const addedItems = [];
    for (const p of people) {
        if (!mails.includes(p.email)) {
            mails.push(p.email);
            u.empfaengerMails = mails.join('; ');
        }
        const exists = state.antworten.find(
            (a) => a.umfrageId === u.id && a.lehrerEmail === p.email
        );
        if (!exists) {
            const row = normalizeAntwort({
                umfrageId: u.id,
                lehrerEmail: p.email,
                lehrerName: p.name,
                status: ANTWORT_STATUS.OFFEN,
                werte: {}
            });
            state.antworten.push(row);
            addedItems.push(row);
            added += 1;
        }
    }
    return { ok: true, errors: [], added, addedItems, umfrage: u, state: saveState(state) };
}

/**
 * @param {object} draft
 */
export function submitAntwort(draft) {
    const state = loadState();
    const a = normalizeAntwort({
        ...draft,
        status: ANTWORT_STATUS.ABGEGEBEN,
        abgegebenAm: new Date().toISOString()
    });
    const u = state.umfragen.find((x) => x.id === a.umfrageId);
    if (!u) return { ok: false, errors: ['Umfrage nicht gefunden.'], state };
    if (u.status !== STATUS.AKTIV) {
        return { ok: false, errors: ['Umfrage ist nicht aktiv.'], state };
    }
    if (!teacherInAudience(u, { email: a.lehrerEmail })) {
        return { ok: false, errors: ['Sie sind nicht in der Zielgruppe.'], state };
    }
    const v = validateAntwortPayload(u, a);
    if (!v.ok) return { ok: false, errors: v.errors, state };

    const idx = state.antworten.findIndex(
        (x) => x.umfrageId === a.umfrageId && x.lehrerEmail === a.lehrerEmail
    );
    if (idx >= 0) {
        state.antworten[idx] = {
            ...state.antworten[idx],
            ...a,
            id: state.antworten[idx].id,
            spoItemId: state.antworten[idx].spoItemId || a.spoItemId
        };
        return { ok: true, errors: [], item: state.antworten[idx], state: saveState(state) };
    }
    state.antworten.unshift(a);
    return { ok: true, errors: [], item: a, state: saveState(state) };
}

/**
 * @param {string} umfrageId
 */
export function getUmfrage(umfrageId) {
    return loadState().umfragen.find((x) => x.id === String(umfrageId || '').trim()) || null;
}

/**
 * @param {string} umfrageId
 */
export function listAntworten(umfrageId) {
    const id = String(umfrageId || '').trim();
    return loadState().antworten.filter((a) => a.umfrageId === id);
}

/**
 * @param {object} umfrage
 */
export function statsFor(umfrage) {
    return progressStats(umfrage, loadState().antworten);
}

export { isDemoSeedItem };
