/**
 * localStorage-Store Einsammler (P1: intern, ohne Public-Link).
 */

import { buildEinsammlerDemoState, EINSAMMLER_DEMO_SEED_TAG } from './einsammler-demo.js';
import {
    parseEmpfaengerText,
    validateAntwort,
    validateEinsammler
} from './einsammler-logic.js';
import {
    ANTWORT_STATUS,
    EINSAMMLER_STATUS,
    emptyState,
    isDemoSeedItem,
    normalizeAntwort,
    normalizeEinsammler,
    normalizeState,
    STORAGE_KEY
} from './einsammler-model.js';

function isDemoRow(row) {
    return isDemoSeedItem(row);
}

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
    // Kein Auto-Demo mehr – sonst sehen Lehrkräfte fremde Demo-Daten oder gar nichts Sinnvolles.
    return loadState();
}

/**
 * @param {{ email?: string, name?: string }} [owner] zuordnen an aktuelle Lehrkraft
 */
export function loadDemoData(owner) {
    const demo = buildEinsammlerDemoState();
    const email = String((owner && owner.email) || '')
        .trim()
        .toLowerCase();
    const name = String((owner && owner.name) || email).trim();
    if (email) {
        demo.einsammler = (demo.einsammler || []).map((e) => ({
            ...e,
            erstellerEmail: email,
            erstellerName: name,
            quelle: 'demo',
            seedTag: EINSAMMLER_DEMO_SEED_TAG
        }));
        demo.antworten = (demo.antworten || []).map((a) => ({
            ...a,
            quelle: 'demo',
            seedTag: EINSAMMLER_DEMO_SEED_TAG
        }));
    }
    const current = loadState();
    const keptHeads = (current.einsammler || []).filter((e) => !isDemoRow(e));
    const keptAnt = (current.antworten || []).filter((a) => !isDemoRow(a));
    return saveState({
        einsammler: [...demo.einsammler, ...keptHeads],
        antworten: [...demo.antworten, ...keptAnt],
        localDemoOnly: true
    });
}

/** Nur Demo-Seed-Zeilen entfernen (echte Vorgänge bleiben). */
export function clearDemoSeeds() {
    const state = loadState();
    const next = {
        einsammler: (state.einsammler || []).filter((e) => !isDemoRow(e)),
        antworten: (state.antworten || []).filter((a) => !isDemoRow(a)),
        localDemoOnly: false
    };
    if (!(next.einsammler || []).length && !(next.antworten || []).length) {
        next.localDemoOnly = false;
    }
    return saveState(next);
}

export function clearAll() {
    return saveState(emptyState());
}

/**
 * @param {object} draft
 */
export function upsertEinsammler(draft) {
    const v = validateEinsammler(draft);
    if (!v.ok) return { ok: false, errors: v.errors, item: null, state: loadState() };
    const state = loadState();
    const item = normalizeEinsammler(draft);
    const idx = state.einsammler.findIndex((x) => x.id === item.id);
    if (idx >= 0) {
        state.einsammler[idx] = {
            ...state.einsammler[idx],
            ...item,
            spoItemId: item.spoItemId || state.einsammler[idx].spoItemId || '',
            seedTag: state.einsammler[idx].seedTag || item.seedTag,
            createdAt: state.einsammler[idx].createdAt || item.createdAt
        };
    } else {
        state.einsammler.unshift(item);
    }
    const saved = state.einsammler.find((x) => x.id === item.id) || item;
    return { ok: true, errors: [], item: saved, state: saveState(state) };
}

/**
 * @param {string} id
 * @param {string} status
 */
export function setEinsammlerStatus(id, status) {
    const state = loadState();
    const want = String(id || '').trim();
    const item = state.einsammler.find((x) => x.id === want);
    if (!item) return { ok: false, errors: ['Einsammler nicht gefunden.'], state };
    item.status = String(status || '').trim().toLowerCase() || item.status;
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * Empfänger aus Text anlegen (Status offen). Bestehende Namen derselben ID bleiben.
 * @param {string} einsammlerId
 * @param {string} text
 */
export function addEmpfaengerFromText(einsammlerId, text) {
    const state = loadState();
    const head = state.einsammler.find((x) => x.id === einsammlerId);
    if (!head) {
        return { ok: false, errors: ['Einsammler nicht gefunden.'], state, added: 0, addedItems: [] };
    }
    const firstKlasse = String(head.klasseCodes || '')
        .split(';')
        .map((s) => s.trim())
        .filter(Boolean)[0];
    const parsed = parseEmpfaengerText(text, firstKlasse);
    if (!parsed.length) {
        return { ok: false, errors: ['Keine Empfänger erkannt.'], state, added: 0, addedItems: [] };
    }
    /** @type {object[]} */
    const addedItems = [];
    for (const row of parsed) {
        const exists = state.antworten.some(
            (a) =>
                a.einsammlerId === einsammlerId &&
                a.schuelerName.toLowerCase() === row.schuelerName.toLowerCase() &&
                a.klasseCode.toLowerCase() === row.klasseCode.toLowerCase()
        );
        if (exists) continue;
        const item = normalizeAntwort({
            einsammlerId,
            schuelerName: row.schuelerName,
            klasseCode: row.klasseCode,
            elternEmail: row.elternEmail,
            status: ANTWORT_STATUS.OFFEN,
            quelle: 'ui'
        });
        state.antworten.push(item);
        addedItems.push(item);
    }
    return {
        ok: true,
        errors: [],
        state: saveState(state),
        added: addedItems.length,
        addedItems
    };
}

/**
 * @param {object} draft
 */
export function upsertAntwort(draft) {
    const v = validateAntwort(draft);
    if (!v.ok) return { ok: false, errors: v.errors, item: null, state: loadState() };
    const state = loadState();
    const item = normalizeAntwort(draft);
    if (item.status !== ANTWORT_STATUS.OFFEN && !item.antwortAm) {
        item.antwortAm = new Date().toISOString();
    }
    if (item.status === ANTWORT_STATUS.OFFEN) item.antwortAm = '';
    const idx = state.antworten.findIndex((x) => x.id === item.id);
    if (idx >= 0) {
        state.antworten[idx] = {
            ...state.antworten[idx],
            ...item,
            spoItemId: item.spoItemId || state.antworten[idx].spoItemId || '',
            seedTag: state.antworten[idx].seedTag || item.seedTag
        };
    } else {
        state.antworten.push(item);
    }
    const saved = state.antworten.find((x) => x.id === item.id) || item;
    return { ok: true, errors: [], item: saved, state: saveState(state) };
}

/**
 * @param {string} antwortId
 * @param {string} status
 * @param {string} [kommentar]
 */
export function setAntwortStatus(antwortId, status, kommentar, anhangMeta) {
    const state = loadState();
    const item = state.antworten.find((x) => x.id === antwortId);
    if (!item) return { ok: false, errors: ['Antwort nicht gefunden.'], state };
    item.status = String(status || '').trim().toLowerCase() || item.status;
    if (kommentar != null) item.kommentar = String(kommentar);
    if (anhangMeta && typeof anhangMeta === 'object') {
        if (anhangMeta.anhangName != null) item.anhangName = String(anhangMeta.anhangName || '');
        if (anhangMeta.anhangMime != null) item.anhangMime = String(anhangMeta.anhangMime || '');
        if (anhangMeta.anhangSize != null) item.anhangSize = Number(anhangMeta.anhangSize) || 0;
        if (anhangMeta.anhangKey != null) item.anhangKey = String(anhangMeta.anhangKey || '');
        if (anhangMeta.anhangUrl != null) item.anhangUrl = String(anhangMeta.anhangUrl || '');
    }
    if (item.status === ANTWORT_STATUS.OFFEN) item.antwortAm = '';
    else if (!item.antwortAm) item.antwortAm = new Date().toISOString();
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} antwortId
 */
export function removeAntwort(antwortId) {
    const state = loadState();
    const id = String(antwortId || '').trim();
    const removed = state.antworten.find((a) => a.id === id) || null;
    state.antworten = state.antworten.filter((a) => a.id !== id);
    return { ok: true, item: removed, state: saveState(state) };
}

/**
 * @param {string} einsammlerId
 */
export function removeEinsammler(einsammlerId) {
    const state = loadState();
    const id = String(einsammlerId || '').trim();
    const removed = state.einsammler.find((e) => e.id === id) || null;
    const removedAntworten = state.antworten.filter((a) => a.einsammlerId === id);
    state.einsammler = state.einsammler.filter((e) => e.id !== id);
    state.antworten = state.antworten.filter((a) => a.einsammlerId !== id);
    return { ok: true, item: removed, antworten: removedAntworten, state: saveState(state) };
}

export { EINSAMMLER_STATUS, isDemoSeedItem };
