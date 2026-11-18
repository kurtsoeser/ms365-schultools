/**
 * localStorage-Store Raumplan (local-first; Plan-Assets bleiben lokal).
 */

import { buildRaumplanDemoState } from './raumplan-demo.js';
import { parseRaeumeCsv, validateRaum } from './raumplan-logic.js';
import {
    STORAGE_KEY,
    emptyState,
    isDemoSeedItem,
    normalizeRaum,
    normalizeState
} from './raumplan-model.js';

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
    if ((state.raeume || []).length) return state;
    return saveState(buildRaumplanDemoState());
}

export function loadDemoData() {
    return saveState(buildRaumplanDemoState());
}

/**
 * Nur Seed-Zeilen entfernen (echte Räume bleiben).
 */
export function clearDemoSeed() {
    const state = loadState();
    state.raeume = (state.raeume || []).filter((r) => !isDemoSeedItem(r));
    if (!(state.raeume || []).length) state.localDemoOnly = false;
    return saveState(state);
}

export function clearAll() {
    return saveState(emptyState());
}

/**
 * @param {object} draft
 */
export function upsertRaum(draft) {
    const v = validateRaum(draft);
    if (!v.ok) return { ok: false, errors: v.errors, item: null, state: loadState() };
    const state = loadState();
    const item = normalizeRaum(draft);
    const dup = state.raeume.find(
        (r) => r.code.toLowerCase() === item.code.toLowerCase() && r.id !== item.id
    );
    if (dup) {
        return {
            ok: false,
            errors: ['Code bereits vergeben: ' + dup.code],
            item: null,
            state
        };
    }
    const idx = state.raeume.findIndex((x) => x.id === item.id);
    if (idx >= 0) {
        const prev = state.raeume[idx];
        const merged = {
            ...prev,
            ...item,
            spoItemId: item.spoItemId || prev.spoItemId,
            seedTag: item.seedTag || prev.seedTag
        };
        state.raeume[idx] = merged;
        return { ok: true, errors: [], item: merged, state: saveState(state) };
    }
    state.raeume.unshift(item);
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 */
export function removeRaum(id) {
    const state = loadState();
    const rid = String(id || '').trim();
    const removed = state.raeume.find((r) => r.id === rid) || null;
    state.raeume = state.raeume.filter((r) => r.id !== rid);
    return { ok: true, errors: [], item: removed, state: saveState(state) };
}

/**
 * @param {string} csvText
 */
export function importRaeumeCsv(csvText) {
    const rows = parseRaeumeCsv(csvText);
    let added = 0;
    let updated = 0;
    const errors = [];
    /** @type {object[]} */
    const changed = [];
    for (const row of rows) {
        const state = loadState();
        const existing = state.raeume.find(
            (r) => r.code.toLowerCase() === String(row.code || '').toLowerCase()
        );
        const res = upsertRaum({
            id: existing ? existing.id : undefined,
            spoItemId: existing ? existing.spoItemId : '',
            seedTag: existing && !isDemoSeedItem(existing) ? existing.seedTag : '',
            ...row
        });
        if (res.ok) {
            if (existing) updated += 1;
            else added += 1;
            if (res.item) changed.push(res.item);
        } else errors.push(...(res.errors || []));
    }
    return { ok: errors.length === 0, added, updated, errors, changed, state: loadState() };
}

/**
 * @param {string} dataUrl
 */
export function setPlanImage(dataUrl) {
    const state = loadState();
    state.planImageDataUrl = String(dataUrl || '').trim();
    if (state.planImageDataUrl) state.planPdfName = '';
    return saveState(state);
}

/**
 * @param {string} fileName
 */
export function setPlanPdfName(fileName) {
    const state = loadState();
    state.planPdfName = String(fileName || '').trim();
    if (state.planPdfName) state.planImageDataUrl = '';
    return saveState(state);
}

export function clearPlan() {
    const state = loadState();
    state.planImageDataUrl = '';
    state.planPdfName = '';
    return saveState(state);
}
