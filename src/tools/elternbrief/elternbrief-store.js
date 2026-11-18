/**
 * localStorage-Store Elternbrief.
 */

import { buildElternbriefDemoState, ELTERNBRIEF_DEMO_SEED_TAG } from './elternbrief-demo.js';
import {
    parseEmpfaengerText,
    parseKlasseCodes,
    validateBrief
} from './elternbrief-logic.js';
import {
    LESE_STATUS,
    STATUS,
    emptyState,
    isDemoSeedItem,
    normalizeBrief,
    normalizeEmpfaenger,
    normalizeState,
    STORAGE_KEY
} from './elternbrief-model.js';
import {
    buildEmpfaengerLinesForKlassen,
    loadEinsammlerStammdaten
} from '../einsammler/einsammler-stammdaten.js';

function storage() {
    if (typeof localStorage === 'undefined') return null;
    return localStorage;
}

function isDemoRow(row) {
    return isDemoSeedItem(row);
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
    return loadState();
}

/**
 * @param {{ email?: string, name?: string }} [owner]
 */
export function loadDemoData(owner) {
    const demo = buildElternbriefDemoState();
    const email = String((owner && owner.email) || '')
        .trim()
        .toLowerCase();
    const name = String((owner && owner.name) || email).trim();
    if (email) {
        demo.briefe = (demo.briefe || []).map((b) => ({
            ...b,
            erstellerEmail: email,
            erstellerName: name,
            absenderEmail: b.absenderEmail || email,
            absenderName: b.absenderName || name,
            quelle: 'demo',
            seedTag: ELTERNBRIEF_DEMO_SEED_TAG
        }));
    }
    demo.empfaenger = (demo.empfaenger || []).map((e) => ({
        ...e,
        quelle: 'demo',
        seedTag: ELTERNBRIEF_DEMO_SEED_TAG
    }));
    const current = loadState();
    const keptBriefe = (current.briefe || []).filter((b) => !isDemoRow(b));
    const keptEmp = (current.empfaenger || []).filter((e) => !isDemoRow(e));
    return saveState({
        briefe: [...demo.briefe, ...keptBriefe],
        empfaenger: [...demo.empfaenger, ...keptEmp],
        updatedAt: new Date().toISOString()
    });
}

/** Nur Demo-Seed-Zeilen entfernen (echte Briefe bleiben). */
export function clearDemoSeeds() {
    const state = loadState();
    return saveState({
        briefe: (state.briefe || []).filter((b) => !isDemoRow(b)),
        empfaenger: (state.empfaenger || []).filter((e) => !isDemoRow(e)),
        updatedAt: new Date().toISOString()
    });
}

export function clearAll() {
    return saveState(emptyState());
}

/**
 * @param {object} draft
 */
export function upsertBrief(draft) {
    const v = validateBrief(draft);
    if (!v.ok) return { ok: false, errors: v.errors, item: null, state: loadState() };
    const state = loadState();
    const item = normalizeBrief(draft);
    const idx = state.briefe.findIndex((x) => x.id === item.id);
    if (idx >= 0) {
        const prev = state.briefe[idx];
        const merged = {
            ...prev,
            ...item,
            spoItemId: item.spoItemId || prev.spoItemId || ''
        };
        state.briefe[idx] = merged;
        return { ok: true, errors: [], item: merged, state: saveState(state) };
    }
    state.briefe.unshift(item);
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * Empfänger upsert (für SPO-Sync spoItemId merken).
 * @param {object} draft
 */
export function upsertEmpfaenger(draft) {
    const state = loadState();
    const item = normalizeEmpfaenger(draft);
    if (!item.id || !item.briefId) {
        return { ok: false, errors: ['Empfänger unvollständig.'], item: null, state };
    }
    const idx = state.empfaenger.findIndex((x) => x.id === item.id);
    if (idx >= 0) {
        const prev = state.empfaenger[idx];
        const merged = {
            ...prev,
            ...item,
            spoItemId: item.spoItemId || prev.spoItemId || ''
        };
        state.empfaenger[idx] = merged;
        return { ok: true, errors: [], item: merged, state: saveState(state) };
    }
    state.empfaenger.push(item);
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 * @param {string} status
 */
export function setBriefStatus(id, status) {
    const state = loadState();
    const item = state.briefe.find((x) => x.id === String(id || '').trim());
    if (!item) return { ok: false, errors: ['Brief nicht gefunden.'], state };
    item.status = String(status || '').toLowerCase();
    if (item.status === STATUS.GESENDET && !item.gesendetAm) {
        item.gesendetAm = new Date().toISOString();
    }
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 */
export function removeBrief(id) {
    const state = loadState();
    const bid = String(id || '').trim();
    const removed = state.briefe.find((b) => b.id === bid) || null;
    const removedEmp = state.empfaenger.filter((e) => e.briefId === bid);
    state.briefe = state.briefe.filter((b) => b.id !== bid);
    state.empfaenger = state.empfaenger.filter((e) => e.briefId !== bid);
    const next = saveState(state);
    return { state: next, brief: removed, empfaenger: removedEmp };
}

/**
 * Empfänger aus Text und/oder Stammdaten-Klassen.
 * @param {string} briefId
 * @param {{ text?: string, fromKlassen?: boolean }} opts
 */
export function addEmpfaenger(briefId, opts = {}) {
    const state = loadState();
    const brief = state.briefe.find((b) => b.id === String(briefId || '').trim());
    if (!brief) return { ok: false, errors: ['Brief nicht gefunden.'], state, added: 0 };

    /** @type {ReturnType<typeof normalizeEmpfaenger>[]} */
    let rows = [];
    if (opts.fromKlassen) {
        const codes = parseKlasseCodes(brief.klasseCodes);
        const lines = buildEmpfaengerLinesForKlassen(codes, loadEinsammlerStammdaten().students);
        rows = parseEmpfaengerText(lines.join('\n'));
    }
    if (opts.text) rows = rows.concat(parseEmpfaengerText(opts.text));

    let added = 0;
    for (const row of rows) {
        const exists = state.empfaenger.some(
            (e) =>
                e.briefId === brief.id &&
                e.schuelerName.toLowerCase() === row.schuelerName.toLowerCase() &&
                e.klasseCode.toLowerCase() === row.klasseCode.toLowerCase()
        );
        if (exists) continue;
        state.empfaenger.push(
            normalizeEmpfaenger({
                ...row,
                briefId: brief.id,
                status: LESE_STATUS.OFFEN
            })
        );
        added++;
    }
    return { ok: true, errors: [], added, state: saveState(state) };
}

/**
 * @param {string} empId
 * @param {string} status
 */
export function setLeseStatus(empId, status) {
    const state = loadState();
    const item = state.empfaenger.find((x) => x.id === String(empId || '').trim());
    if (!item) return { ok: false, errors: ['Empfänger nicht gefunden.'], state };
    item.status = String(status || '').toLowerCase();
    if (item.status === LESE_STATUS.GELESEN) {
        item.gelesenAm = new Date().toISOString();
    } else if (item.status === LESE_STATUS.OFFEN) {
        item.gelesenAm = '';
    }
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} briefId
 * @param {string} empId
 */
export function markGelesen(briefId, empId) {
    const state = loadState();
    const item = state.empfaenger.find(
        (x) => x.id === String(empId || '').trim() && x.briefId === String(briefId || '').trim()
    );
    if (!item) return { ok: false, errors: ['Eintrag nicht gefunden.'], state };
    item.status = LESE_STATUS.GELESEN;
    item.gelesenAm = new Date().toISOString();
    return { ok: true, errors: [], item, state: saveState(state) };
}

export function getBrief(id) {
    return loadState().briefe.find((b) => b.id === String(id || '').trim()) || null;
}

export function listEmpfaenger(briefId) {
    return loadState().empfaenger.filter((e) => e.briefId === String(briefId || '').trim());
}
