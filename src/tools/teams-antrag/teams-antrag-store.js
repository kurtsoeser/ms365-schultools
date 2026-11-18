/**
 * localStorage-Store Teams-Antrag.
 */
import { buildTeamsAntragDemoState } from './teams-antrag-demo.js';
import { buildProvisionPlan, validateAntrag } from './teams-antrag-logic.js';
import {
    STATUS,
    STORAGE_KEY,
    emptyState,
    normalizeAntrag,
    normalizeState
} from './teams-antrag-model.js';

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
    if ((state.antraege || []).length) return state;
    return saveState(buildTeamsAntragDemoState());
}

export function loadDemoData() {
    return saveState(buildTeamsAntragDemoState());
}

export function clearAll() {
    return saveState(emptyState());
}

/**
 * @param {object} draft
 * @param {{ requireApplicant?: boolean }} [opts]
 */
export function upsertAntrag(draft, opts) {
    const v = validateAntrag(draft, opts);
    if (!v.ok) return { ok: false, errors: v.errors, item: null, state: loadState() };
    const state = loadState();
    const item = normalizeAntrag({ ...v.antrag, updatedAt: new Date().toISOString() });
    const idx = state.antraege.findIndex((a) => a.id === item.id);
    if (idx >= 0) state.antraege[idx] = { ...state.antraege[idx], ...item };
    else state.antraege.push(item);
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 * @param {string} status
 * @param {object} [extra]
 */
export function setAntragStatus(id, status, extra) {
    const state = loadState();
    const idx = state.antraege.findIndex((a) => a.id === id);
    if (idx < 0) return { ok: false, errors: ['Antrag nicht gefunden.'], state };
    const patch = { ...(extra || {}), status, updatedAt: new Date().toISOString() };
    if (status === STATUS.GENEHMIGT) patch.genehmigtAt = new Date().toISOString();
    if (status === STATUS.ANGELEGT) patch.angelegtAt = new Date().toISOString();
    state.antraege[idx] = normalizeAntrag({ ...state.antraege[idx], ...patch });
    return { ok: true, errors: [], item: state.antraege[idx], state: saveState(state) };
}

/**
 * @param {string} id
 */
export function removeAntrag(id) {
    const state = loadState();
    state.antraege = state.antraege.filter((a) => a.id !== id);
    return { ok: true, state: saveState(state) };
}

/**
 * @param {string} id
 */
export function getAntrag(id) {
    return loadState().antraege.find((a) => a.id === id) || null;
}

/**
 * Speichert finale Naming-Felder vor Genehmigung.
 * @param {string} id
 * @param {{ finalDisplayName?: string, finalMailNickname?: string, notizIt?: string, besitzerMails?: string, mitgliederMails?: string }} fields
 */
export function patchItFields(id, fields) {
    const state = loadState();
    const idx = state.antraege.findIndex((a) => a.id === id);
    if (idx < 0) return { ok: false, errors: ['Antrag nicht gefunden.'], state };
    state.antraege[idx] = normalizeAntrag({
        ...state.antraege[idx],
        ...(fields || {}),
        updatedAt: new Date().toISOString()
    });
    const plan = buildProvisionPlan(state.antraege[idx]);
    return { ok: true, errors: [], item: state.antraege[idx], plan, state: saveState(state) };
}

export { buildProvisionPlan };
