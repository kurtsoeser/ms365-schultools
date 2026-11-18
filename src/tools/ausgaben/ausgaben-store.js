/**
 * localStorage-Store Ausgabenverwaltung.
 */

import { buildAusgabenDemoState } from './ausgaben-demo.js';
import {
    findOpenVorgang,
    parseGegenstaendeCsv,
    validateGegenstand,
    validateTyp,
    validateVorgang
} from './ausgaben-logic.js';
import {
    VORGANG_STATUS,
    emptyState,
    isDemoSeedItem,
    normalizeGegenstand,
    normalizeState,
    normalizeTyp,
    normalizeVorgang,
    STORAGE_KEY,
    toIsoDateOnly
} from './ausgaben-model.js';
import { materializeTypVorlage } from './ausgaben-templates.js';

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
    if ((state.gegenstaende || []).length || (state.typen || []).length) return state;
    return saveState(buildAusgabenDemoState());
}

export function loadDemoData() {
    return saveState(buildAusgabenDemoState());
}

export function clearAll() {
    return saveState(emptyState());
}

/**
 * Entfernt nur Demo-Seed-Zeilen (Inventar + Vorgänge).
 */
export function clearDemoSeed() {
    const state = loadState();
    state.gegenstaende = (state.gegenstaende || []).filter((g) => !isDemoSeedItem(g));
    state.vorgaenge = (state.vorgaenge || []).filter((v) => !isDemoSeedItem(v));
    if (!(state.gegenstaende || []).length && !(state.vorgaenge || []).length) {
        state.localDemoOnly = false;
    }
    return saveState(state);
}

/**
 * @param {object} draft
 */
export function upsertTyp(draft) {
    const v = validateTyp(draft);
    if (!v.ok) return { ok: false, errors: v.errors, item: null, state: loadState() };
    const state = loadState();
    const item = normalizeTyp(draft);
    const idx = state.typen.findIndex((t) => t.id === item.id);
    if (idx >= 0) state.typen[idx] = { ...state.typen[idx], ...item };
    else state.typen.push(item);
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 */
export function removeTyp(id) {
    const state = loadState();
    const tid = String(id || '').trim();
    const used = state.gegenstaende.some((g) => g.typ === tid);
    if (used) return { ok: false, errors: ['Typ wird noch von Gegenständen genutzt.'], state };
    state.typen = state.typen.filter((t) => t.id !== tid);
    return { ok: true, errors: [], state: saveState(state) };
}

/**
 * Vorlage als Typ übernehmen (+ optional Beispiele).
 * @param {string} vorlageId
 * @param {{ withExamples?: boolean }} [opts]
 */
export function applyTypVorlage(vorlageId, opts = {}) {
    const mat = materializeTypVorlage(vorlageId, opts);
    if (!mat) return { ok: false, errors: ['Vorlage unbekannt.'], state: loadState() };
    const state = loadState();
    if (!state.typen.some((t) => t.id === mat.typ.id)) state.typen.push(normalizeTyp(mat.typ));
    for (const g of mat.gegenstaende) {
        if (state.gegenstaende.some((x) => x.kennung === g.kennung && x.typ === g.typ)) continue;
        state.gegenstaende.push(normalizeGegenstand(g));
    }
    return { ok: true, errors: [], state: saveState(state) };
}

/**
 * @param {object} draft
 */
export function upsertGegenstand(draft) {
    const v = validateGegenstand(draft);
    if (!v.ok) return { ok: false, errors: v.errors, item: null, state: loadState() };
    const state = loadState();
    const item = normalizeGegenstand(draft);
    const typMeta = state.typen.find((t) => t.id === item.typ);
    if (typMeta && !item.typLabel) item.typLabel = typMeta.label;
    const dup = state.gegenstaende.find(
        (g) => g.kennung.toLowerCase() === item.kennung.toLowerCase() && g.id !== item.id
    );
    if (dup) {
        return {
            ok: false,
            errors: ['Kennung bereits vergeben: ' + dup.kennung],
            item: null,
            state
        };
    }
    const idx = state.gegenstaende.findIndex((x) => x.id === item.id);
    if (idx >= 0) state.gegenstaende[idx] = { ...state.gegenstaende[idx], ...item };
    else state.gegenstaende.unshift(item);
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 */
export function removeGegenstand(id) {
    const state = loadState();
    const gid = String(id || '').trim();
    if (findOpenVorgang(gid, state.vorgaenge)) {
        return { ok: false, errors: ['Gegenstand ist noch ausgegeben.'], state };
    }
    state.gegenstaende = state.gegenstaende.filter((g) => g.id !== gid);
    return { ok: true, errors: [], state: saveState(state) };
}

/**
 * @param {object} draft
 */
export function ausgeben(draft) {
    const state = loadState();
    const v = normalizeVorgang({
        ...draft,
        status: VORGANG_STATUS.AUSGEGEBEN,
        zurueckAm: '',
        ausgegebenAm: draft.ausgegebenAm || toIsoDateOnly(new Date())
    });
    const g = state.gegenstaende.find((x) => x.id === v.gegenstandId);
    if (!g || !g.aktiv) {
        return { ok: false, errors: ['Gegenstand nicht gefunden oder inaktiv.'], state };
    }
    if (findOpenVorgang(g.id, state.vorgaenge)) {
        return { ok: false, errors: ['Gegenstand ist bereits ausgegeben.'], state };
    }
    v.kennung = v.kennung || g.kennung;
    const check = validateVorgang(v);
    if (!check.ok) return { ok: false, errors: check.errors, state };
    state.vorgaenge.unshift(v);
    return { ok: true, errors: [], item: v, state: saveState(state) };
}

/**
 * @param {string} vorgangId
 * @param {{ status?: string, zurueckAm?: string, notiz?: string }} [opts]
 */
export function zuruecknehmen(vorgangId, opts = {}) {
    const state = loadState();
    const item = state.vorgaenge.find((x) => x.id === String(vorgangId || '').trim());
    if (!item) return { ok: false, errors: ['Vorgang nicht gefunden.'], state };
    if (item.status !== VORGANG_STATUS.AUSGEGEBEN) {
        return { ok: false, errors: ['Vorgang ist nicht offen ausgegeben.'], state };
    }
    item.status = String(opts.status || VORGANG_STATUS.RETOUR).toLowerCase();
    item.zurueckAm = toIsoDateOnly(opts.zurueckAm) || toIsoDateOnly(new Date());
    if (opts.notiz) item.notiz = String(opts.notiz).trim();
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} vorgangId
 * @param {string} status
 */
export function setVorgangStatus(vorgangId, status) {
    const state = loadState();
    const item = state.vorgaenge.find((x) => x.id === String(vorgangId || '').trim());
    if (!item) return { ok: false, errors: ['Vorgang nicht gefunden.'], state };
    item.status = String(status || '').toLowerCase();
    if (item.status === VORGANG_STATUS.RETOUR && !item.zurueckAm) {
        item.zurueckAm = toIsoDateOnly(new Date());
    }
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} csvText
 * @param {{ defaultTyp?: string }} [opts]
 */
export function importGegenstaendeCsv(csvText, opts = {}) {
    const rows = parseGegenstaendeCsv(csvText);
    let added = 0;
    const errors = [];
    for (const row of rows) {
        const typ = row.typ || opts.defaultTyp || 'sonstiges';
        upsertTyp({ id: typ, label: typ });
        const res = upsertGegenstand({
            typ,
            kennung: row.kennung,
            bezeichnung: row.bezeichnung,
            ortZone: row.ortZone,
            aktiv: true,
            quelle: 'import'
        });
        if (res.ok) added += 1;
        else errors.push(...(res.errors || []));
    }
    return { ok: errors.length === 0, added, errors, state: loadState() };
}
