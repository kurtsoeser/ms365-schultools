/**
 * localStorage-Store Krankmeldung.
 */

import { deleteAttachment } from './krankmeldung-attachments.js';
import { buildKrankmeldungDemoState } from './krankmeldung-demo.js';
import { findDuplicate, validateMeldung } from './krankmeldung-logic.js';
import {
    STATUS,
    emptyState,
    isDemoSeedItem,
    normalizeMeldung,
    normalizeState,
    STORAGE_KEY
} from './krankmeldung-model.js';
import { loadKrankmeldungSettings } from './krankmeldung-settings.js';

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

/**
 * Demo nur aufgerufen wenn UI ?demo=1 oder expliziten Button nutzt.
 * Nie still beim normalen App-Start.
 */
export function seedDemoIfEmpty() {
    const state = loadState();
    if ((state.meldungen || []).length) return state;
    return saveState(buildKrankmeldungDemoState());
}

/** Demo mergen – echte (nicht-Demo) Meldungen bleiben. */
export function loadDemoData() {
    const demo = buildKrankmeldungDemoState();
    const current = loadState();
    const kept = (current.meldungen || []).filter((m) => !isDemoRow(m));
    return saveState({
        meldungen: [...(demo.meldungen || []), ...kept]
    });
}

/** Nur Demo-Seed-Zeilen entfernen. */
export function clearDemoSeeds() {
    const state = loadState();
    const meldungen = (state.meldungen || []).filter((m) => !isDemoRow(m));
    return saveState({
        meldungen,
        localDemoOnly: meldungen.length ? state.localDemoOnly : false
    });
}

export function clearAll() {
    return saveState(emptyState());
}

/**
 * Sync-Merge ohne Validierung (nach SPO Create/Update spoItemId merken).
 * @param {object} draft
 */
export function mergeRemoteMeldung(draft) {
    const state = loadState();
    const item = normalizeMeldung(draft);
    const idx = state.meldungen.findIndex((x) => x.id === item.id);
    if (idx >= 0) {
        state.meldungen[idx] = {
            ...state.meldungen[idx],
            ...item,
            spoItemId: item.spoItemId || state.meldungen[idx].spoItemId || ''
        };
        return { ok: true, item: state.meldungen[idx], state: saveState(state) };
    }
    state.meldungen.unshift(item);
    return { ok: true, item, state: saveState(state) };
}

/**
 * @param {object} draft
 * @param {{ allowDuplicate?: boolean, requireAnhang?: boolean }} [opts]
 */
export function upsertMeldung(draft, opts = {}) {
    const settings = loadKrankmeldungSettings();
    const v = validateMeldung(draft, {
        requireAnhang: !!(opts.requireAnhang || (settings.anhangPflicht && settings.allowAnhang))
    });
    if (!v.ok) return { ok: false, errors: v.errors, warnings: v.warnings || [], item: null, state: loadState(), duplicate: null };
    const state = loadState();
    let item = normalizeMeldung(draft);
    const dup = opts.allowDuplicate
        ? null
        : findDuplicate(item, state.meldungen, item.id);
    if (dup) {
        return {
            ok: false,
            errors: [
                'Ähnliche Meldung existiert bereits (' +
                    dup.schuelerName +
                    ', ' +
                    dup.von +
                    '–' +
                    dup.bis +
                    ').'
            ],
            item: null,
            state,
            duplicate: dup
        };
    }
    const idx = state.meldungen.findIndex((x) => x.id === item.id);
    if (idx >= 0) {
        const prev = state.meldungen[idx];
        state.meldungen[idx] = {
            ...prev,
            ...item,
            spoItemId: item.spoItemId || prev.spoItemId || '',
            seedTag: prev.seedTag || item.seedTag || '',
            gemeldetAm: prev.gemeldetAm || item.gemeldetAm
        };
        item = state.meldungen[idx];
    } else {
        state.meldungen.unshift(item);
    }
    return {
        ok: true,
        errors: [],
        warnings: v.warnings || [],
        item,
        state: saveState(state),
        duplicate: null
    };
}

/**
 * Öffentliche Eingabe – Duplikat nur Warnung, speichern trotzdem mit allow.
 * @param {object} draft
 */
export function submitPublicMeldung(draft) {
    const state = loadState();
    const settings = loadKrankmeldungSettings();
    const item = normalizeMeldung({
        ...draft,
        status: STATUS.NEU,
        quelle: 'public',
        gemeldetAm: new Date().toISOString()
    });
    const v = validateMeldung(item, {
        requireAnhang: !!(settings.allowAnhang && settings.anhangPflicht)
    });
    if (!v.ok) {
        return {
            ok: false,
            errors: v.errors,
            warnings: v.warnings || [],
            item: null,
            duplicate: null,
            state
        };
    }
    const dup = findDuplicate(item, state.meldungen);
    state.meldungen.unshift(item);
    return {
        ok: true,
        errors: [],
        warnings: v.warnings || [],
        item,
        duplicate: dup || null,
        state: saveState(state)
    };
}

/**
 * @param {string} id
 * @param {string} status
 */
export function setMeldungStatus(id, status) {
    const state = loadState();
    const item = state.meldungen.find((x) => x.id === String(id || '').trim());
    if (!item) return { ok: false, errors: ['Meldung nicht gefunden.'], state };
    item.status = String(status || '').trim().toLowerCase() || item.status;
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 * @param {string} email
 */
export function markKenntnis(id, email) {
    const state = loadState();
    const item = state.meldungen.find((x) => x.id === String(id || '').trim());
    if (!item) return { ok: false, errors: ['Meldung nicht gefunden.'], state };
    item.kenntnisLehrerEmail = String(email || '').trim().toLowerCase();
    if (item.status === STATUS.NEU) item.status = STATUS.ZUR_KENNTNIS;
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {string} id
 */
export function removeMeldung(id) {
    const state = loadState();
    const want = String(id || '').trim();
    const victim = state.meldungen.find((m) => m.id === want);
    if (victim && victim.anhangKey) {
        void deleteAttachment(victim.anhangKey);
    }
    state.meldungen = state.meldungen.filter((m) => m.id !== want);
    return { ok: true, state: saveState(state) };
}

export { STATUS };
