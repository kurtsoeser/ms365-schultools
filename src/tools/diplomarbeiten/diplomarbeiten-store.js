/**
 * localStorage: Maturajahrgänge + Arbeiten (Kanal-Modus).
 */
import { labelsForArbeitTyp, validateArbeit } from './diplomarbeiten-logic.js';
import { loadDiplomSettings } from './diplomarbeiten-settings.js';

export const STORAGE_KEY = 'ms365-diplomarbeiten-state-v1';

export const ARBEIT_STATUS = {
    GEPLANT: 'geplant',
    AKTIV: 'aktiv',
    ABGESCHLOSSEN: 'abgeschlossen',
    ARCHIVIERT: 'archiviert'
};

export const ARBEIT_STATUS_LABEL = {
    geplant: 'Geplant',
    aktiv: 'Aktiv',
    abgeschlossen: 'Abgeschlossen',
    archiviert: 'Archiviert'
};

function newId(prefix) {
    const p = String(prefix || 'da').trim() || 'da';
    try {
        if (typeof crypto !== 'undefined' && crypto.randomUUID) {
            return p + '-' + crypto.randomUUID().replace(/-/g, '').slice(0, 12);
        }
    } catch {
        /* ignore */
    }
    return p + '-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function storage() {
    if (typeof localStorage === 'undefined') return null;
    return localStorage;
}

/**
 * @returns {{ cohorts: object[], arbeiten: object[], updatedAt: string }}
 */
export function emptyState() {
    return { cohorts: [], arbeiten: [], updatedAt: new Date().toISOString() };
}

function normalizeCohort(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    return {
        id: String(r.id || newId('dj')),
        maturajahrgang: String(r.maturajahrgang || '').trim(),
        displayName: String(r.displayName || '').trim(),
        mailNickname: String(r.mailNickname || '').trim().toLowerCase(),
        teamId: String(r.teamId || '').trim(),
        teamWebUrl: String(r.teamWebUrl || '').trim(),
        notes: String(r.notes || '').trim(),
        createdAt: String(r.createdAt || new Date().toISOString())
    };
}

function normalizeArbeit(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    const status = String(r.status || ARBEIT_STATUS.GEPLANT).toLowerCase();
    const mails = Array.isArray(r.studentEmails)
        ? r.studentEmails
        : String(r.studentEmails || '')
              .split(/[\n;,]+/)
              .map((s) => s.trim().toLowerCase())
              .filter((s) => s.includes('@'));
    const names = Array.isArray(r.studentNames)
        ? r.studentNames.map((s) => String(s || '').trim()).filter(Boolean)
        : [];
    return {
        id: String(r.id || newId('da')),
        cohortId: String(r.cohortId || '').trim(),
        topic: String(r.topic || '').trim(),
        student: String(r.student || '').trim(),
        mentorName: String(r.mentorName || '').trim(),
        mentorEmail: String(r.mentorEmail || '')
            .trim()
            .toLowerCase(),
        studentEmails: mails,
        studentNames: names.length ? names : mails.map(() => ''),
        status: Object.prototype.hasOwnProperty.call(ARBEIT_STATUS_LABEL, status)
            ? status
            : ARBEIT_STATUS.GEPLANT,
        channelId: String(r.channelId || '').trim(),
        channelName: String(r.channelName || '').trim(),
        channelWebUrl: String(r.channelWebUrl || '').trim(),
        createdAt: String(r.createdAt || new Date().toISOString()),
        updatedAt: String(r.updatedAt || new Date().toISOString())
    };
}

export function normalizeState(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    return {
        cohorts: Array.isArray(r.cohorts) ? r.cohorts.map(normalizeCohort) : [],
        arbeiten: Array.isArray(r.arbeiten) ? r.arbeiten.map(normalizeArbeit) : [],
        updatedAt: String(r.updatedAt || new Date().toISOString())
    };
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

export function clearState() {
    return saveState(emptyState());
}

/**
 * @param {object} draft
 */
export function upsertCohort(draft) {
    const state = loadState();
    const item = normalizeCohort(draft);
    if (!item.maturajahrgang) {
        return { ok: false, errors: ['Maturajahrgang fehlt (z. B. 2026/27).'], item: null, state };
    }
    const idx = state.cohorts.findIndex((c) => c.id === item.id);
    if (idx >= 0) state.cohorts[idx] = { ...state.cohorts[idx], ...item };
    else state.cohorts.push(item);
    return { ok: true, errors: [], item, state: saveState(state) };
}

/**
 * @param {object} draft
 */
export function upsertArbeit(draft) {
    const settings = loadDiplomSettings();
    const v = validateArbeit({ ...draft, arbeitTyp: settings.arbeitTyp });
    if (!v.ok) return { ok: false, errors: v.errors, item: null, state: loadState() };
    if (!String(draft && draft.cohortId || '').trim()) {
        return { ok: false, errors: ['Bitte einen Maturajahrgang wählen.'], item: null, state: loadState() };
    }
    const state = loadState();
    const item = normalizeArbeit({
        ...draft,
        channelName: (draft && draft.channelName) || v.channelName,
        studentEmails: v.studentEmails,
        updatedAt: new Date().toISOString()
    });
    const idx = state.arbeiten.findIndex((a) => a.id === item.id);
    if (idx >= 0) state.arbeiten[idx] = { ...state.arbeiten[idx], ...item };
    else state.arbeiten.push(item);
    return { ok: true, errors: [], item, state: saveState(state), labels: labelsForArbeitTyp(settings.arbeitTyp) };
}

/**
 * @param {string} id
 * @param {object} patch
 */
export function patchArbeit(id, patch) {
    const state = loadState();
    const idx = state.arbeiten.findIndex((a) => a.id === id);
    if (idx < 0) return { ok: false, errors: ['Arbeit nicht gefunden.'], state };
    state.arbeiten[idx] = normalizeArbeit({
        ...state.arbeiten[idx],
        ...(patch || {}),
        updatedAt: new Date().toISOString()
    });
    return { ok: true, errors: [], item: state.arbeiten[idx], state: saveState(state) };
}

export function getCohort(id) {
    return loadState().cohorts.find((c) => c.id === id) || null;
}

export function getArbeit(id) {
    return loadState().arbeiten.find((a) => a.id === id) || null;
}
