/**
 * Schul-Portal-Standard (IT): angeheftete Home-Apps und optionale Sektions-Reihenfolge.
 * Persönliche Favoriten bleiben getrennt (ms365-dashboard-favorites-v1).
 */
import { notifyAppLocalDataChanged } from './app-local-data-notify.js';
import {
    DASHBOARD_LEHRER_SCHUELER_APPS,
    DASHBOARD_PORTAL_SECTIONS_BY_VIEW
} from './dashboard-audience-catalog.js';

export const DASHBOARD_PORTAL_DEFAULTS_STORAGE_KEY = 'ms365-dashboard-portal-defaults-v1';

/**
 * @typedef {{
 *   homePinned: string[],
 *   sectionOrder: string[]
 * }} PortalPersonaDefaults
 *
 * @typedef {{
 *   version: number,
 *   lehrer: PortalPersonaDefaults,
 *   schueler: PortalPersonaDefaults
 * }} DashboardPortalDefaultsConfig
 */

/** @returns {PortalPersonaDefaults} */
function emptyPersona() {
    return { homePinned: [], sectionOrder: [] };
}

/**
 * @returns {DashboardPortalDefaultsConfig}
 */
export function emptyPortalDefaultsConfig() {
    return {
        version: 1,
        lehrer: emptyPersona(),
        schueler: emptyPersona()
    };
}

function safeParse(raw) {
    try {
        return JSON.parse(String(raw));
    } catch {
        return null;
    }
}

/**
 * @param {unknown} list
 * @returns {string[]}
 */
export function normalizeToolIdList(list) {
    if (!Array.isArray(list)) return [];
    const out = [];
    const seen = new Set();
    for (const raw of list) {
        const id = String(raw || '').trim();
        if (!id || seen.has(id)) continue;
        if (!DASHBOARD_LEHRER_SCHUELER_APPS.includes(id)) continue;
        seen.add(id);
        out.push(id);
    }
    return out;
}

/**
 * @param {'lehrer'|'schueler'} view
 * @param {unknown} list
 * @returns {string[]}
 */
export function normalizeSectionOrder(view, list) {
    const known = (DASHBOARD_PORTAL_SECTIONS_BY_VIEW[view] || []).map((s) => s.id);
    const knownSet = new Set(known);
    if (!Array.isArray(list)) return [];
    const out = [];
    const seen = new Set();
    for (const raw of list) {
        const id = String(raw || '').trim();
        if (!id || seen.has(id) || !knownSet.has(id)) continue;
        seen.add(id);
        out.push(id);
    }
    return out;
}

/**
 * @param {unknown} data
 * @returns {DashboardPortalDefaultsConfig}
 */
export function normalizePortalDefaultsConfig(data) {
    const base = emptyPortalDefaultsConfig();
    if (!data || typeof data !== 'object') return base;
    const src = /** @type {Record<string, unknown>} */ (data);
    for (const view of /** @type {const} */ (['lehrer', 'schueler'])) {
        const block = src[view];
        if (!block || typeof block !== 'object') continue;
        const b = /** @type {Record<string, unknown>} */ (block);
        base[view] = {
            homePinned: normalizeToolIdList(b.homePinned),
            sectionOrder: normalizeSectionOrder(view, b.sectionOrder)
        };
    }
    return base;
}

/**
 * @param {Storage} [storage]
 * @returns {DashboardPortalDefaultsConfig}
 */
export function loadPortalDefaultsConfig(storage = localStorage) {
    try {
        const raw = storage.getItem(DASHBOARD_PORTAL_DEFAULTS_STORAGE_KEY);
        const data = safeParse(raw);
        return normalizePortalDefaultsConfig(data);
    } catch {
        return emptyPortalDefaultsConfig();
    }
}

/**
 * @param {DashboardPortalDefaultsConfig} config
 * @param {Storage} [storage]
 */
export function savePortalDefaultsConfig(config, storage = localStorage) {
    const payload = normalizePortalDefaultsConfig(config);
    try {
        storage.setItem(DASHBOARD_PORTAL_DEFAULTS_STORAGE_KEY, JSON.stringify(payload));
    } catch {
        /* ignore */
    }
    try {
        window.dispatchEvent(new CustomEvent('ms365-dashboard-portal-defaults-changed'));
    } catch {
        /* ignore */
    }
    notifyAppLocalDataChanged('dashboard-portal-defaults');
    return payload;
}

export function resetPortalDefaultsConfig(storage = localStorage) {
    return savePortalDefaultsConfig(emptyPortalDefaultsConfig(), storage);
}

/**
 * @param {'lehrer'|'schueler'} view
 * @param {string[]} homePinned
 * @param {DashboardPortalDefaultsConfig} [base]
 * @param {Storage} [storage]
 */
export function setPortalHomePinned(view, homePinned, base, storage = localStorage) {
    const cfg = base || loadPortalDefaultsConfig(storage);
    const next = normalizePortalDefaultsConfig(cfg);
    if (view !== 'lehrer' && view !== 'schueler') return next;
    next[view] = {
        ...next[view],
        homePinned: normalizeToolIdList(homePinned)
    };
    return savePortalDefaultsConfig(next, storage);
}

/**
 * @param {'lehrer'|'schueler'} view
 * @param {string[]} sectionOrder
 * @param {DashboardPortalDefaultsConfig} [base]
 * @param {Storage} [storage]
 */
export function setPortalSectionOrder(view, sectionOrder, base, storage = localStorage) {
    const cfg = base || loadPortalDefaultsConfig(storage);
    const next = normalizePortalDefaultsConfig(cfg);
    if (view !== 'lehrer' && view !== 'schueler') return next;
    next[view] = {
        ...next[view],
        sectionOrder: normalizeSectionOrder(view, sectionOrder)
    };
    return savePortalDefaultsConfig(next, storage);
}

/**
 * Home-Apps: Schul-Pins zuerst, dann persönliche Favoriten, dann Fallback-Priorität.
 * Nur IDs, die in visibleIds liegen.
 *
 * @param {{
 *   visibleIds: string[],
 *   schoolPinned?: string[],
 *   personalFavorites?: string[],
 *   fallbackPriority?: string[],
 *   max?: number
 * }} opts
 * @returns {string[]}
 */
export function resolvePortalHomeToolIds(opts) {
    const visible = new Set((opts.visibleIds || []).map(String));
    const max = Math.max(1, Number(opts.max) || 6);
    /** @type {string[]} */
    const out = [];
    const push = (id) => {
        const t = String(id || '').trim();
        if (!t || out.includes(t) || !visible.has(t)) return;
        if (out.length < max) out.push(t);
    };
    (opts.schoolPinned || []).forEach(push);
    (opts.personalFavorites || []).forEach(push);
    (opts.fallbackPriority || []).forEach(push);
    (opts.visibleIds || []).forEach(push);
    return out;
}

/**
 * @template {{ id: string }} T
 * @param {T[]} sections
 * @param {string[]} [sectionOrder]
 * @returns {T[]}
 */
export function applyPortalSectionOrder(sections, sectionOrder) {
    const list = Array.isArray(sections) ? sections.slice() : [];
    const order = Array.isArray(sectionOrder) ? sectionOrder.filter(Boolean) : [];
    if (!order.length || !list.length) return list;
    const byId = new Map(list.map((s) => [s.id, s]));
    /** @type {T[]} */
    const ordered = [];
    const seen = new Set();
    for (const id of order) {
        const sec = byId.get(id);
        if (sec && !seen.has(id)) {
            ordered.push(sec);
            seen.add(id);
        }
    }
    for (const sec of list) {
        if (!seen.has(sec.id)) ordered.push(sec);
    }
    return ordered;
}
