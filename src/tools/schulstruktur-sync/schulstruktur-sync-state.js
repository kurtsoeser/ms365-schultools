/**
 * Persistenz-Layer für „Schulstruktur-Sync".
 *
 * Aus `schulstruktur-sync.js` 1:1 ausgelagert (Phase 2-Pilot). Verhalten
 * identisch: localStorage als Fallback, `window.ms365AppDataV2` als
 * bevorzugter Container, falls verfügbar.
 *
 * Enthält ausschließlich Storage-Helfer – keine UI, keine Graph-API.
 * Damit ohne DOM testbar.
 */

import { safeJsonParse } from '../../shared/utils/json.js';

/** @internal Storage-Keys (intern, nicht exportiert – Zugriff nur über die Helfer). */
const STORAGE_KEY = 'ms365-schulstruktur-sync-v1';
const STORAGE_TENANT_CACHE_KEY = 'ms365-schulstruktur-tenant-cache-v1';
const STORAGE_MATCH_KEY = 'ms365-schulstruktur-match-v1';
/** Markierungen für AD-gesyncte Gruppen (Hinweise an lokalen Admin). */
const STORAGE_AD_FLAGS_KEY = 'ms365-schulstruktur-ad-flags-v1';
/** Geteilte ID-Menge mit der Graph-Ansicht (Tree-Collapse). */
const GRAPH_COLLAPSE_KEY = 'ms365-ss-graph-collapsed-v1';

/**
 * Liest den Strukturbaum, Mitgliedschaften und Settings. Bevorzugt
 * `ms365AppDataV2`, fällt sonst auf `localStorage` zurück.
 *
 * @returns {{ rows: any[], memberships: object, settings: object }}
 */
export function loadState() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
            const c = window.ms365AppDataV2.getContainer();
            if (c && c.structure && typeof c.structure === 'object') {
                const rows = Array.isArray(c.structure.rows) ? c.structure.rows : [];
                const memberships =
                    c.structure.memberships && typeof c.structure.memberships === 'object' ? c.structure.memberships : {};
                const settings =
                    c.structure.settings && typeof c.structure.settings === 'object' ? c.structure.settings : {};
                return { rows, memberships, settings };
            }
        }
        const raw = localStorage.getItem(STORAGE_KEY);
        const obj = raw ? safeJsonParse(raw) : null;
        const rows = obj && Array.isArray(obj.rows) ? obj.rows : [];
        const memberships =
            obj && obj.memberships && typeof obj.memberships === 'object' ? obj.memberships : {};
        const settings =
            obj && obj.settings && typeof obj.settings === 'object' ? obj.settings : {};
        return { rows, memberships, settings };
    } catch {
        return { rows: [], memberships: {}, settings: {} };
    }
}

/**
 * Persistiert den Strukturbaum. Schreibt parallel in `localStorage` UND
 * `ms365AppDataV2`, damit beide Quellen synchron bleiben.
 *
 * Vor dem Schreiben wird frisch geladen (Reload-before-save).
 * `organisationAssist` aus dem Speicher wird nicht von veraltetem
 * In-Memory-Settings der Gruppenverwaltung ueberschrieben, es sei denn
 * `organisationAssistSource: 'incoming'` (Org-Assistent).
 *
 * @param {{ rows?: any[], memberships?: object, settings?: object, organisationAssistSource?: 'incoming'|'fresh' }} state
 * @returns {{ rows: any[], memberships: object, settings: object }}
 */
export function saveState(state) {
    const fresh = loadState();
    const rows = state && Array.isArray(state.rows) ? state.rows : fresh.rows;
    const memberships =
        state && state.memberships && typeof state.memberships === 'object' ? state.memberships : fresh.memberships;
    const incoming = state && state.settings && typeof state.settings === 'object' ? state.settings : {};
    const settings = Object.assign({}, fresh.settings || {}, incoming);
    if (state && state.organisationAssistSource === 'incoming') {
        /* Org-Assistent: eingehendes organisationAssist behalten */
    } else if (fresh.settings && fresh.settings.organisationAssist !== undefined) {
        settings.organisationAssist = fresh.settings.organisationAssist;
    }
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ rows, memberships, settings }));
    } catch {
        // ignore
    }
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function' && typeof window.ms365AppDataV2.setContainer === 'function') {
            if (typeof window.ms365AppDataV2.invalidateCache === 'function') {
                window.ms365AppDataV2.invalidateCache();
            }
            const c = window.ms365AppDataV2.getContainer();
            c.structure = { rows, memberships, settings };
            window.ms365AppDataV2.setContainer(c);
        }
    } catch {
        // ignore
    }
    return { rows, memberships, settings };
}

/**
 * Multi-Tab: bei Aenderung der Struktur-Keys Handler aufrufen.
 * @param {(ev: StorageEvent) => void} handler
 * @returns {() => void} unsubscribe
 */
export function wireStructureStorageListener(handler) {
    if (typeof window === 'undefined' || typeof handler !== 'function') {
        return function () {};
    }
    const keys = new Set([STORAGE_KEY, STORAGE_MATCH_KEY, 'ms365-schooltool-data-v2']);
    const fn = function (e) {
        if (!e || !e.key || !keys.has(e.key)) return;
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.invalidateCache === 'function') {
            window.ms365AppDataV2.invalidateCache();
        }
        handler(e);
    };
    window.addEventListener('storage', fn);
    return function () {
        window.removeEventListener('storage', fn);
    };
}

/**
 * Liest die Match-Links (Struktur-ID → Tenant-Group-ID / -User-ID + Notiz).
 *
 * @returns {{ links: object }}
 */
export function loadMatchState() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
            const c = window.ms365AppDataV2.getContainer();
            if (c && c.match && c.match.links && typeof c.match.links === 'object') {
                return { links: c.match.links };
            }
        }
        const raw = localStorage.getItem(STORAGE_MATCH_KEY);
        const obj = raw ? safeJsonParse(raw) : null;
        const links = obj && obj.links && typeof obj.links === 'object' ? obj.links : {};
        return { links };
    } catch {
        return { links: {} };
    }
}

/**
 * Persistiert die Match-Links.
 * @param {object} links
 * @returns {object} die persistierten Links (immer ein Objekt)
 */
export function saveMatchState(links) {
    const out = links && typeof links === 'object' ? links : {};
    try {
        localStorage.setItem(STORAGE_MATCH_KEY, JSON.stringify({ links: out }));
    } catch {
        // ignore
    }
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function' && typeof window.ms365AppDataV2.setContainer === 'function') {
            const c = window.ms365AppDataV2.getContainer();
            c.match = { links: out };
            window.ms365AppDataV2.setContainer(c);
        }
    } catch {
        // ignore
    }
    return out;
}

/**
 * Liest den zwischengespeicherten Tenant-Inventar-Snapshot (Gruppen + User
 * + Ladestempel).
 *
 * @returns {{ rows: any[], users: any[], loadedAt: string }}
 */
export function loadTenantCache() {
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
            const c = window.ms365AppDataV2.getContainer();
            const cache = c && c.tenant && c.tenant.cache && typeof c.tenant.cache === 'object' ? c.tenant.cache : null;
            if (cache) {
                const rows = Array.isArray(cache.rows) ? cache.rows : [];
                const users = Array.isArray(cache.users) ? cache.users : [];
                return { rows, users, loadedAt: cache.loadedAt ? String(cache.loadedAt) : '' };
            }
        }
        const raw = localStorage.getItem(STORAGE_TENANT_CACHE_KEY);
        const obj = raw ? safeJsonParse(raw) : null;
        const rows = obj && Array.isArray(obj.rows) ? obj.rows : [];
        const users = obj && Array.isArray(obj.users) ? obj.users : [];
        return { rows, users, loadedAt: obj && obj.loadedAt ? String(obj.loadedAt) : '' };
    } catch {
        return { rows: [], users: [], loadedAt: '' };
    }
}

/**
 * Schreibt den Tenant-Inventar-Snapshot. Wenn `users` nicht angegeben ist,
 * wird der bestehende User-Cache beibehalten.
 *
 * @param {any[]} rows
 * @param {any[]} [users]
 */
export function saveTenantCache(rows, users) {
    const out = Array.isArray(rows) ? rows : [];
    const prev = loadTenantCache();
    const u = users !== undefined ? (Array.isArray(users) ? users : []) : prev.users || [];
    try {
        localStorage.setItem(
            STORAGE_TENANT_CACHE_KEY,
            JSON.stringify({ rows: out, users: u, loadedAt: new Date().toISOString() })
        );
    } catch {
        // ignore
    }
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function' && typeof window.ms365AppDataV2.setContainer === 'function') {
            const c = window.ms365AppDataV2.getContainer();
            c.tenant = { cache: { rows: out, users: u, loadedAt: new Date().toISOString() } };
            window.ms365AppDataV2.setContainer(c);
        }
    } catch {
        // ignore
    }
}

/**
 * Liest die Liste der eingeklappten Knoten-IDs aus der Graph-Ansicht
 * als `Set<string>`. Schluckt JSON-Fehler und liefert dann ein leeres Set.
 *
 * @returns {Set<string>}
 */
export function loadGraphCollapsedSet() {
    try {
        const raw = localStorage.getItem(GRAPH_COLLAPSE_KEY);
        if (!raw) return new Set();
        const arr = JSON.parse(raw);
        if (!Array.isArray(arr)) return new Set();
        return new Set(arr.map((x) => String(x)));
    } catch {
        return new Set();
    }
}

/**
 * Persistiert die Collapsed-IDs.
 * @param {Set<string>} set
 */
export function saveGraphCollapsedSet(set) {
    try {
        localStorage.setItem(GRAPH_COLLAPSE_KEY, JSON.stringify(Array.from(set || []).map((x) => String(x))));
    } catch {
        // ignore
    }
}

/**
 * Lokale Markierungen für AD-gesyncte / hybrid Gruppen.
 * @returns {Record<string, { flagged: boolean, note: string, flaggedAt: string }>}
 */
export function loadAdGroupFlags() {
    try {
        const raw = localStorage.getItem(STORAGE_AD_FLAGS_KEY);
        const obj = raw ? safeJsonParse(raw) : null;
        if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return {};
        /** @type {Record<string, { flagged: boolean, note: string, flaggedAt: string }>} */
        const out = {};
        Object.keys(obj).forEach((k) => {
            const id = String(k || '').trim();
            if (!id) return;
            const v = obj[k];
            if (!v || typeof v !== 'object') return;
            out[id] = {
                flagged: !!v.flagged,
                note: v.note != null ? String(v.note) : '',
                flaggedAt: v.flaggedAt != null ? String(v.flaggedAt) : ''
            };
        });
        return out;
    } catch {
        return {};
    }
}

/**
 * @param {Record<string, { flagged?: boolean, note?: string, flaggedAt?: string }>} map
 */
export function saveAdGroupFlags(map) {
    const src = map && typeof map === 'object' ? map : {};
    /** @type {Record<string, { flagged: boolean, note: string, flaggedAt: string }>} */
    const clean = {};
    Object.keys(src).forEach((k) => {
        const id = String(k || '').trim();
        if (!id) return;
        const v = src[k];
        if (!v || typeof v !== 'object') return;
        const flagged = !!v.flagged;
        const note = v.note != null ? String(v.note) : '';
        if (!flagged && !note.trim()) return;
        clean[id] = {
            flagged: flagged,
            note: note,
            flaggedAt: v.flaggedAt != null ? String(v.flaggedAt) : flagged ? new Date().toISOString() : ''
        };
    });
    try {
        localStorage.setItem(STORAGE_AD_FLAGS_KEY, JSON.stringify(clean));
    } catch {
        // ignore
    }
}

/**
 * @param {string} groupId
 * @param {{ flagged?: boolean, note?: string }} patch
 */
export function patchAdGroupFlag(groupId, patch) {
    const id = String(groupId || '').trim();
    if (!id) return loadAdGroupFlags();
    const map = loadAdGroupFlags();
    const prev = map[id] || { flagged: false, note: '', flaggedAt: '' };
    const nextFlagged = patch && patch.flagged !== undefined ? !!patch.flagged : !!prev.flagged;
    const nextNote = patch && patch.note !== undefined ? String(patch.note) : String(prev.note || '');
    if (!nextFlagged && !String(nextNote || '').trim()) {
        delete map[id];
    } else {
        map[id] = {
            flagged: nextFlagged,
            note: nextNote,
            flaggedAt:
                nextFlagged && !prev.flagged
                    ? new Date().toISOString()
                    : prev.flaggedAt || (nextFlagged ? new Date().toISOString() : '')
        };
    }
    saveAdGroupFlags(map);
    return map;
}
