/**
 * String/clone/persist helpers for Schul-Organigramm edit.
 */
import { normStr, normEmail, normCode, escapeHtml } from '../../shared/utils/strings.js';
import { normalizeSchulOrganigramm } from './schul-organigramm-logic.js';

export { normStr, normEmail, normCode, escapeHtml };

/**
 * Eindeutiges Lehrerkürzel vorschlagen (für Stammdaten-Persistenz nötig).
 * @param {object[]} teachers
 * @param {string} name
 * @param {string} email
 */
export function suggestTeacherCode(teachers, name, email) {
    let base = '';
    const nm = normStr(name);
    if (nm) {
        const parts = nm.split(/\s+/).filter(Boolean);
        if (parts.length >= 2) {
            base = normCode(parts[parts.length - 1].slice(0, 3) + parts[0].slice(0, 2));
        } else {
            base = normCode(nm.slice(0, 4));
        }
    }
    if (!base && email) {
        base = normCode(String(email).split('@')[0]).slice(0, 6);
    }
    if (!base) base = 'LEH';
    const used = new Set(
        (Array.isArray(teachers) ? teachers : [])
            .map(function (t) {
                return normCode(t && t.code).toLowerCase();
            })
            .filter(Boolean)
    );
    if (!used.has(base.toLowerCase())) return base;
    let i = 2;
    while (used.has((base + i).toLowerCase()) && i < 999) i += 1;
    return base + i;
}

/**
 * Lehrkraft in settings.teachers[] anlegen/aktualisieren (mit Kürzel).
 * @param {object} settings
 * @param {{ name?: string, email?: string }} person
 */
export function ensureTeacherInList(settings, person) {
    const email = normEmail(person && person.email);
    const name = normStr(person && person.name);
    if (!email && !name) return null;
    if (!Array.isArray(settings.teachers)) settings.teachers = [];
    const tIdx = settings.teachers.findIndex(function (t) {
        return email && normEmail(t && t.email) === email;
    });
    if (tIdx >= 0) {
        const row = settings.teachers[tIdx];
        if (name && !normStr(row.name)) row.name = name;
        if (email && !normEmail(row.email)) row.email = email;
        if (!normCode(row.code)) {
            row.code = suggestTeacherCode(settings.teachers, row.name || name, row.email || email);
        }
        return row;
    }
    const code = suggestTeacherCode(settings.teachers, name, email);
    const row = { code: code, name: name || email, email: email };
    settings.teachers.push(row);
    return row;
}

/** Ephemere Anreicherungen (years.*) nicht in Tenant-Settings schreiben. */
export function stripEphemeralSchoolSettings(settings) {
    if (!settings || typeof settings !== 'object') return settings;
    delete settings.classChats;
    delete settings.unterrichtsbelegung;
    return settings;
}

/**
 * Tiefenbegrenzter Clone (Zyklen + extreme Nesting) – wirft nicht „Maximum call stack“.
 * @param {unknown} value
 * @param {number} [maxDepth]
 */
export function safeStructuredClone(value, maxDepth) {
    const limit = typeof maxDepth === 'number' && maxDepth > 0 ? maxDepth : 48;
    /** @type {WeakMap<object, unknown>} */
    const seen = new WeakMap();
    /**
     * @param {unknown} v
     * @param {number} depth
     */
    function walk(v, depth) {
        if (v == null) return v;
        const t = typeof v;
        if (t === 'string' || t === 'number' || t === 'boolean') return v;
        if (t !== 'object') return undefined;
        if (depth <= 0) return Array.isArray(v) ? [] : {};
        const obj = /** @type {object} */ (v);
        if (seen.has(obj)) return seen.get(obj);
        if (Array.isArray(v)) {
            /** @type {unknown[]} */
            const arr = [];
            seen.set(obj, arr);
            const n = Math.min(v.length, 100000);
            for (let i = 0; i < n; i += 1) {
                const item = v[i];
                const it = typeof item;
                if (item == null || it === 'string' || it === 'number' || it === 'boolean') {
                    arr.push(item);
                } else if (it === 'object') {
                    arr.push(walk(item, depth - 1));
                }
            }
            return arr;
        }
        /** @type {Record<string, unknown>} */
        const out = {};
        seen.set(obj, out);
        Object.keys(obj).forEach(function (key) {
            if (!key || key === '__proto__' || key === 'constructor') return;
            const item = /** @type {Record<string, unknown>} */ (obj)[key];
            const it = typeof item;
            if (item == null || it === 'string' || it === 'number' || it === 'boolean') {
                out[key] = item;
            } else if (it === 'object') {
                out[key] = walk(item, depth - 1);
            }
        });
        return out;
    }
    return walk(value, limit);
}

/**
 * Layout/Settings als reines JSON materialisieren (bricht Zyklen/Proxies/tiefe Nesting).
 * @param {unknown} value
 * @param {unknown} [fallback]
 */
export function plainJsonClone(value, fallback) {
    try {
        return JSON.parse(JSON.stringify(value));
    } catch {
        try {
            return safeStructuredClone(value);
        } catch {
            return fallback !== undefined ? fallback : value;
        }
    }
}

/**
 * Settings vor Persistenz bereinigen (keine classChats/Belegung, JSON-sicher).
 * @param {object} settings
 */
export function settingsForPersist(settings) {
    const src = stripEphemeralSchoolSettings(
        settings && typeof settings === 'object' ? Object.assign({}, settings) : {}
    );
    const clean = plainJsonClone(src, null);
    if (!clean || typeof clean !== 'object' || Array.isArray(clean)) {
        throw new Error('Einstellungen nicht speicherbar (Datenstruktur prüfen).');
    }
    if (clean.schulOrganigramm) {
        clean.schulOrganigramm = normalizeSchulOrganigramm(clean.schulOrganigramm);
    }
    stripEphemeralSchoolSettings(clean);
    return clean;
}

export function isBlankPerson(p) {
    return !normStr(p && p.name) && !normEmail(p && p.email);
}

export function personMatchesRole(p, role) {
    const name = normStr(role && role.name).toLowerCase();
    const code = normStr(role && role.code).toLowerCase();
    const r = normStr(p && p.role).toLowerCase();
    const dk = normStr(p && p.defaultKey).toLowerCase();
    return (name && (r === name || dk === name)) || (code && (r === code || dk === code));
}

export function loadSettings() {
    let settings = {};
    if (typeof window.ms365TenantSettingsLoad === 'function') {
        settings = window.ms365TenantSettingsLoad() || {};
    }
    // Belegung + KlassenChats liegen in years.* (App-Data v2) – nötig für Lehrkräfte in KlassenChats
    try {
        const api = window.ms365AppDataV2;
        if (!settings.unterrichtsbelegung && api && typeof api.getUnterrichtsbelegung === 'function') {
            const snap = api.getUnterrichtsbelegung();
            if (snap && Array.isArray(snap.rows) && snap.rows.length) {
                settings = Object.assign({}, settings, { unterrichtsbelegung: snap });
            }
        }
        if (api && typeof api.getClassChats === 'function') {
            const chats = api.getClassChats();
            if (chats && typeof chats === 'object') {
                settings = Object.assign({}, settings, { classChats: chats });
            }
        }
    } catch {
        /* ignore */
    }
    return settings;
}

export function saveSettings(settings) {
    const clean = settingsForPersist(settings);
    if (typeof window.ms365TenantSettingsSave === 'function') {
        return window.ms365TenantSettingsSave(clean);
    }
    return clean;
}

export function toast(msg, kind) {
    if (typeof window.ms365ShowToast === 'function') {
        window.ms365ShowToast(msg, { kind: kind || 'info', title: 'Schul-Organigramm' });
        return;
    }
    if (typeof window.ms365ToastOrAlert === 'function') {
        window.ms365ToastOrAlert(msg, { kind: kind || 'info', title: 'Schul-Organigramm' });
        return;
    }
    if (typeof window.ms365AppDialogAlert === 'function') {
        window.ms365AppDialogAlert(msg, { kind: kind || 'info', title: 'Schul-Organigramm' });
        return;
    }
    if (typeof window.ms365Toast === 'function') window.ms365Toast(msg);
}

export function confirmDlg(msg, opts) {
    if (typeof window.ms365AppDialogConfirm === 'function') {
        return window.ms365AppDialogConfirm(msg, opts || {});
    }
    return Promise.resolve(window.confirm(msg));
}

/**
 * Lokale Platzhalter-Mail für Personen ohne Entra-Konto.
 * @param {string} name
 */
export function suggestLocalPersonEmail(name) {
    const slug =
        normStr(name)
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/ä/g, 'ae')
            .replace(/ö/g, 'oe')
            .replace(/ü/g, 'ue')
            .replace(/ß/g, 'ss')
            .replace(/[^a-z0-9]+/g, '.')
            .replace(/^\.+|\.+$/g, '')
            .slice(0, 40) || 'person';
    return slug + '.' + Date.now().toString(36) + '@local.schul-org';
}
