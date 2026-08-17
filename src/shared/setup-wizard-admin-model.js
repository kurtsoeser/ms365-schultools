/**
 * Verwaltungs-/Admin-Datenmodell für den Einrichtungs-Assistenten.
 * Aus setup-wizard.js ausgelagert (gleiches Verhalten).
 *
 * Die String-Helper (`normStr`, `normEmail`, `normCode`, `escapeHtml`) werden
 * aus `./utils/strings.js` re-exportiert, damit bestehende Imports
 * `import { normStr } from '.../setup-wizard-admin-model.js'` weiterhin
 * funktionieren. Neue Aufrufer sollten direkt aus `utils/strings.js` importieren.
 */

export { normStr, normEmail, normCode, escapeHtml } from './utils/strings.js';

import { normStr, normEmail } from './utils/strings.js';

export const SW_ADMIN_DEFAULT_ROLES = [
    'Direktion',
    'Sekretariat',
    'Administration',
    'Schularzt',
    'Schulwart',
    'IT-Support',
    'Bibliothek'
];

export function adminRoleCodeFromName(name) {
    const c = normStr(name)
        .toUpperCase()
        .replace(/\s+/g, '')
        .replace(/[^A-Z0-9ÄÖÜß-]/g, '')
        .slice(0, 24);
    return c;
}

function uniqueAdminRoleCode(desired, used) {
    let code = adminRoleCodeFromName(desired) || 'ROLLE';
    const usedSet = used instanceof Set ? used : new Set();
    if (!usedSet.has(code.toLowerCase())) return code;
    let i = 2;
    while (usedSet.has((code + String(i)).toLowerCase())) i += 1;
    return (code + String(i)).slice(0, 24);
}

export function defaultAdminRoleCatalog() {
    const used = new Set();
    return SW_ADMIN_DEFAULT_ROLES.map(function (name) {
        const code = uniqueAdminRoleCode(name, used);
        used.add(code.toLowerCase());
        return { code: code, name: name };
    });
}

export function personMatchesAdminRole(row, role) {
    if (!row || !role) return false;
    const r = normStr(row.role).toLowerCase();
    const dk = normStr(row.defaultKey).toLowerCase();
    const n = normStr(role.name).toLowerCase();
    const c = normStr(role.code).toLowerCase();
    if (n && (r === n || dk === n)) return true;
    if (c && (r === c || dk === c)) return true;
    return false;
}

export function normalizeAdminRoleCatalog(rolesIn, adminIn) {
    const seen = new Set();
    const roles = [];
    (Array.isArray(rolesIn) ? rolesIn : []).forEach(function (raw) {
        const name = normStr(raw && (raw.name || raw.role || raw.bezeichnung));
        let code = normStr(raw && raw.code).toUpperCase();
        if (!code && name) code = adminRoleCodeFromName(name);
        if (!code && !name) return;
        if (!code) code = uniqueAdminRoleCode(name, seen);
        const key = code.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        roles.push({ code: code, name: name || code });
    });
    (Array.isArray(adminIn) ? adminIn : []).forEach(function (a) {
        const name = normStr(a && (a.role || a.rolle || a.title));
        if (!name) return;
        const exists = roles.some(function (r) {
            return r.name.toLowerCase() === name.toLowerCase() || r.code.toLowerCase() === name.toLowerCase();
        });
        if (exists) return;
        const code = uniqueAdminRoleCode(name, seen);
        seen.add(code.toLowerCase());
        roles.push({ code: code, name: name });
    });
    return roles;
}

export function renameAdminRole(roles, admin, fromName, toName) {
    const from = normStr(fromName);
    const to = normStr(toName);
    const fromL = from.toLowerCase();
    const nextRoles = (Array.isArray(roles) ? roles : []).map(function (r) {
        const o = Object.assign({}, r);
        if (normStr(o.name).toLowerCase() === fromL || normStr(o.code).toLowerCase() === fromL) {
            o.name = to || o.name;
            if (to && (!o.code || normStr(o.code).toLowerCase() === fromL)) {
                o.code = adminRoleCodeFromName(to) || o.code;
            }
        }
        return o;
    });
    const nextAdmin = (Array.isArray(admin) ? admin : []).map(function (row) {
        const o = Object.assign({}, row);
        if (normStr(o.role).toLowerCase() === fromL) o.role = to;
        if (normStr(o.defaultKey).toLowerCase() === fromL) o.defaultKey = to;
        return o;
    });
    return { roles: normalizeAdminRoleCatalog(nextRoles, nextAdmin), admin: nextAdmin };
}

/** Kanonischer Standard-Slot nur bei gesetztem defaultKey (vermeidet Kollision mit freier Rolle „Direktion“). */
export function resolveAdminSlotFromRow(row) {
    if (!row) return null;
    const dk = normStr(row.defaultKey);
    if (!dk) return null;
    const m = SW_ADMIN_DEFAULT_ROLES.find(function (x) {
        return x.toLowerCase() === dk.toLowerCase();
    });
    return m || null;
}

export function migrateAdminRowDefaultKey(row) {
    const o = Object.assign({}, row);
    if (normStr(o.defaultKey)) return o;
    const rl = normStr(o.role);
    const m = SW_ADMIN_DEFAULT_ROLES.find(function (x) {
        return x.toLowerCase() === rl.toLowerCase();
    });
    if (m) o.defaultKey = m;
    return o;
}

export function isDirektionRole(roleRaw) {
    const r = normStr(roleRaw).toLowerCase();
    if (!r) return false;
    return r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1;
}

export function isDirektionAdminRow(row) {
    const slot = resolveAdminSlotFromRow(row);
    if (slot && slot.toLowerCase() === 'direktion') return true;
    return isDirektionRole(row && row.role);
}

export function getAdminDisplayRowsFromSettings(settings) {
    const adminArr = settings && Array.isArray(settings.admin) ? settings.admin : [];
    if (adminArr.length) {
        return adminArr.map(function (r) {
            return migrateAdminRowDefaultKey(r);
        });
    }
    return SW_ADMIN_DEFAULT_ROLES.map(function (slot) {
        return { defaultKey: slot, role: slot, name: '', email: '' };
    });
}

export function collectDirektionOwnerEmails(settings) {
    const out = [];
    const seen = new Set();
    const admin = settings && Array.isArray(settings.admin) ? settings.admin : [];
    admin.forEach(function (row) {
        if (!isDirektionAdminRow(row)) return;
        const em = normEmail(row && row.email);
        if (!em || em.indexOf('@') === -1) return;
        if (seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

export function collectAdminOwnerEmails(settings) {
    const out = [];
    const seen = new Set();
    const admin = settings && Array.isArray(settings.admin) ? settings.admin : [];
    admin.forEach(function (row) {
        const em = normEmail(row && row.email);
        if (!em || em.indexOf('@') === -1) return;
        if (seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

export function collectEmails(arr) {
    const out = [];
    const seen = new Set();
    (Array.isArray(arr) ? arr : []).forEach(function (row) {
        const em = normEmail(row && row.email);
        if (!em || em.indexOf('@') === -1) return;
        if (seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

export function randomTempPassword() {
    const u = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
    const l = 'abcdefghijkmnopqrstuvwxyz';
    const d = '23456789';
    const s = '!@#$%&*';
    function pick(set) {
        return set.charAt(Math.floor(Math.random() * set.length));
    }
    let pwd = pick(u) + pick(l) + pick(d) + pick(s);
    for (let i = 0; i < 12; i++) {
        pwd += pick(u + l + d);
    }
    return pwd;
}
