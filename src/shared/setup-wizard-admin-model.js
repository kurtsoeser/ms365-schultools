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
    const groups = adminRolesAndAdminToGroups(roles, admin);
    const nextGroups = renameAdminRoleInGroups(groups, fromName, toName);
    return {
        roles: deriveAdminRolesFromGroups(nextGroups),
        admin: deriveAdminFromGroups(nextGroups)
    };
}

export function isAdministrationGrouped(entries) {
    return (
        Array.isArray(entries) &&
        entries.some(function (entry) {
            return entry && Array.isArray(entry.people);
        })
    );
}

function normalizeAdminPersonEntry(row) {
    const person = {
        name: normStr(row && row.name),
        email: normEmail(row && row.email)
    };
    const defaultKey = normStr(row && row.defaultKey);
    if (defaultKey) person.defaultKey = defaultKey;
    return person;
}

export function deriveAdminFromGroups(groups) {
    const out = [];
    (Array.isArray(groups) ? groups : []).forEach(function (group) {
        const roleName = normStr(group && group.name);
        (Array.isArray(group && group.people) ? group.people : []).forEach(function (personRow) {
            const person = normalizeAdminPersonEntry(personRow);
            if (!roleName && !person.name && !person.email) return;
            const row = { role: roleName, name: person.name, email: person.email };
            if (person.defaultKey) row.defaultKey = person.defaultKey;
            out.push(migrateAdminRowDefaultKey(row));
        });
    });
    return out;
}

export function deriveAdminRolesFromGroups(groups) {
    const out = [];
    const seen = new Set();
    (Array.isArray(groups) ? groups : []).forEach(function (group) {
        const name = normStr(group && group.name);
        let code = normStr(group && group.code).toUpperCase();
        if (!code && name) code = adminRoleCodeFromName(name);
        if (!code && !name) return;
        const key = (name || code).toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push({ code: code || adminRoleCodeFromName(name), name: name || code });
    });
    return out;
}

export function adminRolesAndAdminToGroups(rolesIn, adminIn) {
    const roles = normalizeAdminRoleCatalog(rolesIn, adminIn);
    const groupMap = new Map();
    roles.forEach(function (role) {
        const name = normStr(role && role.name);
        const code = normStr(role && role.code).toUpperCase() || adminRoleCodeFromName(name);
        if (!name && !code) return;
        groupMap.set(name.toLowerCase(), { code: code, name: name || code, people: [] });
    });
    (Array.isArray(adminIn) ? adminIn : []).forEach(function (row) {
        const migrated = migrateAdminRowDefaultKey(row);
        const roleName = normStr(migrated.role);
        if (!roleName) return;
        const key = roleName.toLowerCase();
        if (!groupMap.has(key)) {
            groupMap.set(key, {
                code: adminRoleCodeFromName(roleName),
                name: roleName,
                people: []
            });
        }
        const person = normalizeAdminPersonEntry(migrated);
        if (!person.name && !person.email) return;
        groupMap.get(key).people.push(person);
    });
    return Array.from(groupMap.values());
}

function flatKindEntriesToGroups(entries) {
    const administration = Array.isArray(entries) ? entries : [];
    const roles = normalizeAdminRoleCatalog(
        administration
            .filter(function (entry) {
                return entry && entry.kind === 'role';
            })
            .map(function (entry) {
                return { code: normStr(entry.code).toUpperCase(), name: normStr(entry.name) };
            }),
        administration
            .filter(function (entry) {
                return entry && entry.kind === 'person';
            })
            .map(function (entry) {
                return {
                    role: normStr(entry.role),
                    name: normStr(entry.name),
                    email: normEmail(entry.email),
                    defaultKey: normStr(entry.defaultKey)
                };
            })
    );
    const admin = administration
        .filter(function (entry) {
            return entry && entry.kind === 'person';
        })
        .map(function (entry) {
            const row = {
                role: normStr(entry.role),
                name: normStr(entry.name),
                email: normEmail(entry.email)
            };
            const defaultKey = normStr(entry.defaultKey);
            if (defaultKey) row.defaultKey = defaultKey;
            return migrateAdminRowDefaultKey(row);
        });
    return adminRolesAndAdminToGroups(roles, admin);
}

export function normalizeAdministrationGroups(entriesIn, rolesIn, adminIn) {
    if (isAdministrationGrouped(entriesIn)) {
        return normalizeGroupedAdministrationEntries(entriesIn);
    }
    if (
        Array.isArray(entriesIn) &&
        entriesIn.some(function (entry) {
            return entry && (entry.kind === 'role' || entry.kind === 'person');
        })
    ) {
        return flatKindEntriesToGroups(entriesIn);
    }
    return adminRolesAndAdminToGroups(rolesIn, adminIn);
}

export function normalizeGroupedAdministrationEntries(groupsIn) {
    const out = [];
    const seen = new Set();
    (Array.isArray(groupsIn) ? groupsIn : []).forEach(function (group) {
        const name = normStr(group && group.name);
        const code = normStr(group && group.code).toUpperCase() || adminRoleCodeFromName(name);
        if (!name && !code) return;
        const key = (name || code).toLowerCase();
        let target = out.find(function (entry) {
            return (entry.name || entry.code).toLowerCase() === key;
        });
        if (!target) {
            target = { code: code, name: name || code, people: [] };
            out.push(target);
            seen.add(key);
        }
        const peopleSeen = new Set(
            target.people.map(function (person) {
                return [person.name, person.email].join('\u0001').toLowerCase();
            })
        );
        (Array.isArray(group && group.people) ? group.people : []).forEach(function (personRow) {
            const person = normalizeAdminPersonEntry(personRow);
            if (!person.name && !person.email) return;
            const personKey = [person.name, person.email].join('\u0001').toLowerCase();
            if (peopleSeen.has(personKey)) return;
            peopleSeen.add(personKey);
            target.people.push(person);
        });
    });
    return out;
}

export function groupsToDisplayRows(groups) {
    const rows = [];
    (Array.isArray(groups) ? groups : []).forEach(function (group, groupIdx) {
        const people = Array.isArray(group.people) ? group.people : [];
        if (!people.length) {
            rows.push({
                groupIdx: groupIdx,
                personIdx: -1,
                code: group.code || '',
                name: group.name || '',
                personName: '',
                email: ''
            });
            return;
        }
        people.forEach(function (person, personIdx) {
            rows.push({
                groupIdx: groupIdx,
                personIdx: personIdx,
                code: group.code || '',
                name: group.name || '',
                personName: person.name || '',
                email: person.email || '',
                defaultKey: person.defaultKey || ''
            });
        });
    });
    return rows;
}

export function adminGroupsToLines(groups) {
    const lines = [];
    (Array.isArray(groups) ? groups : []).forEach(function (group) {
        const code = normStr(group && group.code);
        const name = normStr(group && group.name);
        const people = Array.isArray(group && group.people) ? group.people : [];
        if (!people.length) {
            lines.push([code, name, '', ''].join(';'));
            return;
        }
        people.forEach(function (person) {
            lines.push([code, name, normStr(person && person.name), normEmail(person && person.email)].join(';'));
        });
    });
    return lines.join('\n').trim();
}

export function renameAdminRoleInGroups(groups, fromName, toName) {
    const fromL = normStr(fromName).toLowerCase();
    const to = normStr(toName);
    return (Array.isArray(groups) ? groups : []).map(function (group) {
        const next = Object.assign({}, group, {
            people: (Array.isArray(group.people) ? group.people : []).map(function (person) {
                return Object.assign({}, person);
            })
        });
        if (normStr(next.name).toLowerCase() === fromL || normStr(next.code).toLowerCase() === fromL) {
            next.name = to || next.name;
            if (to && (!next.code || normStr(next.code).toLowerCase() === fromL)) {
                next.code = adminRoleCodeFromName(to) || next.code;
            }
        }
        return next;
    });
}

export function deriveAdministrationEntries(roles, admin) {
    return adminRolesAndAdminToGroups(roles, admin);
}

export function splitAdministrationEntries(entries) {
    if (isAdministrationGrouped(entries)) {
        return {
            roles: deriveAdminRolesFromGroups(entries),
            admin: deriveAdminFromGroups(entries)
        };
    }
    return {
        roles: deriveAdminRolesFromGroups(flatKindEntriesToGroups(entries)),
        admin: deriveAdminFromGroups(flatKindEntriesToGroups(entries))
    };
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
    const split =
        settings && Array.isArray(settings.administration)
            ? splitAdministrationEntries(settings.administration)
            : { admin: settings && Array.isArray(settings.admin) ? settings.admin : [] };
    const adminArr = split.admin;
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
    const split =
        settings && Array.isArray(settings.administration)
            ? splitAdministrationEntries(settings.administration)
            : { admin: settings && Array.isArray(settings.admin) ? settings.admin : [] };
    const admin = split.admin;
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
    const split =
        settings && Array.isArray(settings.administration)
            ? splitAdministrationEntries(settings.administration)
            : { admin: settings && Array.isArray(settings.admin) ? settings.admin : [] };
    const admin = split.admin;
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
