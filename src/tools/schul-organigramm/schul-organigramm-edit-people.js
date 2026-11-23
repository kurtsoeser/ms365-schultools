/**
 * People listing helpers for Schul-Organigramm edit.
 */
import { normStr, normEmail, normCode, isBlankPerson } from './schul-organigramm-edit-utils.js';
import {
    buildClassChatPlans,
    buildKvByKlasse
} from '../klassenchats/klassenchats-logic.js';
import { findArge, findClass } from './schul-organigramm-edit-find.js';
import {
    isArgeNode,
    isArgeAreaNode,
    isClassNode,
    isClassAreaNode,
    isElternClassNode,
    isKlassenchatClassNode,
    isSammelgruppeArea,
    isFunctionAggregateArea,
    functionAggregateDefForArea,
    listSchulOrgClassFunctions,
    listSchulOrgArgeFunctions,
    collectFunctionAggregatePeople,
    sgaScopeForArea,
    isSgaScopeArea,
    isSgaBoardArea,
    svSeatForNode,
    personForSvSeat,
    SCHUL_ORG_SV_SEATS,
    SCHUL_ORG_BOARD_REF,
    NODE,
    sammelgruppeKeyForArea,
    areaAcceptsDirectPeople,
    childrenOf,
    normalizeSchulOrganigramm,
    isAreaNode,
    canonicalSchulOrgClassCode
} from './schul-organigramm-logic.js';

export function roleCodeFromName(name) {
    if (typeof window.ms365TenantSettingsAdminRoleCodeFromName === 'function') {
        return window.ms365TenantSettingsAdminRoleCodeFromName(name) || 'ROLLE';
    }
    return (
        normStr(name)
            .toUpperCase()
            .replace(/\s+/g, '')
            .replace(/[^A-Z0-9ÄÖÜß-]/g, '')
            .slice(0, 24) || 'ROLLE'
    );
}

/**
 * administration ist beim Speichern die Quelle der Wahrheit (grouped).
 * Nach Änderungen an admin / adminRoles immer mitsynchronisieren.
 * @param {object} settings
 */
export function syncAdministrationFromAdmin(settings) {
    if (!settings) return;
    const roles = Array.isArray(settings.adminRoles) ? settings.adminRoles : [];
    const admin = Array.isArray(settings.admin) ? settings.admin : [];
    if (
        typeof window !== 'undefined' &&
        typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function'
    ) {
        settings.administration = window.ms365TenantSettingsAdminRolesAndAdminToGroups(roles, admin);
        return;
    }
    /** @type {Map<string, { code: string, name: string, people: object[], tier?: string, seatMode?: string }>} */
    const byName = new Map();
    roles.forEach(function (r) {
        const name = normStr(r && r.name) || normStr(r && r.code);
        if (!name) return;
        const group = {
            code: normStr(r && r.code) || name,
            name: name,
            people: []
        };
        if (r && r.tier) group.tier = normStr(r.tier);
        if (r && r.seatMode) group.seatMode = normStr(r.seatMode);
        byName.set(name.toLowerCase(), group);
    });
    admin.forEach(function (p) {
        const roleName = normStr(p && p.role);
        if (!roleName) return;
        const key = roleName.toLowerCase();
        if (!byName.has(key)) {
            byName.set(key, { code: roleName, name: roleName, people: [] });
        }
        const row = { name: normStr(p && p.name), email: normEmail(p && p.email) };
        const dk = normStr(p && p.defaultKey);
        if (dk) row.defaultKey = dk;
        byName.get(key).people.push(row);
    });
    settings.administration = Array.from(byName.values());
}

/**
 * @param {object} settings
 * @param {string} refKey – code oder name
 */
export function findAdminRole(settings, refKey) {
    const key = normStr(refKey).toUpperCase();
    const keyL = normStr(refKey).toLowerCase();
    const roles = Array.isArray(settings.adminRoles) ? settings.adminRoles : [];
    return (
        roles.find(function (r) {
            return normStr(r.code).toUpperCase() === key || normStr(r.name).toLowerCase() === keyL;
        }) || null
    );
}

/**
 * @param {object} settings
 * @param {object} role
 */
export function peopleForAdminRole(settings, role) {
    if (!role) return [];
    const name = normStr(role.name).toLowerCase();
    const code = normStr(role.code).toLowerCase();
    return (Array.isArray(settings.admin) ? settings.admin : []).filter(function (p) {
        const r = normStr(p && p.role).toLowerCase();
        const dk = normStr(p && p.defaultKey).toLowerCase();
        return (name && (r === name || dk === name)) || (code && (r === code || dk === code));
    });
}

/** Max. Personen-Chips auf einer Kachel (Rest über „+N“ / Bearbeiten). */
export const SCHUL_ORG_PEOPLE_PREVIEW_LIMIT = 3;

/**
 * @param {object} cls
 * @param {object} student
 */
export function studentMatchesClass(cls, student) {
    const k = normStr(student && student.klasse).toLowerCase();
    if (!k) return false;
    const code = normStr(cls && cls.code).toLowerCase();
    const name = normStr(cls && cls.name).toLowerCase();
    return (!!code && k === code) || (!!name && k === name);
}

/**
 * @param {object} settings
 * @param {object} cls
 */
export function studentsForClass(settings, cls) {
    if (!cls) return [];
    return (Array.isArray(settings.students) ? settings.students : []).filter(function (s) {
        return studentMatchesClass(cls, s) && (normStr(s && s.name) || normEmail(s && s.email));
    });
}

/**
 * Erziehungsberechtigte einer Klasse aus kanonischen guardianIds→guardians[]
 * (+ parentPairs als Fallback/Komfort-Projektion).
 * @param {object} settings
 * @param {string|{ code?: string, name?: string }} classRef
 * @returns {{ name: string, email: string }[]}
 */
export function guardiansForClass(settings, classRef) {
    const code =
        typeof classRef === 'string'
            ? normStr(classRef).toUpperCase()
            : normStr(classRef && classRef.code).toUpperCase();
    const nameL =
        typeof classRef === 'string' ? '' : normStr(classRef && classRef.name).toLowerCase();
    if (!code && !nameL) return [];
    /** @type {Map<string, { name: string, email: string }>} */
    const byEmail = new Map();
    const pushPair = function (p) {
        const em = normEmail(p && p.email);
        if (!em || em.indexOf('@') === -1) return;
        if (byEmail.has(em)) {
            const prev = byEmail.get(em);
            if (!prev.name && normStr(p && p.name)) prev.name = normStr(p.name);
            return;
        }
        byEmail.set(em, { name: normStr(p && p.name) || em, email: em });
    };
    const guardiansById = new Map();
    (Array.isArray(settings && settings.guardians) ? settings.guardians : []).forEach(function (g) {
        const id = normStr(g && g.id);
        if (id) guardiansById.set(id, g);
    });
    (Array.isArray(settings && settings.students) ? settings.students : []).forEach(function (s) {
        const k = normStr(s && s.klasse).toUpperCase();
        const kL = normStr(s && s.klasse).toLowerCase();
        if (!((code && k === code) || (nameL && kL === nameL))) return;
        const ids = Array.isArray(s && s.guardianIds) ? s.guardianIds : [];
        let resolved = 0;
        ids.forEach(function (gid) {
            const g = guardiansById.get(normStr(gid));
            if (!g) return;
            pushPair(g);
            resolved += 1;
        });
        // parentPairs: Fallback wenn keine/ungueltige guardianIds, sonst Komfort-Ergaenzung
        if (!resolved || Array.isArray(s && s.parentPairs)) {
            (Array.isArray(s && s.parentPairs) ? s.parentPairs : []).forEach(pushPair);
        }
    });
    return Array.from(byEmail.values()).sort(function (a, b) {
        return (a.name || a.email).localeCompare(b.name || b.email, 'de');
    });
}

/**
 * Lehrkräfte für den KlassenChat einer Klasse (Belegung + KV), Namen aus Stammdaten.
 * @param {object} settings
 * @param {string|{ code?: string, name?: string }} classRef
 * @returns {{ name: string, email: string, role: 'lehrer'|'kv', code: string }[]}
 */
/**
 * Lehrkräfte einer ARGE: Leitung + Lehrende der ARGE-Fächer (Unterrichtsbelegung).
 * @param {object} settings
 * @param {string|object} argeRef
 * @returns {{ name: string, email: string, role: 'head'|'lehrer', code: string }[]}
 */
/**
 * Besitzer einer M365-Gruppe → ARGE-Leitung in den Stammdaten.
 * Erster Besitzer (bevorzugt Lehrkraft) wird headName/headEmail.
 * @param {object} settings
 * @param {string} argeRef
 * @param {Array<{ displayName?: string, mail?: string, userPrincipalName?: string }>} owners
 * @returns {{ settings: object, changed: boolean, head: { name: string, email: string }|null }}
 */
export function applyArgeHeadFromGraphOwners(settings, argeRef, owners) {
    const s = settings && typeof settings === 'object' ? Object.assign({}, settings) : {};
    const code = normStr(argeRef).toUpperCase();
    if (!code) return { settings: s, changed: false, head: null };
    const list = Array.isArray(s.arges) ? s.arges.slice() : [];
    const idx = list.findIndex(function (a) {
        return normStr(a && a.code).toUpperCase() === code;
    });
    if (idx < 0) return { settings: s, changed: false, head: null };
    const teacherByEmail = new Map();
    (Array.isArray(s.teachers) ? s.teachers : []).forEach(function (t) {
        const em = normEmail(t && t.email);
        if (em) teacherByEmail.set(em, t);
    });
    /** @type {{ name: string, email: string }|null} */
    let head = null;
    const ownerList = Array.isArray(owners) ? owners : [];
    for (let i = 0; i < ownerList.length; i += 1) {
        const u = ownerList[i];
        const em = normEmail((u && (u.mail || u.userPrincipalName)) || '');
        if (!em) continue;
        const t = teacherByEmail.get(em);
        head = {
            email: em,
            name: (t && normStr(t.name)) || normStr(u && u.displayName) || em
        };
        if (t) break;
    }
    if (!head) return { settings: s, changed: false, head: null };
    const prev = list[idx] || {};
    if (normEmail(prev.headEmail) === head.email && normStr(prev.headName) === head.name) {
        return { settings: s, changed: false, head: head };
    }
    list[idx] = Object.assign({}, prev, {
        headName: head.name,
        headEmail: head.email
    });
    s.arges = list;
    return { settings: s, changed: true, head: head };
}

/**
 * Besitzer einer M365-Klassengruppe → Klassenvorstand in den Stammdaten.
 * Erster Besitzer (bevorzugt Lehrkraft) wird headName/headEmail.
 * @param {object} settings
 * @param {string} classRef
 * @param {Array<{ displayName?: string, mail?: string, userPrincipalName?: string }>} owners
 * @returns {{ settings: object, changed: boolean, head: { name: string, email: string }|null }}
 */
export function applyClassHeadFromGraphOwners(settings, classRef, owners) {
    const s = settings && typeof settings === 'object' ? Object.assign({}, settings) : {};
    const key = normStr(classRef);
    if (!key) return { settings: s, changed: false, head: null };
    const list = Array.isArray(s.classes) ? s.classes.slice() : [];
    const idx = list.findIndex(function (c) {
        return (
            normStr(c && c.code).toUpperCase() === key.toUpperCase() ||
            normStr(c && c.name).toLowerCase() === key.toLowerCase()
        );
    });
    if (idx < 0) return { settings: s, changed: false, head: null };
    const teacherByEmail = new Map();
    (Array.isArray(s.teachers) ? s.teachers : []).forEach(function (t) {
        const em = normEmail(t && t.email);
        if (em) teacherByEmail.set(em, t);
    });
    /** @type {{ name: string, email: string }|null} */
    let head = null;
    const ownerList = Array.isArray(owners) ? owners : [];
    for (let i = 0; i < ownerList.length; i += 1) {
        const u = ownerList[i];
        const em = normEmail((u && (u.mail || u.userPrincipalName)) || '');
        if (!em) continue;
        const t = teacherByEmail.get(em);
        head = {
            email: em,
            name: (t && normStr(t.name)) || normStr(u && u.displayName) || em
        };
        if (t) break;
    }
    if (!head) return { settings: s, changed: false, head: null };
    const prev = list[idx] || {};
    if (normEmail(prev.headEmail) === head.email && normStr(prev.headName) === head.name) {
        return { settings: s, changed: false, head: head };
    }
    list[idx] = Object.assign({}, prev, {
        headName: head.name,
        headEmail: head.email
    });
    s.classes = list;
    return { settings: s, changed: true, head: head };
}

export function teachersForArge(settings, argeRef) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const code =
        typeof argeRef === 'string'
            ? normStr(argeRef).toUpperCase()
            : normStr(argeRef && argeRef.code).toUpperCase();
    if (!code) return [];
    const arge = findArge(s, code);
    const subjectSet = new Set(
        (arge && Array.isArray(arge.subjects) ? arge.subjects : [])
            .map(function (x) {
                return normStr(x).toUpperCase();
            })
            .filter(Boolean)
    );
    /** @type {Map<string, { name: string, email: string }>} */
    const teacherByEmail = new Map();
    (Array.isArray(s.teachers) ? s.teachers : []).forEach(function (t) {
        const em = normEmail(t && t.email);
        if (!em) return;
        teacherByEmail.set(em, { name: normStr(t && t.name), email: em });
    });
    /** @type {Map<string, { name: string, email: string, role: 'head'|'lehrer', code: string }>} */
    const byEmail = new Map();
    const headEmail = arge ? normEmail(arge.headEmail) : '';
    const headName = arge ? normStr(arge.headName) : '';
    if (headEmail || headName) {
        const key = headEmail || 'name:' + headName.toLowerCase();
        byEmail.set(key, {
            name: headName || (teacherByEmail.get(headEmail) && teacherByEmail.get(headEmail).name) || headEmail,
            email: headEmail,
            role: 'head',
            code: ''
        });
    }
    if (subjectSet.size) {
        const belegung = s.unterrichtsbelegung && typeof s.unterrichtsbelegung === 'object'
            ? s.unterrichtsbelegung
            : null;
        const rows = belegung && Array.isArray(belegung.rows) ? belegung.rows : [];
        rows.forEach(function (r) {
            const fach = normStr(r && r.fach).toUpperCase();
            if (!fach || !subjectSet.has(fach)) return;
            const email = normEmail(r && r.lehrerEmail);
            if (!email) return;
            if (byEmail.has(email)) return;
            const t = teacherByEmail.get(email);
            byEmail.set(email, {
                name: normStr(r && r.lehrerName) || (t && t.name) || email,
                email: email,
                role: 'lehrer',
                code: normStr(r && r.lehrerCode).toUpperCase()
            });
        });
    }
    return Array.from(byEmail.values()).sort(function (a, b) {
        if (a.role === 'head' && b.role !== 'head') return -1;
        if (b.role === 'head' && a.role !== 'head') return 1;
        return (a.name || a.email).localeCompare(b.name || b.email, 'de');
    });
}

export function teachersForClassChat(settings, classRef) {
    const s = settings && typeof settings === 'object' ? settings : {};
    const rawRef =
        typeof classRef === 'string'
            ? classRef
            : (classRef && (classRef.code || classRef.refKey || classRef.name)) || '';
    const code =
        canonicalSchulOrgClassCode(rawRef) ||
        normStr(rawRef).toUpperCase();
    if (!code) return [];
    const codeUpper = code.toUpperCase();
    const snap = s.classChats && typeof s.classChats === 'object' ? s.classChats : null;
    const plans = buildClassChatPlans(s.unterrichtsbelegung || null, buildKvByKlasse(s.classes || []), {
        knownClassCodes: Array.isArray(s.classes) ? s.classes : [],
        namePattern: snap && Array.isArray(snap.namePattern) ? snap.namePattern : undefined,
        yearPrefix: snap && snap.yearPrefix ? String(snap.yearPrefix) : undefined
    });
    function sameKlasse(k) {
        const raw = normStr(k);
        if (!raw) return false;
        const canon = canonicalSchulOrgClassCode(raw) || raw.toUpperCase();
        return canon.toUpperCase() === codeUpper || raw.toUpperCase() === codeUpper;
    }
    const plan = plans.find(function (p) {
        return sameKlasse(p && p.klasse);
    });
    /** @type {Map<string, { name: string, email: string }>} */
    const teacherByEmail = new Map();
    (Array.isArray(s.teachers) ? s.teachers : []).forEach(function (t) {
        const em = normEmail(t && t.email);
        if (!em) return;
        teacherByEmail.set(em, { name: normStr(t && t.name), email: em });
    });
    const members = plan && Array.isArray(plan.members) ? plan.members : [];
    // Fallback: gespeicherte Chat-Mitglieder ohne aktuelle Belegung
    if (!members.length && snap && Array.isArray(snap.items)) {
        const item = snap.items.find(function (row) {
            return sameKlasse(row && row.klasse);
        });
        if (item && Array.isArray(item.memberEmails)) {
            return item.memberEmails
                .map(function (em) {
                    const email = normEmail(em);
                    if (!email) return null;
                    const t = teacherByEmail.get(email);
                    return {
                        name: (t && t.name) || email,
                        email: email,
                        role: /** @type {'lehrer'} */ ('lehrer'),
                        code: ''
                    };
                })
                .filter(Boolean)
                .sort(function (a, b) {
                    return (a.name || a.email).localeCompare(b.name || b.email, 'de');
                });
        }
    }
    // Fallback: nur KV aus Klassen-Stammdaten, wenn noch keine Belegung
    if (!members.length) {
        const cls = findClass(s, code) || findClass(s, rawRef);
        const kvEmail = cls ? normEmail(cls.headEmail) : '';
        if (kvEmail) {
            return [
                {
                    name: normStr(cls.headName) || kvEmail,
                    email: kvEmail,
                    role: /** @type {'kv'} */ ('kv'),
                    code: ''
                }
            ];
        }
    }
    return members
        .map(function (m) {
            const email = normEmail(m && m.email);
            if (!email) return null;
            const t = teacherByEmail.get(email);
            return {
                name: normStr(m && m.name) || (t && t.name) || email,
                email: email,
                role: m && m.role === 'kv' ? /** @type {'kv'} */ ('kv') : /** @type {'lehrer'} */ ('lehrer'),
                code: normStr(m && m.code).toUpperCase()
            };
        })
        .filter(Boolean)
        .sort(function (a, b) {
            if (a.role === 'kv' && b.role !== 'kv') return -1;
            if (b.role === 'kv' && a.role !== 'kv') return 1;
            return (a.name || a.email).localeCompare(b.name || b.email, 'de');
        });
}

/**
 * @param {{ kind: string, name?: string, email?: string, label?: string }} p
 */
export function personChipKey(p) {
    const em = normEmail(p && p.email);
    if (em) return String(p.kind || 'p') + ':e:' + em;
    return String(p.kind || 'p') + ':n:' + normStr(p && p.name).toLowerCase();
}

export {
    listPeopleForSchulOrgNode,
    listM365PeopleForArea,
    listMailboxDelegatePeopleForArea,
    previewPeopleForSchulOrgNode
} from './schul-organigramm-edit-people-node.js';
