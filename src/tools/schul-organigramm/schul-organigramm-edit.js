/**
 * Schul-Organigramm bearbeiten (schreibt Stammdaten).
 * Phase 2: Rollen · Phase 3: ARGEs · Phase 4: Klassen/KV, Vertretung/SGA.
 * Custom-Verwaltungsgruppen → verwaltungAudienceGroups (+ Organigramm-Section).
 */
import {
    addCustomVerwaltungAudienceGroup,
    removeCustomVerwaltungAudienceGroup,
    normalizeVerwaltungAudienceGroups,
    normalizeAdminAudienceMemberships
} from '../../shared/administration-audience-groups.js';
import {
    buildClassChatPlans,
    buildKvByKlasse
} from '../klassenchats/klassenchats-logic.js';
import { resolveClassGraphGroupId } from '../../shared/membership-hygiene.js';
import {
    mountInlineEntraUserPicker,
    pickEntraUser
} from '../../shared/entra-user-picker.js';
import { mountInlineEntraGroupPicker } from '../../shared/entra-group-picker.js';
import {
    audienceSectionId,
    audienceIdForArea,
    sammelgruppeKeyForArea,
    isSammelgruppeArea,
    isArgeLeiterArea,
    collectArgeLeiterPeopleFromSettings,
    sgaScopeForArea,
    isSgaScopeArea,
    isSgaBoardArea,
    svSeatForNode,
    personForSvSeat,
    jahrgangKeyForArea,
    schulstufeFromClassCode,
    jahrgangAreaId,
    isElternClassNode,
    isKlassenchatClassNode,
    isArgeNode,
    isClassNode,
    isClassAreaNode,
    ensureSchulOrganigrammSgaStructure,
    canDeleteSchulOrganigrammArea,
    isOptionalDeletableBuiltinArea,
    withRemovedSchulOrganigrammArea,
    SCHUL_ORG_SV_SEATS,
    SCHUL_ORG_SGA_SV_AREA_ID,
    areaAcceptsDirectPeople,
    isRoleCapableSection,
    canPlaceSchulOrgNodeInSection,
    isAreaNode,
    normalizeSchulOrganigramm,
    upsertSchulOrganigrammAreaPerson,
    removeSchulOrganigrammAreaPerson,
    setSchulOrganigrammAreaM365,
    setSchulOrganigrammNodePosition,
    childrenOf,
    SCHUL_ORG_BOARD_REF,
    NODE,
    normalizeSchulOrgGroupKind,
    schulOrgGroupKindLabel,
    inferSchulOrgGroupKind,
    suggestSchulOrgMailNickname,
    isSchulOrgLocalPlaceholderEmail
} from './schul-organigramm-logic.js';

function normStr(v) {
    return String(v ?? '').trim();
}

function normEmail(v) {
    return String(v ?? '')
        .trim()
        .toLowerCase();
}

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * Einheitliche Dialog-Hülle (App-Dialog-/Picker-Stil).
 * @param {{ icon?: string, title: string, sub?: string, body: string, foot: string }} opts
 */
function editDialogHtml(opts) {
    const icon = opts.icon || 'bi-pencil-square';
    const title = opts.title || 'Bearbeiten';
    const sub = opts.sub || '';
    const body = opts.body || '';
    const foot = opts.foot || '';
    return (
        '<div class="modal-box schul-org-edit-box" role="dialog" aria-modal="true" aria-labelledby="schulOrgEditTitle" tabindex="-1">' +
        '<header class="schul-org-edit-head">' +
        '<div class="schul-org-edit-head-main">' +
        '<span class="schul-org-edit-icon" aria-hidden="true"><i class="bi ' +
        icon +
        '"></i></span>' +
        '<div class="schul-org-edit-titles">' +
        '<h3 id="schulOrgEditTitle" class="schul-org-edit-title">' +
        escapeHtml(title) +
        '</h3>' +
        (sub ? '<p class="schul-org-edit-sub">' + escapeHtml(sub) + '</p>' : '') +
        '</div></div>' +
        '<button type="button" class="schul-org-edit-close" data-schul-org-close aria-label="Schließen">' +
        '<i class="bi bi-x-lg" aria-hidden="true"></i></button>' +
        '</header>' +
        '<div class="schul-org-edit-body">' +
        body +
        '</div>' +
        '<footer class="schul-org-edit-foot">' +
        foot +
        '</footer></div>'
    );
}

function editFootHtml(delAttr, delLabel, saveAttr) {
    return (
        '<button type="button" class="btn btn-danger schul-org-edit-btn-danger" ' +
        delAttr +
        '><i class="bi bi-trash" aria-hidden="true"></i> ' +
        escapeHtml(delLabel) +
        '</button>' +
        '<span class="schul-org-edit-foot-spacer"></span>' +
        '<button type="button" class="btn schul-org-edit-btn-ghost" data-schul-org-close>Schließen</button>' +
        '<button type="button" class="btn btn-success schul-org-edit-btn-save" ' +
        saveAttr +
        '><i class="bi bi-check-lg" aria-hidden="true"></i> Speichern</button>'
    );
}

function loadSettings() {
    if (typeof window.ms365TenantSettingsLoad === 'function') {
        return window.ms365TenantSettingsLoad() || {};
    }
    return {};
}

function saveSettings(settings) {
    if (typeof window.ms365TenantSettingsSave === 'function') {
        return window.ms365TenantSettingsSave(settings);
    }
    return settings;
}

function toast(msg, kind) {
    if (typeof window.ms365AppDialogAlert === 'function') {
        window.ms365AppDialogAlert(msg, { kind: kind || 'info', title: 'Schul-Organigramm' });
        return;
    }
    if (typeof window.ms365Toast === 'function') window.ms365Toast(msg);
}

function confirmDlg(msg, opts) {
    if (typeof window.ms365AppDialogConfirm === 'function') {
        return window.ms365AppDialogConfirm(msg, opts || {});
    }
    return Promise.resolve(window.confirm(msg));
}

function roleCodeFromName(name) {
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
function syncAdministrationFromAdmin(settings) {
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
function studentMatchesClass(cls, student) {
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
 * Erziehungsberechtigte einer Klasse aus Schüler-parentPairs (und optional guardians[]).
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
    (Array.isArray(settings.students) ? settings.students : []).forEach(function (s) {
        const k = normStr(s && s.klasse).toUpperCase();
        const kL = normStr(s && s.klasse).toLowerCase();
        if (!((code && k === code) || (nameL && kL === nameL))) return;
        (Array.isArray(s && s.parentPairs) ? s.parentPairs : []).forEach(pushPair);
    });
    // Fallback: globale guardians-Liste ohne Klassenbezug nicht zuordnen
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
    const code =
        typeof classRef === 'string'
            ? normStr(classRef).toUpperCase()
            : normStr(classRef && classRef.code).toUpperCase();
    if (!code) return [];
    const snap = s.classChats && typeof s.classChats === 'object' ? s.classChats : null;
    const plans = buildClassChatPlans(s.unterrichtsbelegung || null, buildKvByKlasse(s.classes || []), {
        knownClassCodes: Array.isArray(s.classes) ? s.classes : [],
        namePattern: snap && Array.isArray(snap.namePattern) ? snap.namePattern : undefined,
        yearPrefix: snap && snap.yearPrefix ? String(snap.yearPrefix) : undefined
    });
    const plan = plans.find(function (p) {
        return normStr(p && p.klasse).toUpperCase() === code;
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
            return normStr(row && row.klasse).toUpperCase() === code;
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
function personChipKey(p) {
    const em = normEmail(p && p.email);
    if (em) return String(p.kind || 'p') + ':e:' + em;
    return String(p.kind || 'p') + ':n:' + normStr(p && p.name).toLowerCase();
}

/**
 * Personen für eine Organigramm-Kachel (Anzeige + DnD).
 * @param {object} settings
 * @param {object} node
 * @returns {{ kind: string, key: string, name: string, email: string, label: string, movable: boolean }[]}
 */
export function listPeopleForSchulOrgNode(settings, node) {
    if (!node || !node.type) return [];
    const out = [];
    const push = function (row) {
        const name = normStr(row.name);
        const email = normEmail(row.email);
        const vacant = !!row.vacant || !(name || email);
        if (!name && !email && !row.vacant && !row.seatLabel) return;
        const kind = row.kind || 'person';
        const seatLabel = normStr(row.seatLabel);
        const label = name || email || (seatLabel ? seatLabel + ' (offen)' : '');
        /** @type {{ kind: string, key: string, name: string, email: string, label: string, movable: boolean, scope?: string, klasse?: string, seatLabel?: string, vacant?: boolean }} */
        const item = {
            kind: kind,
            key: personChipKey({ kind: kind, name: name || seatLabel, email: email }),
            name: name,
            email: email,
            label: label,
            movable: row.movable !== false
        };
        if (seatLabel) item.seatLabel = seatLabel;
        if (vacant) item.vacant = true;
        if (row.scope) item.scope = String(row.scope);
        if (row.klasse) item.klasse = normStr(row.klasse);
        out.push(item);
    };

    if (node.type === 'role' && node.refKind === 'adminRole' && node.refKey) {
        const role = findAdminRole(settings, node.refKey);
        peopleForAdminRole(settings, role)
            .filter(function (p) {
                return !isBlankPerson(p);
            })
            .forEach(function (p) {
                push({ kind: 'admin', name: p.name, email: p.email });
            });
        return out;
    }

    if (isArgeNode(node) && node.refKey) {
        const members = teachersForArge(settings, node.refKey);
        if (!members.length) {
            push({
                kind: 'head',
                seatLabel: 'ARGE-Leitung',
                name: '',
                email: '',
                label: 'ARGE-Leitung (offen)',
                vacant: true,
                movable: true
            });
            return out;
        }
        members.forEach(function (t) {
            push({
                kind: t.role === 'head' ? 'head' : 'teacher',
                name: t.name,
                email: t.email,
                label:
                    t.role === 'head'
                        ? t.name || t.email || 'ARGE-Leitung (offen)'
                        : t.name || t.email,
                vacant: t.role === 'head' && !(t.name || t.email),
                movable: t.role === 'head',
                scope: 'teacher',
                seatLabel: t.role === 'head' ? 'ARGE-Leitung' : ''
            });
        });
        return out;
    }

    if (isClassNode(node) && node.refKey) {
        const cls = findClass(settings, node.refKey);
        const headName = cls ? normStr(cls.headName) : '';
        const headEmail = cls ? normEmail(cls.headEmail) : '';
        push({
            kind: 'head',
            seatLabel: 'Klassenvorstand',
            name: headName,
            email: headEmail,
            label: headName || headEmail || 'Klassenvorstand (offen)',
            vacant: !(headName || headEmail),
            movable: true
        });
        studentsForClass(settings, cls)
            .sort(function (a, b) {
                return normStr(a.name).localeCompare(normStr(b.name), 'de');
            })
            .forEach(function (s) {
                push({ kind: 'student', name: s.name, email: s.email });
            });
        return out;
    }

    if (isElternClassNode(node) && node.refKey) {
        guardiansForClass(settings, node.refKey).forEach(function (g) {
            push({
                kind: 'guardian',
                name: g.name,
                email: g.email,
                label: g.name || g.email,
                movable: false,
                scope: 'external'
            });
        });
        return out;
    }

    if (isKlassenchatClassNode(node) && node.refKey) {
        teachersForClassChat(settings, node.refKey).forEach(function (t) {
            push({
                kind: t.role === 'kv' ? 'head' : 'teacher',
                name: t.name,
                email: t.email,
                label: t.role === 'kv' ? (t.name || t.email) + ' (KV)' : t.name || t.email,
                movable: false,
                scope: 'teacher',
                seatLabel: t.role === 'kv' ? 'Klassenvorstand' : ''
            });
        });
        return out;
    }

    if (node.type === NODE.KURSTEAM || node.type === 'kursteam') {
        const meta = node.meta && typeof node.meta === 'object' ? node.meta : {};
        let headName = normStr(meta.lehrerName);
        let headEmail = normEmail(meta.lehrerEmail);
        if (!headEmail && meta.lehrerCode) {
            const code = normStr(meta.lehrerCode).toUpperCase();
            const t = (Array.isArray(settings.teachers) ? settings.teachers : []).find(function (row) {
                return normStr(row && row.code).toUpperCase() === code;
            });
            if (t) {
                headName = headName || normStr(t.name);
                headEmail = normEmail(t.email);
            }
        }
        push({
            kind: 'head',
            seatLabel: 'Besitzer',
            name: headName,
            email: headEmail,
            label: headName || headEmail || 'Besitzer (offen)',
            vacant: !(headName || headEmail),
            movable: false,
            scope: 'teacher'
        });
        return out;
    }

    if (node.type === 'council' && node.refKind === 'studentCouncil') {
        const seat = svSeatForNode(node);
        if (seat) {
            const person = personForSvSeat(settings, seat.seat);
            if (person) {
                push({
                    kind: 'council',
                    name: person.name,
                    email: person.email,
                    label: person.name || person.email,
                    movable: false,
                    scope: 'student',
                    klasse: person.klasse || '',
                    seatLabel: seat.label
                });
            }
            return out;
        }
        if (node.refKey === SCHUL_ORG_BOARD_REF || node.id === 'council_board') {
            SCHUL_ORG_SV_SEATS.forEach(function (s) {
                const person = personForSvSeat(settings, s.seat);
                if (!person) return;
                push({
                    kind: 'council',
                    name: person.name,
                    email: person.email,
                    label: s.label + ': ' + (person.name || person.email),
                    movable: false,
                    scope: 'student',
                    klasse: person.klasse || '',
                    seatLabel: s.label
                });
            });
            return out;
        }
        const row = findStudentCouncil(settings, node.refKey);
        if (row) push({ kind: 'council', name: row.name, email: row.email, movable: false });
        else if (node.meta) push({ kind: 'council', name: node.label, email: node.meta.email, movable: false });
        return out;
    }

    if (node.type === 'sga' && node.refKind === 'sga') {
        if (node.refKey === SCHUL_ORG_BOARD_REF || node.id === 'sga_board') {
            // Personen liegen in den Scope-Unterbereichen
            return out;
        }
        const row = findSga(settings, node.refKey);
        if (row) push({ kind: 'sga', name: row.name, email: row.email, movable: false });
        else if (node.meta) push({ kind: 'sga', name: node.label, email: node.meta.email, movable: false });
        return out;
    }

    if (isAreaNode(node) || node.type === NODE.PERSON) {
        if (node.type === NODE.PERSON && node.refKind === 'adminPerson') {
            push({
                kind: 'areaPerson',
                name: node.label,
                email: node.refKey,
                movable: true
            });
            return out;
        }
        if (isSgaBoardArea(node)) {
            return out;
        }
        if (isArgeLeiterArea(node)) {
            collectArgeLeiterPeopleFromSettings(settings).forEach(function (p) {
                push({
                    kind: 'head',
                    seatLabel: 'ARGE-Leitung',
                    name: p.name,
                    email: p.email,
                    label: p.label + (p.argeName ? ' · ' + p.argeName : ''),
                    vacant: false,
                    movable: false,
                    scope: 'teacher',
                    argeCode: p.argeCode,
                    argeName: p.argeName
                });
            });
            return out;
        }
        const sgaScope = sgaScopeForArea(node);
        if (sgaScope) {
            if (sgaScope === 'student' || node.id === SCHUL_ORG_SGA_SV_AREA_ID) {
                // Schülervertretung im SGA = drei Funktionen
                SCHUL_ORG_SV_SEATS.forEach(function (s) {
                    const person = personForSvSeat(settings, s.seat);
                    if (!person) return;
                    push({
                        kind: 'council',
                        name: person.name,
                        email: person.email,
                        label: s.label + ': ' + (person.name || person.email),
                        movable: false,
                        scope: 'student',
                        klasse: person.klasse || '',
                        seatLabel: s.label
                    });
                });
                return out;
            }
            (Array.isArray(settings.sga) ? settings.sga : [])
                .filter(function (row) {
                    if (!(normStr(row && row.name) || normEmail(row && row.email))) return false;
                    return (normalizeSgaScope(row.scope) || 'student') === sgaScope;
                })
                .sort(function (a, b) {
                    return normStr(a.name).localeCompare(normStr(b.name), 'de');
                })
                .forEach(function (row) {
                    push({
                        kind: 'sga',
                        name: row.name,
                        email: row.email,
                        label: normStr(row.name) || normEmail(row.email),
                        movable: false,
                        scope: sgaScope
                    });
                });
            return out;
        }
        const sammelKey = sammelgruppeKeyForArea(node);
        if (sammelKey === 'lehrer') {
            (Array.isArray(settings.teachers) ? settings.teachers : [])
                .filter(function (t) {
                    return normStr(t && t.name) || normEmail(t && t.email);
                })
                .sort(function (a, b) {
                    return normStr(a.name).localeCompare(normStr(b.name), 'de');
                })
                .forEach(function (t) {
                    push({
                        kind: 'teacher',
                        name: t.name,
                        email: t.email,
                        label: normStr(t.name) || normEmail(t.email),
                        movable: false,
                        scope: 'teacher'
                    });
                });
            return out;
        }
        if (sammelKey === 'schueler') {
            // Keine flache Schülerliste – Klassen liegen in den Jahrgangs-Bereichen darunter
            return out;
        }
        if (sammelKey === 'eltern') {
            // Keine flache Elternliste – Klassen liegen in den Jahrgangs-Bereichen darunter
            return out;
        }
        if (sammelKey === 'klassenchats') {
            // Keine flache Chat-Liste – Klassen liegen in den Jahrgangs-Bereichen darunter
            return out;
        }
        const seen = new Set();
        const audId = audienceIdForArea(node);
        if (audId) {
            normalizeAdminAudienceMemberships(settings.adminAudienceMemberships).forEach(function (m) {
                if (String(m.groupId) !== audId) return;
                if (seen.has(m.email)) return;
                seen.add(m.email);
                const adminHit = (Array.isArray(settings.admin) ? settings.admin : []).find(function (a) {
                    return normEmail(a && a.email) === m.email;
                });
                push({
                    kind: 'areaPerson',
                    name: m.name || (adminHit && adminHit.name) || '',
                    email: m.email,
                    movable: true
                });
            });
        }
        const layout = normalizeSchulOrganigramm(settings.schulOrganigramm);
        childrenOf(layout, node.id).forEach(function (ch) {
            if (ch.type !== NODE.PERSON || ch.refKind !== 'adminPerson') return;
            const em = normEmail(ch.refKey);
            if (!em || seen.has(em)) return;
            seen.add(em);
            push({
                kind: 'areaPerson',
                name: ch.label,
                email: em,
                movable: true
            });
        });
        out.sort(function (a, b) {
            return normStr(a.name).localeCompare(normStr(b.name), 'de');
        });
        return out;
    }

    return out;
}

/**
 * Personen für M365-Verknüpfung eines Bereichs (Gruppe/Postfach):
 * direkte Bereichspersonen, Rollen darunter sowie rekursiv alle Unterbereiche
 * (flach als Mitglieder/Berechtigte – keine verschachtelten M365-Gruppen).
 * @param {object} settings
 * @param {object} areaNode
 * @returns {Array<{ email: string, name: string, label: string, source: string }>}
 */
export function listM365PeopleForArea(settings, areaNode) {
    if (!areaNode || !isAreaNode(areaNode)) return [];
    /** @type {Map<string, { email: string, name: string, label: string, source: string }>} */
    const byMail = new Map();
    const pushPerson = function (p, source) {
        const email = normEmail(p && p.email);
        if (!email || isSchulOrgLocalPlaceholderEmail(email)) return;
        if (p && p.vacant) return;
        if (byMail.has(email)) return;
        const name = normStr(p && p.name);
        byMail.set(email, {
            email: email,
            name: name,
            label: name || email,
            source: source || 'area'
        });
    };

    const layout = normalizeSchulOrganigramm(settings && settings.schulOrganigramm);
    /** @type {Set<string>} */
    const visitedAreas = new Set();

    /**
     * @param {object} area
     * @param {boolean} nested
     */
    function walkArea(area, nested) {
        if (!area || !isAreaNode(area) || !area.id) return;
        const aid = String(area.id);
        if (visitedAreas.has(aid)) return;
        visitedAreas.add(aid);

        listPeopleForSchulOrgNode(settings, area).forEach(function (p) {
            pushPerson(p, nested ? 'nested' : 'area');
        });

        childrenOf(layout, aid).forEach(function (ch) {
            if (!ch) return;
            if (isAreaNode(ch)) {
                walkArea(ch, true);
                return;
            }
            const ct = String(ch.type || '').toLowerCase();
            if (ct !== NODE.ROLE && ct !== 'role') return;
            listPeopleForSchulOrgNode(settings, ch).forEach(function (p) {
                pushPerson(p, nested ? 'nested' : 'role');
            });
        });
    }

    walkArea(areaNode, false);

    return Array.from(byMail.values()).sort(function (a, b) {
        return normStr(a.name || a.email).localeCompare(normStr(b.name || b.email), 'de');
    });
}

/** @deprecated Alias – gleiche Liste wie listM365PeopleForArea */
export function listMailboxDelegatePeopleForArea(settings, areaNode) {
    return listM365PeopleForArea(settings, areaNode);
}

/**
 * @param {object} settings
 * @param {object} node
 * @param {number} [limit]
 */
export function previewPeopleForSchulOrgNode(settings, node, limit) {
    const max =
        typeof limit === 'number' && limit >= 0 ? limit : SCHUL_ORG_PEOPLE_PREVIEW_LIMIT;
    const all = listPeopleForSchulOrgNode(settings, node);
    return {
        all: all,
        shown: all.slice(0, max),
        more: Math.max(0, all.length - max),
        total: all.length
    };
}

function clearHeadOnRecord(rec) {
    if (!rec) return;
    rec.headName = '';
    rec.headEmail = '';
}

function setHeadOnRecord(rec, person) {
    if (!rec || !person) return;
    rec.headName = normStr(person.name) || normStr(person.label);
    rec.headEmail = normEmail(person.email);
}

function findAdminIndex(admin, person) {
    const em = normEmail(person && person.email);
    const name = normStr(person && person.name).toLowerCase();
    if (em) {
        const byMail = admin.findIndex(function (p) {
            return normEmail(p && p.email) === em;
        });
        if (byMail >= 0) return byMail;
    }
    if (name) {
        return admin.findIndex(function (p) {
            return normStr(p && p.name).toLowerCase() === name && !isBlankPerson(p);
        });
    }
    return -1;
}

function findStudentIndex(students, person, preferKlasse) {
    const em = normEmail(person && person.email);
    const name = normStr(person && person.name).toLowerCase();
    const klasse = normStr(preferKlasse).toLowerCase();
    if (em) {
        const byMail = students.findIndex(function (s) {
            return normEmail(s && s.email) === em;
        });
        if (byMail >= 0) return byMail;
    }
    return students.findIndex(function (s) {
        if (normStr(s && s.name).toLowerCase() !== name || !name) return false;
        if (!klasse) return true;
        return normStr(s && s.klasse).toLowerCase() === klasse;
    });
}

/**
 * Person von einer Kachel auf eine andere verschieben (schreibt Stammdaten-Objekt).
 * @param {object} settings
 * @param {{ kind: string, key?: string, name?: string, email?: string, label?: string, sourceNodeId?: string, sourceRefKind?: string, sourceRefKey?: string }} person
 * @param {object} targetNode
 * @returns {{ ok: boolean, message?: string, settings?: object }}
 */
export function moveSchulOrgPerson(settings, person, targetNode) {
    if (!settings || !person || !targetNode) {
        return { ok: false, message: 'Verschieben nicht möglich.' };
    }
    const kind = normStr(person.kind).toLowerCase();
    const targetType = normStr(targetNode.type).toLowerCase();
    const targetIsArge = isArgeNode(targetNode);
    const targetIsClass = isClassNode(targetNode);
    if (!kind || kind === 'council' || kind === 'sga') {
        return { ok: false, message: 'Dieser Eintrag lässt sich hier nicht verschieben – bitte über Bearbeiten ändern.' };
    }
    const targetIsArea = isAreaNode(targetNode) && !targetIsArge && !targetIsClass;
    if (!targetIsArea && !targetIsArge && !targetIsClass && !targetNode.refKey) {
        return { ok: false, message: 'Ziel ohne Stammdaten-Bezug.' };
    }

    const next = Object.assign({}, settings);
    next.admin = Array.isArray(settings.admin) ? settings.admin.map(function (p) {
        return Object.assign({}, p);
    }) : [];
    next.adminRoles = Array.isArray(settings.adminRoles) ? settings.adminRoles.slice() : [];
    next.arges = Array.isArray(settings.arges)
        ? settings.arges.map(function (a) {
              return Object.assign({}, a, {
                  subjects: Array.isArray(a.subjects) ? a.subjects.slice() : []
              });
          })
        : [];
    next.classes = Array.isArray(settings.classes)
        ? settings.classes.map(function (c) {
              return Object.assign({}, c);
          })
        : [];
    next.students = Array.isArray(settings.students)
        ? settings.students.map(function (s) {
              return Object.assign({}, s);
          })
        : [];
    next.adminAudienceMemberships = normalizeAdminAudienceMemberships(
        settings.adminAudienceMemberships
    ).map(function (m) {
        return Object.assign({}, m);
    });
    next.schulOrganigramm = normalizeSchulOrganigramm(settings.schulOrganigramm);

    const sourceRefKind = normStr(person.sourceRefKind);
    const sourceRefKey = normStr(person.sourceRefKey);
    let sameTarget = false;
    if (targetIsArea) {
        sameTarget = !!person.sourceNodeId && person.sourceNodeId === targetNode.id;
    } else if (sourceRefKind && sourceRefKey) {
        sameTarget =
            sourceRefKind === normStr(targetNode.refKind) &&
            sourceRefKey.toLowerCase() === normStr(targetNode.refKey).toLowerCase();
    }
    if (sameTarget) {
        return { ok: false, message: 'Bereits auf dieser Kachel.' };
    }

    // —— Admin-/Bereichsperson → Rolle (z. B. Eva aus Bereich auf „Personal“) ——
    if (
        (kind === 'admin' || kind === 'areaperson' || kind === 'person') &&
        targetType === 'role' &&
        normStr(targetNode.refKind) === 'adminRole'
    ) {
        const targetRole = findAdminRole(next, targetNode.refKey);
        if (!targetRole) return { ok: false, message: 'Zielrolle nicht gefunden.' };
        const email = normEmail(person.email);
        const name = normStr(person.name) || normStr(person.label);
        if (!email && !name) return { ok: false, message: 'Person unvollständig.' };
        let idx = findAdminIndex(next.admin, person);
        if (idx < 0) {
            next.admin.push({
                role: targetRole.name || targetRole.code,
                name: name || email,
                email: email,
                defaultKey: targetRole.name || targetRole.code
            });
        } else {
            next.admin[idx].role = targetRole.name || targetRole.code;
            next.admin[idx].defaultKey = targetRole.name || targetRole.code;
            if (name && !next.admin[idx].name) next.admin[idx].name = name;
            if (email && !next.admin[idx].email) next.admin[idx].email = email;
        }
        syncSeatMode(targetRole, peopleForAdminRole(next, targetRole));
        if (sourceRefKind === 'adminRole' && sourceRefKey) {
            const srcRole = findAdminRole(next, sourceRefKey);
            if (srcRole) syncSeatMode(srcRole, peopleForAdminRole(next, srcRole));
        }
        // Aus Quell-Bereich entfernen (Audience-Membership / Area-Person-Knoten)
        if (person.sourceNodeId) {
            removePersonFromAreaSettings(next, person.sourceNodeId, email);
        } else if (sourceRefKind === 'area' && sourceRefKey) {
            removePersonFromAreaSettings(next, sourceRefKey, email);
        }
        syncAdministrationFromAdmin(next);
        return { ok: true, settings: next, message: 'Person der Rolle zugeordnet.' };
    }

    // —— Admin-Person → ARGE-/Klassen-Leitung ——
    if (kind === 'admin' && (targetIsArge || targetIsClass)) {
        const list = targetIsArge ? next.arges : next.classes;
        const findFn = targetIsArge ? findArge : findClass;
        const rec = findFn(next, targetNode.refKey);
        if (!rec) return { ok: false, message: 'Ziel nicht gefunden.' };
        setHeadOnRecord(rec, person);
        if (targetIsArge) next.arges = list;
        else next.classes = list;
        return {
            ok: true,
            settings: next,
            message: targetIsArge ? 'Als ARGE-Leitung gesetzt.' : 'Als Klassenvorstand gesetzt.'
        };
    }

    // —— Leitung (ARGE/Klasse) → andere Leitung gleichen Typs ——
    if (kind === 'head' && (targetIsArge || targetIsClass)) {
        const findFn = targetIsArge ? findArge : findClass;
        const wantKind = targetIsArge ? 'arge' : 'class';
        if (sourceRefKind && sourceRefKind !== wantKind) {
            return { ok: false, message: 'Leitung nur innerhalb gleicher Bereiche verschieben (oder zuerst in Rolle).' };
        }
        const target = findFn(next, targetNode.refKey);
        if (!target) return { ok: false, message: 'Ziel nicht gefunden.' };
        if (sourceRefKey) {
            const src = findFn(next, sourceRefKey);
            if (src) clearHeadOnRecord(src);
        }
        setHeadOnRecord(target, person);
        return {
            ok: true,
            settings: next,
            message: targetIsArge ? 'ARGE-Leitung verschoben.' : 'Klassenvorstand verschoben.'
        };
    }

    // —— Leitung → Verwaltungsrolle ——
    if (kind === 'head' && targetType === 'role' && targetNode.refKind === 'adminRole') {
        const targetRole = findAdminRole(next, targetNode.refKey);
        if (!targetRole) return { ok: false, message: 'Zielrolle nicht gefunden.' };
        let idx = findAdminIndex(next.admin, person);
        if (idx < 0) {
            next.admin.push({
                role: targetRole.name || targetRole.code,
                name: normStr(person.name) || normStr(person.label),
                email: normEmail(person.email),
                defaultKey: targetRole.name || targetRole.code
            });
        } else {
            next.admin[idx].role = targetRole.name || targetRole.code;
            next.admin[idx].defaultKey = targetRole.name || targetRole.code;
            if (!next.admin[idx].name) next.admin[idx].name = normStr(person.name);
        }
        if (sourceRefKind === 'arge' && sourceRefKey) {
            const src = findArge(next, sourceRefKey);
            if (src) clearHeadOnRecord(src);
        } else if (sourceRefKind === 'class' && sourceRefKey) {
            const src = findClass(next, sourceRefKey);
            if (src) clearHeadOnRecord(src);
        }
        syncSeatMode(targetRole, peopleForAdminRole(next, targetRole));
        syncAdministrationFromAdmin(next);
        return { ok: true, settings: next, message: 'Person der Rolle zugeordnet.' };
    }

    // —— Schüler → andere Klasse ——
    if (kind === 'student' && targetIsClass && targetNode.refKind === 'class') {
        const targetCls = findClass(next, targetNode.refKey);
        if (!targetCls) return { ok: false, message: 'Zielklasse nicht gefunden.' };
        const preferKlasse =
            sourceRefKind === 'class' && sourceRefKey
                ? (findClass(settings, sourceRefKey) || {}).code || sourceRefKey
                : '';
        const idx = findStudentIndex(next.students, person, preferKlasse);
        if (idx < 0) return { ok: false, message: 'Schüler:in in den Stammdaten nicht gefunden.' };
        next.students[idx].klasse = targetCls.code || targetCls.name;
        return { ok: true, settings: next, message: 'Schüler:in der Klasse zugeordnet.' };
    }

    // —— Person → Area (direkter Slot / Audience-Membership) ——
    if (
        targetIsArea &&
        (kind === 'admin' || kind === 'head' || kind === 'areaperson' || kind === 'person')
    ) {
        const email = normEmail(person.email);
        const name = normStr(person.name) || normStr(person.label);
        if (!email) return { ok: false, message: 'Person ohne E-Mail.' };
        const placed = placePersonOnArea(next, targetNode, { email: email, name: name });
        if (!placed.ok) return placed;
        // Slot aus Quell-Area entfernen
        if (sourceRefKind === 'area' && sourceRefKey && sourceRefKey !== targetNode.id) {
            removePersonFromAreaSettings(next, sourceRefKey, email);
        } else if (person.sourceNodeId && person.sourceNodeId !== targetNode.id) {
            const srcNode = (next.schulOrganigramm.nodes || []).find(function (n) {
                return n.id === person.sourceNodeId;
            });
            if (srcNode && isAreaNode(srcNode)) {
                removePersonFromAreaSettings(next, srcNode.id, email);
            }
        }
        return {
            ok: true,
            settings: next,
            message: 'Person dem Bereich zugeordnet.'
        };
    }

    return {
        ok: false,
        message: 'Diese Verschiebung ist nicht vorgesehen. Mitglieder bitte über Bearbeiten pflegen.'
    };
}

/**
 * @param {object} settings
 * @param {object} areaNode
 * @param {{ email: string, name?: string }} person
 */
function placePersonOnArea(settings, areaNode, person) {
    if (!areaAcceptsDirectPeople(areaNode)) {
        return { ok: false, message: 'Bereich nimmt keine Personen auf.' };
    }
    if (isSammelgruppeArea(areaNode)) {
        return {
            ok: false,
            message:
                sammelgruppeKeyForArea(areaNode) === 'lehrer'
                    ? 'Lehrkräfte bitte über die Lehrerliste in den Stammdaten pflegen.'
                    : 'Schüler:innen bitte über die Schülerliste in den Stammdaten pflegen.'
        };
    }
    if (isSgaBoardArea(areaNode)) {
        return {
            ok: false,
            message: 'Personen bitte in Schüler, Lehrer oder Eltern (unter SGA) hinzufügen.'
        };
    }
    const email = normEmail(person && person.email);
    const name = normStr(person && person.name);
    if (!email) return { ok: false, message: 'E-Mail fehlt.' };
    const sgaScope = sgaScopeForArea(areaNode);
    if (sgaScope) {
        if (!Array.isArray(settings.sga)) settings.sga = [];
        const exists = settings.sga.some(function (row) {
            return (
                normEmail(row && row.email) === email &&
                (normalizeSgaScope(row.scope) || 'student') === sgaScope
            );
        });
        if (!exists) {
            settings.sga.push({ scope: sgaScope, name: name || email, email: email });
        }
        return { ok: true };
    }
    const audId = audienceIdForArea(areaNode);
    if (audId) {
        const mem = normalizeAdminAudienceMemberships(settings.adminAudienceMemberships);
        const exists = mem.some(function (m) {
            return m.groupId === audId && m.email === email;
        });
        if (!exists) {
            const row = { groupId: audId, email: email };
            if (name) row.name = name;
            mem.push(row);
        } else {
            mem.forEach(function (m) {
                if (m.groupId === audId && m.email === email && name) m.name = name;
            });
        }
        settings.adminAudienceMemberships = mem;
        // Person ggf. auch in admin-Liste sichern (ohne Rolle zu überschreiben)
        if (!Array.isArray(settings.admin)) settings.admin = [];
        const idx = findAdminIndex(settings.admin, { email: email, name: name });
        if (idx < 0) {
            settings.admin.push({ role: '', name: name || email, email: email, defaultKey: '' });
        } else if (name && !settings.admin[idx].name) {
            settings.admin[idx].name = name;
        }
        return { ok: true };
    }
    settings.schulOrganigramm = upsertSchulOrganigrammAreaPerson(
        settings.schulOrganigramm,
        areaNode.id,
        { email: email, name: name }
    );
    return { ok: true };
}

/**
 * @param {object} settings
 * @param {string} areaId
 * @param {string} email
 */
function removePersonFromAreaSettings(settings, areaId, email) {
    const em = normEmail(email);
    if (!em || !areaId) return;
    const layout = normalizeSchulOrganigramm(settings.schulOrganigramm);
    const area = layout.nodes.find(function (n) {
        return n.id === areaId;
    });
    const sgaScope = sgaScopeForArea(area);
    if (sgaScope) {
        settings.sga = (Array.isArray(settings.sga) ? settings.sga : []).filter(function (row) {
            if (normEmail(row && row.email) !== em) return true;
            return (normalizeSgaScope(row.scope) || 'student') !== sgaScope;
        });
        return;
    }
    const audId = audienceIdForArea(area);
    if (audId) {
        settings.adminAudienceMemberships = normalizeAdminAudienceMemberships(
            settings.adminAudienceMemberships
        ).filter(function (m) {
            return !(m.groupId === audId && m.email === em);
        });
    }
    settings.schulOrganigramm = removeSchulOrganigrammAreaPerson(
        settings.schulOrganigramm,
        em,
        areaId
    );
}

/**
 * Person einem Bereich zuordnen (Entra-Picker / manuell).
 * @param {object} settings
 * @param {object} areaNode
 * @param {{ email: string, name?: string }} person
 */
export function addSchulOrgPersonToArea(settings, areaNode, person) {
    if (!settings || !areaNode) return { ok: false, message: 'Bereich fehlt.' };
    const next = Object.assign({}, settings);
    next.admin = Array.isArray(settings.admin)
        ? settings.admin.map(function (p) {
              return Object.assign({}, p);
          })
        : [];
    next.adminAudienceMemberships = normalizeAdminAudienceMemberships(
        settings.adminAudienceMemberships
    ).map(function (m) {
        return Object.assign({}, m);
    });
    next.schulOrganigramm = normalizeSchulOrganigramm(settings.schulOrganigramm);
    const placed = placePersonOnArea(next, areaNode, person);
    if (!placed.ok) return placed;
    return { ok: true, settings: next, message: 'Person hinzugefügt.' };
}

/**
 * @param {object} settings
 * @param {object} areaNode
 * @param {string} email
 */
export function removeSchulOrgPersonFromArea(settings, areaNode, email) {
    if (!settings || !areaNode) return { ok: false, message: 'Bereich fehlt.' };
    const next = Object.assign({}, settings);
    next.adminAudienceMemberships = normalizeAdminAudienceMemberships(
        settings.adminAudienceMemberships
    ).map(function (m) {
        return Object.assign({}, m);
    });
    next.schulOrganigramm = normalizeSchulOrganigramm(settings.schulOrganigramm);
    removePersonFromAreaSettings(next, areaNode.id, email);
    return { ok: true, settings: next, message: 'Person entfernt.' };
}

function personMatchesRole(p, role) {
    const name = normStr(role && role.name).toLowerCase();
    const code = normStr(role && role.code).toLowerCase();
    const r = normStr(p && p.role).toLowerCase();
    const dk = normStr(p && p.defaultKey).toLowerCase();
    return (name && (r === name || dk === name)) || (code && (r === code || dk === code));
}

function isBlankPerson(p) {
    return !normStr(p && p.name) && !normEmail(p && p.email);
}

function syncSeatMode(role, people) {
    const filled = people.filter(function (p) {
        return !isBlankPerson(p);
    }).length;
    if (filled >= 2) role.seatMode = 'multi';
    else if (!role.seatMode) role.seatMode = 'single';
}

function patchSchulOrganigrammRefLabel(settings, type, oldRefKey, nextLabel, nextRefKey) {
    const layout = settings.schulOrganigramm;
    if (!layout || !Array.isArray(layout.nodes)) return;
    const oldKey = normStr(oldRefKey).toLowerCase();
    layout.nodes = layout.nodes.map(function (n) {
        if (n.type !== type) return n;
        if (normStr(n.refKey).toLowerCase() !== oldKey) return n;
        const next = Object.assign({}, n, { label: nextLabel || n.label });
        if (nextRefKey != null && String(nextRefKey) !== '') next.refKey = normStr(nextRefKey);
        return next;
    });
    settings.schulOrganigramm = layout;
}

function removeSchulOrganigrammByRef(settings, type, refKey) {
    const layout = settings.schulOrganigramm;
    if (!layout || !Array.isArray(layout.nodes)) return;
    const key = normStr(refKey).toLowerCase();
    settings.schulOrganigramm = {
        version: 1,
        nodes: layout.nodes.filter(function (n) {
            return !(n.type === type && normStr(n.refKey).toLowerCase() === key);
        })
    };
}

function sgaScopeLabel(scope) {
    const s = normStr(scope).toLowerCase();
    if (s === 'teacher') return 'Lehrer';
    if (s === 'student') return 'Schüler';
    if (s === 'external') return 'Extern';
    return s || '–';
}

function normalizeSgaScope(raw) {
    const s = normStr(raw).toLowerCase();
    if (s === 'lehrer' || s === 'teacher') return 'teacher';
    if (s === 'schueler' || s === 'schüler' || s === 'student') return 'student';
    if (s === 'extern' || s === 'external' || s === 'eltern') return 'external';
    return '';
}

/**
 * @param {HTMLSelectElement} selectEl
 * @param {object[]} teachers
 */
function fillTeacherSelect(selectEl, teachers) {
    (teachers || [])
        .filter(function (t) {
            return normEmail(t && t.email).indexOf('@') !== -1;
        })
        .sort(function (a, b) {
            return normStr(a.name).localeCompare(normStr(b.name), 'de');
        })
        .forEach(function (t) {
            const opt = document.createElement('option');
            opt.value = normEmail(t.email);
            opt.textContent =
                (normStr(t.name) || normEmail(t.email)) +
                (t.code ? ' (' + normStr(t.code) + ')' : '') +
                ' – ' +
                normEmail(t.email);
            selectEl.appendChild(opt);
        });
}

/**
 * Lehrerliste ↔ Name/E-Mail-Felder für Leitung/KV verdrahten.
 * @param {{ teacherEl: HTMLSelectElement, headNameEl: HTMLInputElement, headEmailEl: HTMLInputElement, hintEl: Element|null, teachers: object[], vacantText: string }} opts
 */
function wireTeacherHeadFields(opts) {
    const teacherEl = opts.teacherEl;
    const headNameEl = opts.headNameEl;
    const headEmailEl = opts.headEmailEl;
    const hintEl = opts.hintEl;
    const teachers = opts.teachers || [];
    const vacantText = opts.vacantText || 'Keine Person gesetzt.';

    function syncTeacherSelect() {
        const em = normEmail(headEmailEl.value);
        teacherEl.value = em && teachers.some(function (t) {
            return normEmail(t.email) === em;
        })
            ? em
            : '';
    }

    function updateHeadHint() {
        if (!hintEl) return;
        const em = normEmail(headEmailEl.value);
        if (!em) {
            hintEl.textContent = vacantText;
            return;
        }
        const t = findTeacherByEmail({ teachers: teachers }, em);
        if (t) {
            hintEl.textContent =
                'Verknüpft mit Lehrerliste: ' +
                (t.name || t.code || em) +
                (t.code ? ' · ' + t.code : '');
        } else {
            hintEl.textContent =
                'E-Mail gesetzt, aber kein Treffer in der Lehrerliste (trotzdem gültig).';
        }
    }

    teacherEl.addEventListener('change', function () {
        const em = normEmail(teacherEl.value);
        if (!em) return;
        const t = findTeacherByEmail({ teachers: teachers }, em);
        headEmailEl.value = em;
        if (t && t.name) headNameEl.value = t.name;
        updateHeadHint();
    });
    headEmailEl.addEventListener('input', function () {
        syncTeacherSelect();
        updateHeadHint();
    });
    headNameEl.addEventListener('input', updateHeadHint);
    syncTeacherSelect();
    updateHeadHint();
    return { syncTeacherSelect: syncTeacherSelect, updateHeadHint: updateHeadHint };
}

/**
 * @param {object} settings
 * @param {string} refKey
 */
export function findArge(settings, refKey) {
    const key = normStr(refKey).toUpperCase();
    const keyL = normStr(refKey).toLowerCase();
    const arges = Array.isArray(settings.arges) ? settings.arges : [];
    return (
        arges.find(function (a) {
            return normStr(a.code).toUpperCase() === key || normStr(a.name).toLowerCase() === keyL;
        }) || null
    );
}

/**
 * @param {object} settings
 * @param {string} email
 */
export function findTeacherByEmail(settings, email) {
    const em = normEmail(email);
    if (!em || em.indexOf('@') === -1) return null;
    const teachers = Array.isArray(settings.teachers) ? settings.teachers : [];
    return (
        teachers.find(function (t) {
            return normEmail(t && t.email) === em;
        }) || null
    );
}

/**
 * @param {object} settings
 * @param {string} refKey – Klassen-Code oder Name
 */
export function findClass(settings, refKey) {
    const key = normStr(refKey).toUpperCase();
    const keyL = normStr(refKey).toLowerCase();
    const classes = Array.isArray(settings.classes) ? settings.classes : [];
    return (
        classes.find(function (c) {
            return normStr(c.code).toUpperCase() === key || normStr(c.name).toLowerCase() === keyL;
        }) || null
    );
}

/**
 * @param {object} settings
 * @param {string} refKey – E-Mail oder Name
 */
export function findStudentCouncil(settings, refKey) {
    const key = normStr(refKey);
    const keyL = key.toLowerCase();
    const list = Array.isArray(settings.studentCouncil) ? settings.studentCouncil : [];
    if (keyL.indexOf('@') !== -1) {
        const byMail = list.find(function (s) {
            return normEmail(s.email) === keyL;
        });
        if (byMail) return byMail;
    }
    return (
        list.find(function (s) {
            return normStr(s.name).toLowerCase() === keyL;
        }) || null
    );
}

/**
 * @param {object} settings
 * @param {string} refKey – E-Mail oder Name
 */
export function findSga(settings, refKey) {
    const key = normStr(refKey);
    const keyL = key.toLowerCase();
    const list = Array.isArray(settings.sga) ? settings.sga : [];
    if (keyL.indexOf('@') !== -1) {
        const byMail = list.find(function (s) {
            return normEmail(s.email) === keyL;
        });
        if (byMail) return byMail;
    }
    return (
        list.find(function (s) {
            return normStr(s.name).toLowerCase() === keyL;
        }) || null
    );
}

/** @type {HTMLElement|null} */
let activeOverlay = null;
/** @type {function(): void}|null */
let onChangedCb = null;
/** Zuletzt genutzter Bereich für neue Funktionen (Rechtsklick / Toolbar). */
let lastRoleParentId = 'sec_verwaltung';
const LAST_ROLE_PARENT_KEY = 'ms365-schul-org-last-role-parent-v1';

function readLastRoleParentId() {
    try {
        if (typeof localStorage !== 'undefined') {
            const v = normStr(localStorage.getItem(LAST_ROLE_PARENT_KEY));
            if (v) lastRoleParentId = v;
        }
    } catch {
        /* ignore */
    }
    return lastRoleParentId;
}

/**
 * @param {string} areaId
 */
export function rememberSchulOrgRoleParent(areaId) {
    const id = normStr(areaId);
    if (!id) return;
    lastRoleParentId = id;
    try {
        if (typeof localStorage !== 'undefined') {
            localStorage.setItem(LAST_ROLE_PARENT_KEY, id);
        }
    } catch {
        /* ignore */
    }
}

/**
 * @returns {string}
 */
export function getLastSchulOrgRoleParent() {
    return readLastRoleParentId();
}

function closeOverlay() {
    document.removeEventListener('keydown', onEsc, true);
    if (activeOverlay && activeOverlay.parentNode) activeOverlay.parentNode.removeChild(activeOverlay);
    activeOverlay = null;
}

function onEsc(ev) {
    if (ev.key === 'Escape') {
        ev.preventDefault();
        closeOverlay();
    }
}

/**
 * Entra-Personensuche im offenen Dialog einbetten (kein zweites Overlay).
 * @param {HTMLElement|null} host
 * @param {{ api: { destroy: () => void, focus: () => void }|null }} state
 * @param {{ multi?: boolean, hint?: string, excludeEmails?: string[]|Set<string>, onPick?: Function, onApply?: Function }} opts
 * @param {{ api: { destroy: () => void }|null }[]} [closeOthers]
 */
/**
 * Inline-Picker als Overlay im Dialog-Body (ohne Höhenwachstum).
 * @param {HTMLElement|null} host
 * @param {{
 *   api: { destroy: () => void }|null,
 *   _eupHost?: HTMLElement|null,
 *   _eupHome?: { parent: Node, next: Node|null }|null
 * }} state
 * @param {boolean} open
 */
function setInlinePickerOverlay(host, state, open) {
    if (!state) return;
    const el = host || state._eupHost;
    if (!el) return;
    const box = el.closest ? el.closest('.schul-org-edit-box') : null;
    const body = box ? box.querySelector('.schul-org-edit-body') : null;
    if (open) {
        state._eupHost = el;
        if (box) box.classList.add('is-inline-picker-open');
        if (body && el.parentElement !== body) {
            state._eupHome = { parent: el.parentElement, next: el.nextSibling };
            body.appendChild(el);
        }
        return;
    }
    if (box) box.classList.remove('is-inline-picker-open');
    const home = state._eupHome;
    if (home && home.parent && el.parentElement !== home.parent) {
        try {
            home.parent.insertBefore(el, home.next);
        } catch {
            home.parent.appendChild(el);
        }
    }
    state._eupHome = null;
    state._eupHost = null;
}

async function toggleInlineEntraUserPicker(host, state, opts, closeOthers) {
    if (!host || !state) return;
    if (state.api) {
        state.api.destroy();
        state.api = null;
        setInlinePickerOverlay(host, state, false);
        return;
    }
    (closeOthers || []).forEach(function (s) {
        if (s && s.api && s !== state) {
            const otherHost = s._eupHost || null;
            s.api.destroy();
            s.api = null;
            setInlinePickerOverlay(otherHost, s, false);
        }
    });
    const options = opts || {};
    setInlinePickerOverlay(host, state, true);
    state.api = mountInlineEntraUserPicker(host, {
        multi: !!options.multi,
        hint: options.hint,
        excludeEmails: options.excludeEmails,
        onPick: function (user) {
            state.api = null;
            setInlinePickerOverlay(host, state, false);
            if (typeof options.onPick === 'function') options.onPick(user);
        },
        onApply: function (users) {
            state.api = null;
            setInlinePickerOverlay(host, state, false);
            if (typeof options.onApply === 'function') options.onApply(users);
        },
        onCancel: function () {
            state.api = null;
            setInlinePickerOverlay(host, state, false);
        }
    });
}

/**
 * Entra-Gruppensuche im offenen Dialog einbetten (kein zweites Overlay).
 * @param {HTMLElement|null} host
 * @param {{ api: { destroy: () => void, focus: () => void }|null }} state
 * @param {{ hint?: string, onPick?: Function }} opts
 * @param {{ api: { destroy: () => void }|null }[]} [closeOthers]
 */
async function toggleInlineEntraGroupPicker(host, state, opts, closeOthers) {
    if (!host || !state) return;
    if (state.api) {
        state.api.destroy();
        state.api = null;
        setInlinePickerOverlay(host, state, false);
        return;
    }
    (closeOthers || []).forEach(function (s) {
        if (s && s.api && s !== state) {
            const otherHost = s._eupHost || null;
            s.api.destroy();
            s.api = null;
            setInlinePickerOverlay(otherHost, s, false);
        }
    });
    const options = opts || {};
    setInlinePickerOverlay(host, state, true);
    state.api = mountInlineEntraGroupPicker(host, {
        hint: options.hint,
        onPick: function (group) {
            state.api = null;
            setInlinePickerOverlay(host, state, false);
            if (typeof options.onPick === 'function') options.onPick(group);
        },
        onCancel: function () {
            state.api = null;
            setInlinePickerOverlay(host, state, false);
        }
    });
}

/**
 * @param {{ onChanged?: function(): void }} [deps]
 */
export function createSchulOrganigrammEditor(deps) {
    onChangedCb = deps && typeof deps.onChanged === 'function' ? deps.onChanged : null;

    function notify() {
        if (onChangedCb) onChangedCb();
    }

    /**
     * @param {string} refKey
     */
    function openRoleEditor(refKey) {
        closeOverlay();
        const settings = loadSettings();
        let role = findAdminRole(settings, refKey);
        if (!role) {
            toast('Rolle nicht in den Stammdaten gefunden.', 'warning');
            return;
        }
        // Arbeitskopie
        role = Object.assign({}, role);
        let people = peopleForAdminRole(settings, role).map(function (p) {
            return Object.assign({}, p);
        });

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: 'bi-person-vcard',
            title: 'Rolle bearbeiten',
            sub: 'Änderungen schreiben in die Stammdaten (schulweit).',
            body:
                '<section class="schul-org-edit-panel">' +
                '<h4 class="schul-org-edit-panel-title">Stammdaten</h4>' +
                '<div class="form-grid schul-org-edit-grid">' +
                '<div class="field"><label for="schulOrgRoleName">Bezeichnung</label>' +
                '<input type="text" id="schulOrgRoleName" maxlength="80" autocomplete="off"></div>' +
                '<div class="field"><label for="schulOrgRoleCode">Kürzel</label>' +
                '<input type="text" id="schulOrgRoleCode" maxlength="24" autocomplete="off" spellcheck="false"></div>' +
                '<div class="field schul-org-edit-field-full"><label for="schulOrgRoleSeat">Besetzung (Hinweis)</label>' +
                '<select id="schulOrgRoleSeat">' +
                '<option value="single">Einzelrolle – typisch eine Person</option>' +
                '<option value="multi">Team-Rolle – mehrere Personen</option>' +
                '</select>' +
                '<p class="schul-org-edit-hint">Weitere Personen sind immer möglich; die Farbe richtet sich nach der Besetzung.</p></div>' +
                '</div></section>' +
                '<section class="schul-org-edit-panel">' +
                '<div class="schul-org-edit-panel-head">' +
                '<h4 class="schul-org-edit-panel-title">Personen</h4>' +
                '<p class="schul-org-edit-hint" id="schulOrgPeopleHint"></p></div>' +
                '<div class="schul-org-edit-table-wrap"><table class="schul-org-edit-table" aria-label="Personen der Rolle">' +
                '<thead><tr><th>Name</th><th>E-Mail</th><th></th></tr></thead>' +
                '<tbody id="schulOrgPeopleBody"></tbody></table></div>' +
                '<div class="schul-org-edit-people-actions">' +
                '<button type="button" class="btn btn-sm btn-success" data-schul-org-add-person><i class="bi bi-person-plus" aria-hidden="true"></i> Person</button>' +
                '<button type="button" class="btn btn-sm btn-primary" data-schul-org-add-m365><i class="bi bi-microsoft" aria-hidden="true"></i> Aus Microsoft 365</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>' +
                '</section>',
            foot: editFootHtml('data-schul-org-del-role', 'Rolle löschen', 'data-schul-org-save-role')
        });

        document.body.appendChild(overlay);
        activeOverlay = overlay;
        document.addEventListener('keydown', onEsc, true);

        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgRoleName'));
        const codeEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgRoleCode'));
        const seatEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgRoleSeat'));
        const tbody = overlay.querySelector('#schulOrgPeopleBody');
        const hintEl = overlay.querySelector('#schulOrgPeopleHint');
        const titleEl = overlay.querySelector('#schulOrgEditTitle');
        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };

        nameEl.value = role.name || '';
        codeEl.value = role.code || '';
        seatEl.value = role.seatMode === 'multi' ? 'multi' : 'single';

        function updateHint() {
            const n = people.filter(function (p) {
                return !isBlankPerson(p);
            }).length;
            if (titleEl) titleEl.textContent = nameEl.value.trim() || role.name || 'Rolle';
            if (!hintEl) return;
            if (n === 0) {
                hintEl.textContent = 'Noch niemand – Person hinzufügen oder aus Microsoft 365 übernehmen.';
            } else if (n === 1) {
                hintEl.textContent =
                    'Aktuell eine Person (Einzelbesetzung). Weitere jederzeit möglich → Team-Rolle.';
            } else {
                hintEl.textContent = 'Team-Rolle mit ' + n + ' Personen.';
            }
            if (n >= 2) seatEl.value = 'multi';
        }

        function renderPeople() {
            if (!tbody) return;
            tbody.replaceChildren();
            if (!people.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 3;
                td.className = 'muted';
                td.textContent = 'Keine Personen.';
                tr.appendChild(td);
                tbody.appendChild(tr);
                updateHint();
                return;
            }
            people.forEach(function (person, idx) {
                const tr = document.createElement('tr');
                const tdN = document.createElement('td');
                const inpN = document.createElement('input');
                inpN.type = 'text';
                inpN.value = person.name || '';
                inpN.setAttribute('aria-label', 'Name');
                inpN.addEventListener('input', function () {
                    people[idx].name = inpN.value;
                    updateHint();
                });
                tdN.appendChild(inpN);
                const tdE = document.createElement('td');
                const inpE = document.createElement('input');
                inpE.type = 'email';
                inpE.value = person.email || '';
                inpE.setAttribute('aria-label', 'E-Mail');
                inpE.addEventListener('input', function () {
                    people[idx].email = inpE.value;
                });
                tdE.appendChild(inpE);
                const tdA = document.createElement('td');
                tdA.className = 'schul-org-edit-action';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'btn btn-sm btn-danger';
                btnDel.title = 'Person entfernen';
                btnDel.innerHTML = '<i class="bi bi-x-lg" aria-hidden="true"></i>';
                btnDel.addEventListener('click', function () {
                    people.splice(idx, 1);
                    renderPeople();
                });
                tdA.appendChild(btnDel);
                tr.append(tdN, tdE, tdA);
                tbody.appendChild(tr);
            });
            updateHint();
        }

        function persistRole(closeAfter) {
            const nextName = normStr(nameEl.value);
            const nextCode = normStr(codeEl.value).toUpperCase() || roleCodeFromName(nextName);
            if (!nextName) {
                toast('Bitte eine Bezeichnung eingeben.', 'warning');
                return;
            }
            const cur = loadSettings();
            let roles = Array.isArray(cur.adminRoles) ? cur.adminRoles.slice() : [];
            let admin = Array.isArray(cur.admin) ? cur.admin.slice() : [];
            const oldName = role.name;
            const oldCode = role.code;

            // Rename people role labels if name changed
            if (nextName !== oldName && typeof window.ms365TenantSettingsRenameAdminRole === 'function') {
                const renamed = window.ms365TenantSettingsRenameAdminRole(roles, admin, oldName, nextName);
                roles = renamed.roles;
                admin = renamed.admin;
            }

            let found =
                roles.find(function (r) {
                    return (
                        normStr(r.name).toLowerCase() === nextName.toLowerCase() ||
                        normStr(r.code).toUpperCase() === normStr(oldCode).toUpperCase()
                    );
                }) || null;
            if (!found) {
                found = {
                    code: nextCode,
                    name: nextName,
                    seatMode: seatEl.value === 'multi' ? 'multi' : 'single',
                    tier: role.tier || 'verwaltung'
                };
                roles.push(found);
            } else {
                found.name = nextName;
                const clash = roles.some(function (r) {
                    return r !== found && normStr(r.code).toUpperCase() === nextCode;
                });
                if (!clash) found.code = nextCode;
                found.seatMode = seatEl.value === 'multi' ? 'multi' : 'single';
                if (role.m365ResourceType) found.m365ResourceType = role.m365ResourceType;
                if (role.m365ResourceId) found.m365ResourceId = role.m365ResourceId;
                if (role.m365ResourceEmail) found.m365ResourceEmail = role.m365ResourceEmail;
                if (role.m365ResourceLabel) found.m365ResourceLabel = role.m365ResourceLabel;
            }
            syncSeatMode(found, people);

            // Replace people for this role
            admin = admin.filter(function (p) {
                return !personMatchesRole(p, { name: oldName, code: oldCode }) && !personMatchesRole(p, found);
            });
            people.forEach(function (p) {
                admin.push({
                    role: found.name,
                    name: normStr(p.name),
                    email: normEmail(p.email),
                    defaultKey: found.name
                });
            });

            cur.adminRoles = roles;
            cur.admin = admin;
            if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
                cur.administration = window.ms365TenantSettingsAdminRolesAndAdminToGroups(roles, admin);
            }
            patchSchulOrganigrammRefLabel(cur, 'role', oldCode || refKey, found.name, found.code);
            saveSettings(cur);
            role = Object.assign({}, found);
            toast('Rolle gespeichert.', 'success');
            notify();
            if (closeAfter) closeOverlay();
            else {
                people = peopleForAdminRole(loadSettings(), found).map(function (p) {
                    return Object.assign({}, p);
                });
                nameEl.value = found.name || '';
                codeEl.value = found.code || '';
                seatEl.value = found.seatMode === 'multi' ? 'multi' : 'single';
                renderPeople();
            }
        }

        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });
        overlay.querySelector('[data-schul-org-save-role]').addEventListener('click', function () {
            persistRole(true);
        });
        overlay.querySelector('[data-schul-org-add-person]').addEventListener('click', function () {
            people.push({ role: role.name, name: '', email: '', defaultKey: role.name });
            syncSeatMode(role, people);
            if (people.filter(function (p) {
                return !isBlankPerson(p);
            }).length +
                people.filter(isBlankPerson).length >=
                2) {
                seatEl.value = 'multi';
            }
            renderPeople();
            const last = tbody && tbody.querySelector('tr:last-child input');
            if (last) last.focus();
        });
        overlay.querySelector('[data-schul-org-add-m365]').addEventListener('click', function () {
            const existing = new Set(
                people
                    .map(function (p) {
                        return normEmail(p.email);
                    })
                    .filter(function (e) {
                        return e.indexOf('@') !== -1;
                    })
            );
            void toggleInlineEntraUserPicker(inlineEupHost, inlineEupState, {
                multi: true,
                hint:
                    'Personen für „' +
                    (nameEl.value || role.name) +
                    '“ suchen, anhaken und übernehmen.',
                excludeEmails: existing,
                onApply: function (picked) {
                    if (!picked || !picked.length) return;
                    let added = 0;
                    picked.forEach(function (u) {
                        const mail = normEmail(u.mail || u.userPrincipalName || u.email);
                        const name = normStr(u.displayName || u.name);
                        if (!mail || mail.indexOf('@') === -1 || existing.has(mail)) return;
                        const draft = people.find(isBlankPerson);
                        if (draft) {
                            draft.name = name;
                            draft.email = mail;
                        } else {
                            people.push({
                                role: role.name,
                                name: name,
                                email: mail,
                                defaultKey: role.name
                            });
                        }
                        existing.add(mail);
                        added += 1;
                    });
                    if (!added) {
                        toast('Keine neue Person übernommen.', 'info');
                        return;
                    }
                    seatEl.value =
                        people.filter(function (p) {
                            return !isBlankPerson(p);
                        }).length >= 2
                            ? 'multi'
                            : seatEl.value;
                    renderPeople();
                    toast(added + ' Person(en) aus Microsoft 365 übernommen.', 'success');
                }
            }).catch(function (e) {
                toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
            });
        });
        overlay.querySelector('[data-schul-org-del-role]').addEventListener('click', function () {
            void confirmDlg('Rolle „' + (role.name || role.code) + '“ wirklich löschen? Personen müssen vorher entfernt sein.', {
                title: 'Rolle löschen',
                okText: 'Löschen',
                cancelText: 'Abbrechen'
            }).then(function (ok) {
                if (!ok) return;
                const filled = people.filter(function (p) {
                    return !isBlankPerson(p);
                });
                if (filled.length) {
                    toast('Zuerst alle Personen entfernen.', 'warning');
                    return;
                }
                const cur = loadSettings();
                const code = normStr(role.code).toUpperCase();
                cur.adminRoles = (cur.adminRoles || []).filter(function (r) {
                    return normStr(r.code).toUpperCase() !== code;
                });
                cur.admin = (cur.admin || []).filter(function (p) {
                    return !personMatchesRole(p, role);
                });
                if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
                    cur.administration = window.ms365TenantSettingsAdminRolesAndAdminToGroups(
                        cur.adminRoles,
                        cur.admin
                    );
                }
                if (cur.schulOrganigramm && Array.isArray(cur.schulOrganigramm.nodes)) {
                    cur.schulOrganigramm = {
                        version: 1,
                        nodes: cur.schulOrganigramm.nodes.filter(function (n) {
                            return !(
                                n.type === 'role' &&
                                normStr(n.refKey).toUpperCase() === code
                            );
                        })
                    };
                }
                saveSettings(cur);
                toast('Rolle gelöscht.', 'success');
                closeOverlay();
                notify();
            });
        });

        renderPeople();
        nameEl.focus();
    }

    /**
     * Neue Rolle unter einem Section-Knoten anlegen.
     * @param {string} parentSectionId
     */
    async function createRoleUnderSection(parentSectionId) {
        let name = '';
        if (typeof window.ms365AppDialogPrompt === 'function') {
            name = await window.ms365AppDialogPrompt('Bezeichnung der neuen Rolle', '', {
                title: 'Rolle anlegen',
                inputLabel: 'Rolle'
            });
        } else {
            name = window.prompt('Bezeichnung der neuen Rolle', '') || '';
        }
        const label = normStr(name);
        if (!label) return null;
        const cur = loadSettings();
        const exists = (cur.adminRoles || []).some(function (r) {
            return normStr(r.name).toLowerCase() === label.toLowerCase();
        });
        if (exists) {
            toast('Diese Rolle gibt es bereits.', 'warning');
            return null;
        }
        const used = new Set(
            (cur.adminRoles || []).map(function (r) {
                return normStr(r.code).toUpperCase();
            })
        );
        let code = roleCodeFromName(label);
        let i = 2;
        while (used.has(code)) {
            code = (roleCodeFromName(label) + String(i)).slice(0, 24);
            i += 1;
        }
        const layout =
            cur.schulOrganigramm && Array.isArray(cur.schulOrganigramm.nodes)
                ? normalizeSchulOrganigramm(cur.schulOrganigramm)
                : { version: 1, nodes: [] };

        /** @type {string[]} */
        const parentCandidates = [];
        const explicit = normStr(parentSectionId);
        if (explicit) parentCandidates.push(explicit);
        const remembered = readLastRoleParentId();
        if (remembered && parentCandidates.indexOf(remembered) < 0) {
            parentCandidates.push(remembered);
        }
        if (parentCandidates.indexOf('sec_verwaltung') < 0) {
            parentCandidates.push('sec_verwaltung');
        }
        if (parentCandidates.indexOf('sec_schulleitung') < 0) {
            parentCandidates.push('sec_schulleitung');
        }

        let resolvedParent = '';
        for (let p = 0; p < parentCandidates.length; p++) {
            const cand = parentCandidates[p];
            if (canPlaceSchulOrgNodeInSection({ type: NODE.ROLE }, cand, layout)) {
                resolvedParent = cand;
                break;
            }
        }
        if (!resolvedParent) {
            toast('In diesem Bereich können keine Funktionen angelegt werden.', 'warning');
            return null;
        }

        const tier =
            /direktion|direktor|schulleit/i.test(label) || resolvedParent === 'sec_schulleitung'
                ? 'schulleitung'
                : 'verwaltung';
        const role = { code: code, name: label, tier: tier, seatMode: 'single' };
        cur.adminRoles = (cur.adminRoles || []).concat([role]);
        if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
            cur.administration = window.ms365TenantSettingsAdminRolesAndAdminToGroups(
                cur.adminRoles,
                cur.admin || []
            );
        }
        const maxOrder = layout.nodes.reduce(function (m, n) {
            return Math.max(m, n.order || 0);
        }, 0);
        layout.nodes.push({
            id: 'role_' + code,
            type: 'role',
            label: label,
            parentId: resolvedParent,
            order: maxOrder + 1,
            tone: tier === 'schulleitung' || resolvedParent === 'sec_schulleitung' ? 'lead' : 'staff',
            refKind: 'adminRole',
            refKey: code
        });
        cur.schulOrganigramm = layout;
        rememberSchulOrgRoleParent(resolvedParent);
        saveSettings(cur);
        toast(
            'Funktion „' +
                label +
                '“ in „' +
                ((layout.nodes.find(function (n) {
                    return n.id === resolvedParent;
                }) || {}).label || resolvedParent) +
                '“ angelegt.',
            'success'
        );
        notify();
        openRoleEditor(code);
        return { code: code, parentId: resolvedParent };
    }

    /**
     * @param {string} refKey – ARGE-Kürzel
     */
    function openArgeEditor(refKey) {
        closeOverlay();
        const settings = loadSettings();
        let arge = findArge(settings, refKey);
        if (!arge) {
            toast('ARGE nicht in den Stammdaten gefunden.', 'warning');
            return;
        }
        arge = Object.assign({}, arge, {
            subjects: Array.isArray(arge.subjects) ? arge.subjects.slice() : []
        });

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: 'bi-journal-bookmark',
            title: 'ARGE bearbeiten',
            sub: 'Leitung aus Lehrerliste wählen oder aus Microsoft 365 übernehmen.',
            body:
                '<section class="schul-org-edit-panel">' +
                '<h4 class="schul-org-edit-panel-title">Stammdaten</h4>' +
                '<div class="form-grid schul-org-edit-grid">' +
                '<div class="field"><label for="schulOrgArgeName">Bezeichnung</label>' +
                '<input type="text" id="schulOrgArgeName" maxlength="80" autocomplete="off"></div>' +
                '<div class="field"><label for="schulOrgArgeCode">Kürzel</label>' +
                '<input type="text" id="schulOrgArgeCode" maxlength="24" autocomplete="off" spellcheck="false"></div>' +
                '<div class="field schul-org-edit-field-full"><label for="schulOrgArgeSubjects">Fächer (Komma-getrennt)</label>' +
                '<input type="text" id="schulOrgArgeSubjects" maxlength="200" autocomplete="off" placeholder="z. B. MATHE, PHYSIK"></div>' +
                '</div></section>' +
                '<section class="schul-org-edit-panel">' +
                '<div class="schul-org-edit-panel-head">' +
                '<h4 class="schul-org-edit-panel-title">Leitung (Lehrer-Bezug)</h4>' +
                '<p class="schul-org-edit-hint" id="schulOrgArgeHeadHint"></p></div>' +
                '<div class="form-grid schul-org-edit-grid">' +
                '<div class="field schul-org-edit-field-full"><label for="schulOrgArgeTeacher">Aus Lehrerliste</label>' +
                '<select id="schulOrgArgeTeacher"><option value="">– manuell / M365 –</option></select></div>' +
                '<div class="field"><label for="schulOrgArgeHeadName">Name Leitung</label>' +
                '<input type="text" id="schulOrgArgeHeadName" maxlength="120" autocomplete="name"></div>' +
                '<div class="field"><label for="schulOrgArgeHeadEmail">E-Mail Leitung</label>' +
                '<input type="email" id="schulOrgArgeHeadEmail" maxlength="160" autocomplete="email"></div>' +
                '</div>' +
                '<div class="schul-org-edit-people-actions">' +
                '<button type="button" class="btn btn-sm btn-primary" data-schul-org-arge-m365><i class="bi bi-microsoft" aria-hidden="true"></i> Leitung aus Microsoft 365</button>' +
                '<button type="button" class="btn btn-sm schul-org-edit-btn-ghost" data-schul-org-arge-clear-head><i class="bi bi-x-lg" aria-hidden="true"></i> Leitung leeren</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>' +
                '</section>',
            foot: editFootHtml('data-schul-org-del-arge', 'ARGE löschen', 'data-schul-org-save-arge')
        });

        document.body.appendChild(overlay);
        activeOverlay = overlay;
        document.addEventListener('keydown', onEsc, true);

        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgArgeName'));
        const codeEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgArgeCode'));
        const subjEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgArgeSubjects'));
        const teacherEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgArgeTeacher'));
        const headNameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgArgeHeadName'));
        const headEmailEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgArgeHeadEmail'));
        const hintEl = overlay.querySelector('#schulOrgArgeHeadHint');
        const titleEl = overlay.querySelector('#schulOrgEditTitle');
        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };

        nameEl.value = arge.name || '';
        codeEl.value = arge.code || '';
        subjEl.value = (arge.subjects || []).join(', ');
        headNameEl.value = arge.headName || '';
        headEmailEl.value = arge.headEmail || '';

        const teachers = Array.isArray(settings.teachers) ? settings.teachers.slice() : [];
        teachers
            .filter(function (t) {
                return normEmail(t && t.email).indexOf('@') !== -1;
            })
            .sort(function (a, b) {
                return normStr(a.name).localeCompare(normStr(b.name), 'de');
            })
            .forEach(function (t) {
                const opt = document.createElement('option');
                opt.value = normEmail(t.email);
                opt.textContent =
                    (normStr(t.name) || normEmail(t.email)) +
                    (t.code ? ' (' + normStr(t.code) + ')' : '') +
                    ' – ' +
                    normEmail(t.email);
                teacherEl.appendChild(opt);
            });

        function syncTeacherSelect() {
            const em = normEmail(headEmailEl.value);
            teacherEl.value = em && teachers.some(function (t) {
                return normEmail(t.email) === em;
            })
                ? em
                : '';
        }

        function updateHeadHint() {
            if (titleEl) titleEl.textContent = nameEl.value.trim() || arge.name || 'ARGE';
            if (!hintEl) return;
            const em = normEmail(headEmailEl.value);
            if (!em) {
                hintEl.textContent = 'Keine Leitung gesetzt – ARGE ist „offen“.';
                return;
            }
            const t = findTeacherByEmail({ teachers: teachers }, em);
            if (t) {
                hintEl.textContent =
                    'Verknüpft mit Lehrerliste: ' +
                    (t.name || t.code || em) +
                    (t.code ? ' · ' + t.code : '');
            } else {
                hintEl.textContent =
                    'E-Mail gesetzt, aber kein Treffer in der Lehrerliste (trotzdem gültig).';
            }
        }

        teacherEl.addEventListener('change', function () {
            const em = normEmail(teacherEl.value);
            if (!em) return;
            const t = findTeacherByEmail({ teachers: teachers }, em);
            headEmailEl.value = em;
            if (t && t.name) headNameEl.value = t.name;
            updateHeadHint();
        });
        headEmailEl.addEventListener('input', function () {
            syncTeacherSelect();
            updateHeadHint();
        });
        headNameEl.addEventListener('input', updateHeadHint);
        nameEl.addEventListener('input', updateHeadHint);

        syncTeacherSelect();
        updateHeadHint();

        function persistArge(closeAfter) {
            const nextName = normStr(nameEl.value);
            const nextCode = normStr(codeEl.value).toUpperCase() || roleCodeFromName(nextName);
            if (!nextName) {
                toast('Bitte eine Bezeichnung eingeben.', 'warning');
                return;
            }
            if (!nextCode) {
                toast('Bitte ein Kürzel eingeben.', 'warning');
                return;
            }
            const subjects = String(subjEl.value || '')
                .split(/[,;\s]+/)
                .map(function (x) {
                    return normStr(x).toUpperCase();
                })
                .filter(Boolean);
            const cur = loadSettings();
            const arges = Array.isArray(cur.arges) ? cur.arges.slice() : [];
            const oldCode = normStr(arge.code).toUpperCase();
            let found = arges.find(function (a) {
                return normStr(a.code).toUpperCase() === oldCode;
            });
            const clash = arges.some(function (a) {
                return (
                    a !== found &&
                    normStr(a.code).toUpperCase() === nextCode
                );
            });
            if (clash) {
                toast('Kürzel bereits vergeben.', 'warning');
                return;
            }
            if (!found) {
                found = { code: nextCode, name: nextName, subjects: subjects };
                arges.push(found);
            } else {
                found.code = nextCode;
                found.name = nextName;
                found.subjects = subjects;
            }
            found.headName = normStr(headNameEl.value);
            found.headEmail = normEmail(headEmailEl.value);
            cur.arges = arges;
            patchSchulOrganigrammRefLabel(cur, 'arge', oldCode || refKey, found.name, found.code);
            saveSettings(cur);
            arge = Object.assign({}, found);
            toast('ARGE gespeichert.', 'success');
            notify();
            if (closeAfter) closeOverlay();
            else {
                nameEl.value = found.name || '';
                codeEl.value = found.code || '';
                subjEl.value = (found.subjects || []).join(', ');
                headNameEl.value = found.headName || '';
                headEmailEl.value = found.headEmail || '';
                syncTeacherSelect();
                updateHeadHint();
            }
        }

        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });
        overlay.querySelector('[data-schul-org-save-arge]').addEventListener('click', function () {
            persistArge(true);
        });
        overlay.querySelector('[data-schul-org-arge-clear-head]').addEventListener('click', function () {
            headNameEl.value = '';
            headEmailEl.value = '';
            teacherEl.value = '';
            updateHeadHint();
        });
        overlay.querySelector('[data-schul-org-arge-m365]').addEventListener('click', function () {
            void toggleInlineEntraUserPicker(inlineEupHost, inlineEupState, {
                hint: 'ARGE-Leitung in Microsoft 365 suchen und antippen.',
                onPick: function (picked) {
                    const mail = normEmail(picked.mail || picked.userPrincipalName || picked.email);
                    const name = normStr(picked.displayName || picked.name);
                    if (!mail || mail.indexOf('@') === -1) {
                        toast('Keine gültige E-Mail in der Auswahl.', 'warning');
                        return;
                    }
                    headEmailEl.value = mail;
                    if (name) headNameEl.value = name;
                    syncTeacherSelect();
                    updateHeadHint();
                    toast('Leitung aus Microsoft 365 übernommen.', 'success');
                }
            }).catch(function (e) {
                toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
            });
        });
        overlay.querySelector('[data-schul-org-del-arge]').addEventListener('click', function () {
            void confirmDlg('ARGE „' + (arge.name || arge.code) + '“ wirklich löschen?', {
                title: 'ARGE löschen',
                okText: 'Löschen',
                cancelText: 'Abbrechen'
            }).then(function (ok) {
                if (!ok) return;
                const cur = loadSettings();
                const code = normStr(arge.code).toUpperCase();
                cur.arges = (cur.arges || []).filter(function (a) {
                    return normStr(a.code).toUpperCase() !== code;
                });
                if (cur.schulOrganigramm && Array.isArray(cur.schulOrganigramm.nodes)) {
                    cur.schulOrganigramm = {
                        version: 1,
                        nodes: cur.schulOrganigramm.nodes.filter(function (n) {
                            if (!(isArgeNode(n) && normStr(n.refKey).toUpperCase() === code)) return true;
                            return false;
                        })
                    };
                }
                saveSettings(cur);
                toast('ARGE gelöscht.', 'success');
                closeOverlay();
                notify();
            });
        });

        nameEl.focus();
    }

    /**
     * @param {string} [parentSectionId]
     */
    async function createArgeUnderSection(parentSectionId) {
        let name = '';
        if (typeof window.ms365AppDialogPrompt === 'function') {
            name = await window.ms365AppDialogPrompt('Bezeichnung der neuen ARGE / Fachschaft', '', {
                title: 'ARGE anlegen',
                inputLabel: 'Bezeichnung'
            });
        } else {
            name = window.prompt('Bezeichnung der neuen ARGE', '') || '';
        }
        const label = normStr(name);
        if (!label) return;
        const cur = loadSettings();
        const exists = (cur.arges || []).some(function (a) {
            return normStr(a.name).toLowerCase() === label.toLowerCase();
        });
        if (exists) {
            toast('Diese ARGE gibt es bereits.', 'warning');
            return;
        }
        const used = new Set(
            (cur.arges || []).map(function (a) {
                return normStr(a.code).toUpperCase();
            })
        );
        let code = roleCodeFromName(label);
        let i = 2;
        while (used.has(code)) {
            code = (roleCodeFromName(label) + String(i)).slice(0, 24);
            i += 1;
        }
        const arge = {
            code: code,
            name: label,
            subjects: [],
            headName: '',
            headEmail: ''
        };
        cur.arges = (cur.arges || []).concat([arge]);
        const layout =
            cur.schulOrganigramm && Array.isArray(cur.schulOrganigramm.nodes)
                ? cur.schulOrganigramm
                : { version: 1, nodes: [] };
        const maxOrder = layout.nodes.reduce(function (m, n) {
            return Math.max(m, n.order || 0);
        }, 0);
        layout.nodes.push({
            id: 'arge_' + code,
            type: 'area',
            label: label,
            parentId: parentSectionId || 'sec_arge',
            order: maxOrder + 1,
            tone: 'arge',
            refKind: 'arge',
            refKey: code,
            meta: { nestPeople: true, audience: 'arge' }
        });
        cur.schulOrganigramm = layout;
        saveSettings(cur);
        toast('ARGE angelegt.', 'success');
        notify();
        openArgeEditor(code);
    }

    /**
     * @param {string} refKey – Klassen-Code
     */
    function openClassEditor(refKey) {
        closeOverlay();
        const settings = loadSettings();
        let cls = findClass(settings, refKey);
        if (!cls) {
            toast('Klasse nicht in den Stammdaten gefunden.', 'warning');
            return;
        }
        cls = Object.assign({}, cls);

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: 'bi-mortarboard',
            title: 'Klasse bearbeiten',
            sub: 'Klassenvorstand aus Lehrerliste oder Microsoft 365 zuordnen.',
            body:
                '<section class="schul-org-edit-panel">' +
                '<h4 class="schul-org-edit-panel-title">Klasse</h4>' +
                '<div class="form-grid schul-org-edit-grid">' +
                '<div class="field schul-org-edit-field-full"><label for="schulOrgClassName">Bezeichnung</label>' +
                '<input type="text" id="schulOrgClassName" maxlength="80" autocomplete="off"></div>' +
                '<div class="field"><label for="schulOrgClassCode">Kürzel</label>' +
                '<input type="text" id="schulOrgClassCode" maxlength="24" autocomplete="off" spellcheck="false"></div>' +
                '<div class="field"><label for="schulOrgClassYear">Abschlussjahr</label>' +
                '<input type="text" id="schulOrgClassYear" maxlength="4" inputmode="numeric" pattern="[0-9]{4}" placeholder="z. B. 2030" autocomplete="off"></div>' +
                '</div></section>' +
                '<section class="schul-org-edit-panel">' +
                '<div class="schul-org-edit-panel-head">' +
                '<h4 class="schul-org-edit-panel-title">Klassenvorstand</h4>' +
                '<p class="schul-org-edit-hint" id="schulOrgClassHeadHint"></p></div>' +
                '<div class="form-grid schul-org-edit-grid">' +
                '<div class="field schul-org-edit-field-full"><label for="schulOrgClassTeacher">Aus Lehrerliste</label>' +
                '<select id="schulOrgClassTeacher"><option value="">– manuell / M365 –</option></select></div>' +
                '<div class="field"><label for="schulOrgClassHeadName">Name KV</label>' +
                '<input type="text" id="schulOrgClassHeadName" maxlength="120" autocomplete="name"></div>' +
                '<div class="field"><label for="schulOrgClassHeadEmail">E-Mail KV</label>' +
                '<input type="email" id="schulOrgClassHeadEmail" maxlength="160" autocomplete="email"></div>' +
                '</div>' +
                '<div class="schul-org-edit-people-actions">' +
                '<button type="button" class="btn btn-sm btn-primary" data-schul-org-class-m365><i class="bi bi-microsoft" aria-hidden="true"></i> KV aus Microsoft 365</button>' +
                '<button type="button" class="btn btn-sm schul-org-edit-btn-ghost" data-schul-org-class-clear-head><i class="bi bi-x-lg" aria-hidden="true"></i> KV leeren</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>' +
                '</section>',
            foot: editFootHtml('data-schul-org-del-class', 'Klasse löschen', 'data-schul-org-save-class')
        });

        document.body.appendChild(overlay);
        activeOverlay = overlay;
        document.addEventListener('keydown', onEsc, true);

        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgClassName'));
        const codeEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgClassCode'));
        const yearEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgClassYear'));
        const teacherEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgClassTeacher'));
        const headNameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgClassHeadName'));
        const headEmailEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgClassHeadEmail'));
        const hintEl = overlay.querySelector('#schulOrgClassHeadHint');
        const titleEl = overlay.querySelector('#schulOrgEditTitle');
        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };

        nameEl.value = cls.name || '';
        codeEl.value = cls.code || '';
        yearEl.value = cls.year || '';
        headNameEl.value = cls.headName || '';
        headEmailEl.value = cls.headEmail || '';

        const teachers = Array.isArray(settings.teachers) ? settings.teachers.slice() : [];
        fillTeacherSelect(teacherEl, teachers);
        const headWire = wireTeacherHeadFields({
            teacherEl: teacherEl,
            headNameEl: headNameEl,
            headEmailEl: headEmailEl,
            hintEl: hintEl,
            teachers: teachers,
            vacantText: 'Kein Klassenvorstand – KV ist „offen“.'
        });

        nameEl.addEventListener('input', function () {
            if (titleEl) titleEl.textContent = nameEl.value.trim() || cls.name || 'Klasse';
        });
        if (titleEl) titleEl.textContent = nameEl.value.trim() || cls.name || 'Klasse';

        function persistClass(closeAfter) {
            const nextName = normStr(nameEl.value);
            const nextCode = normStr(codeEl.value).toUpperCase() || roleCodeFromName(nextName);
            const yearRaw = normStr(yearEl.value);
            const nextYear = /^\d{4}$/.test(yearRaw) ? yearRaw : '';
            if (!nextName && !nextCode) {
                toast('Bitte Bezeichnung oder Kürzel eingeben.', 'warning');
                return;
            }
            if (!nextCode) {
                toast('Bitte ein Kürzel eingeben.', 'warning');
                return;
            }
            if (yearRaw && !nextYear) {
                toast('Abschlussjahr muss vierstellig sein (z. B. 2030).', 'warning');
                return;
            }
            const cur = loadSettings();
            const classes = Array.isArray(cur.classes) ? cur.classes.slice() : [];
            const oldCode = normStr(cls.code).toUpperCase() || normStr(refKey).toUpperCase();
            let found = classes.find(function (c) {
                return normStr(c.code).toUpperCase() === oldCode;
            });
            const clash = classes.some(function (c) {
                return c !== found && normStr(c.code).toUpperCase() === nextCode;
            });
            if (clash) {
                toast('Klassenkürzel bereits vergeben.', 'warning');
                return;
            }
            if (!found) {
                found = {
                    code: nextCode,
                    name: nextName || nextCode,
                    year: nextYear,
                    headName: '',
                    headEmail: '',
                    stableMailNickname: ''
                };
                classes.push(found);
            } else {
                found.code = nextCode;
                found.name = nextName || nextCode;
                found.year = nextYear;
            }
            found.headName = normStr(headNameEl.value);
            found.headEmail = normEmail(headEmailEl.value);
            cur.classes = classes;
            patchSchulOrganigrammRefLabel(
                cur,
                'class',
                oldCode || refKey,
                found.name || found.code,
                found.code
            );
            saveSettings(cur);
            cls = Object.assign({}, found);
            toast('Klasse gespeichert.', 'success');
            notify();
            if (closeAfter) closeOverlay();
            else {
                nameEl.value = found.name || '';
                codeEl.value = found.code || '';
                yearEl.value = found.year || '';
                headNameEl.value = found.headName || '';
                headEmailEl.value = found.headEmail || '';
                headWire.syncTeacherSelect();
                headWire.updateHeadHint();
                if (titleEl) titleEl.textContent = found.name || found.code || 'Klasse';
            }
        }

        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });
        overlay.querySelector('[data-schul-org-save-class]').addEventListener('click', function () {
            persistClass(true);
        });
        overlay.querySelector('[data-schul-org-class-clear-head]').addEventListener('click', function () {
            headNameEl.value = '';
            headEmailEl.value = '';
            teacherEl.value = '';
            headWire.updateHeadHint();
        });
        overlay.querySelector('[data-schul-org-class-m365]').addEventListener('click', function () {
            void toggleInlineEntraUserPicker(inlineEupHost, inlineEupState, {
                hint: 'Klassenvorstand in Microsoft 365 suchen und antippen.',
                onPick: function (picked) {
                    const mail = normEmail(picked.mail || picked.userPrincipalName || picked.email);
                    const name = normStr(picked.displayName || picked.name);
                    if (!mail || mail.indexOf('@') === -1) {
                        toast('Keine gültige E-Mail in der Auswahl.', 'warning');
                        return;
                    }
                    headEmailEl.value = mail;
                    if (name) headNameEl.value = name;
                    headWire.syncTeacherSelect();
                    headWire.updateHeadHint();
                    toast('Klassenvorstand aus Microsoft 365 übernommen.', 'success');
                }
            }).catch(function (e) {
                toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
            });
        });
        overlay.querySelector('[data-schul-org-del-class]').addEventListener('click', function () {
            void confirmDlg('Klasse „' + (cls.name || cls.code) + '“ wirklich löschen?', {
                title: 'Klasse löschen',
                okText: 'Löschen',
                cancelText: 'Abbrechen'
            }).then(function (ok) {
                if (!ok) return;
                const cur = loadSettings();
                const code = normStr(cls.code).toUpperCase();
                cur.classes = (cur.classes || []).filter(function (c) {
                    return normStr(c.code).toUpperCase() !== code;
                });
                removeSchulOrganigrammByRef(cur, 'class', code);
                saveSettings(cur);
                toast('Klasse gelöscht.', 'success');
                closeOverlay();
                notify();
            });
        });

        nameEl.focus();
    }

    /**
     * @param {string} [parentSectionId]
     */
    async function createClassUnderSection(parentSectionId) {
        let codeIn = '';
        if (typeof window.ms365AppDialogPrompt === 'function') {
            codeIn = await window.ms365AppDialogPrompt('Kürzel der neuen Klasse (z. B. 1A)', '', {
                title: 'Klasse anlegen',
                inputLabel: 'Kürzel'
            });
        } else {
            codeIn = window.prompt('Kürzel der neuen Klasse (z. B. 1A)', '') || '';
        }
        const code = normStr(codeIn).toUpperCase();
        if (!code) return;
        const cur = loadSettings();
        if ((cur.classes || []).some(function (c) {
            return normStr(c.code).toUpperCase() === code;
        })) {
            toast('Diese Klasse gibt es bereits.', 'warning');
            return;
        }
        const cls = {
            code: code,
            name: code,
            year: '',
            headName: '',
            headEmail: '',
            stableMailNickname: ''
        };
        cur.classes = (cur.classes || []).concat([cls]);
        const layout =
            cur.schulOrganigramm && Array.isArray(cur.schulOrganigramm.nodes)
                ? cur.schulOrganigramm
                : { version: 1, nodes: [] };
        const maxOrder = layout.nodes.reduce(function (m, n) {
            return Math.max(m, n.order || 0);
        }, 0);
        layout.nodes.push({
            id: 'class_' + code,
            type: 'class',
            label: code,
            parentId: parentSectionId || jahrgangAreaId(schulstufeFromClassCode(code)),
            order: maxOrder + 1,
            tone: 'class',
            refKind: 'class',
            refKey: code
        });
        cur.schulOrganigramm = layout;
        saveSettings(cur);
        toast('Klasse angelegt.', 'success');
        notify();
        openClassEditor(code);
    }

    /**
     * @param {'council'|'sga'} kind
     * @param {string} refKey
     */
    function openPersonListEditor(kind, refKey) {
        closeOverlay();
        const isCouncil = kind === 'council';
        const settings = loadSettings();
        let row = isCouncil ? findStudentCouncil(settings, refKey) : findSga(settings, refKey);
        if (!row) {
            toast((isCouncil ? 'Schülervertretung' : 'SGA-Eintrag') + ' nicht gefunden.', 'warning');
            return;
        }
        row = Object.assign({}, row);
        const listKey = isCouncil ? 'studentCouncil' : 'sga';
        const nodeType = isCouncil ? 'council' : 'sga';
        const titleDefault = isCouncil ? 'Schülervertretung' : 'SGA';

        const classOptions = (Array.isArray(settings.classes) ? settings.classes : [])
            .map(function (c) {
                return normStr(c.code).toUpperCase() || normStr(c.name);
            })
            .filter(Boolean)
            .sort(function (a, b) {
                return a.localeCompare(b, 'de');
            });

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: isCouncil ? 'bi-people' : 'bi-building',
            title: titleDefault + ' bearbeiten',
            sub: isCouncil
                ? 'Klassenvertretung: Name, E-Mail und Klasse.'
                : 'Schulgemeinschaftsausschuss: Name, E-Mail und Bereich.',
            body:
                '<section class="schul-org-edit-panel">' +
                '<h4 class="schul-org-edit-panel-title">Person</h4>' +
                '<div class="form-grid schul-org-edit-grid">' +
                '<div class="field"><label for="schulOrgPersName">Name</label>' +
                '<input type="text" id="schulOrgPersName" maxlength="120" autocomplete="name"></div>' +
                '<div class="field"><label for="schulOrgPersEmail">E-Mail</label>' +
                '<input type="email" id="schulOrgPersEmail" maxlength="160" autocomplete="email"></div>' +
                (isCouncil
                    ? '<div class="field"><label for="schulOrgPersKlasse">Klasse</label>' +
                      '<select id="schulOrgPersKlasse"><option value="">– wählen / frei –</option></select></div>' +
                      '<div class="field"><label for="schulOrgPersKlasseFree">Klasse (frei)</label>' +
                      '<input type="text" id="schulOrgPersKlasseFree" maxlength="24" autocomplete="off" spellcheck="false"></div>'
                    : '<div class="field schul-org-edit-field-full"><label for="schulOrgPersScope">Bereich</label>' +
                      '<select id="schulOrgPersScope">' +
                      '<option value="teacher">Lehrer</option>' +
                      '<option value="student">Schüler</option>' +
                      '<option value="external">Extern</option>' +
                      '</select></div>') +
                '</div>' +
                '<div class="schul-org-edit-people-actions">' +
                '<button type="button" class="btn btn-sm btn-primary" data-schul-org-pers-m365><i class="bi bi-microsoft" aria-hidden="true"></i> Aus Microsoft 365</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>' +
                '</section>',
            foot: editFootHtml('data-schul-org-del-pers', 'Löschen', 'data-schul-org-save-pers')
        });

        document.body.appendChild(overlay);
        activeOverlay = overlay;
        document.addEventListener('keydown', onEsc, true);

        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgPersName'));
        const emailEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgPersEmail'));
        const titleEl = overlay.querySelector('#schulOrgEditTitle');
        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };
        nameEl.value = row.name || '';
        emailEl.value = row.email || '';

        /** @type {HTMLSelectElement|null} */
        let klasseEl = null;
        /** @type {HTMLInputElement|null} */
        let klasseFreeEl = null;
        /** @type {HTMLSelectElement|null} */
        let scopeEl = null;

        if (isCouncil) {
            klasseEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgPersKlasse'));
            klasseFreeEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgPersKlasseFree'));
            classOptions.forEach(function (c) {
                const opt = document.createElement('option');
                opt.value = c;
                opt.textContent = c;
                klasseEl.appendChild(opt);
            });
            const k = normStr(row.klasse).toUpperCase();
            if (k && classOptions.indexOf(k) !== -1) {
                klasseEl.value = k;
                klasseFreeEl.value = '';
            } else {
                klasseEl.value = '';
                klasseFreeEl.value = normStr(row.klasse);
            }
        } else {
            scopeEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgPersScope'));
            const sc = normalizeSgaScope(row.scope) || 'teacher';
            scopeEl.value = sc;
        }

        function updateTitle() {
            if (titleEl) titleEl.textContent = (nameEl.value.trim() || row.name || titleDefault) + ' bearbeiten';
        }
        nameEl.addEventListener('input', updateTitle);
        updateTitle();

        let oldRef = normEmail(row.email) || normStr(row.name) || refKey;

        function persistPerson(closeAfter) {
            const nextName = normStr(nameEl.value);
            const nextEmail = normEmail(emailEl.value);
            if (!nextName && !nextEmail) {
                toast('Bitte Name oder E-Mail eingeben.', 'warning');
                return;
            }
            if (nextEmail && nextEmail.indexOf('@') === -1) {
                toast('E-Mail ungültig.', 'warning');
                return;
            }
            const cur = loadSettings();
            const list = Array.isArray(cur[listKey]) ? cur[listKey].slice() : [];
            let found =
                (normEmail(row.email)
                    ? list.find(function (s) {
                          return normEmail(s.email) === normEmail(row.email);
                      })
                    : null) ||
                list.find(function (s) {
                    return (
                        normStr(s.name).toLowerCase() === normStr(row.name).toLowerCase() &&
                        (!isCouncil || normStr(s.klasse).toLowerCase() === normStr(row.klasse).toLowerCase()) &&
                        (isCouncil || normalizeSgaScope(s.scope) === normalizeSgaScope(row.scope))
                    );
                });
            if (nextEmail) {
                const clash = list.some(function (s) {
                    return s !== found && normEmail(s.email) === nextEmail;
                });
                if (clash) {
                    toast('Diese E-Mail ist bereits eingetragen.', 'warning');
                    return;
                }
            }
            if (!found) {
                found = isCouncil
                    ? { klasse: '', name: nextName, email: nextEmail }
                    : { scope: 'teacher', name: nextName, email: nextEmail };
                list.push(found);
            }
            found.name = nextName;
            found.email = nextEmail;
            if (isCouncil) {
                const fromSelect = klasseEl ? normStr(klasseEl.value).toUpperCase() : '';
                const fromFree = klasseFreeEl ? normStr(klasseFreeEl.value).toUpperCase() : '';
                found.klasse = fromSelect || fromFree;
            } else {
                found.scope = scopeEl ? normalizeSgaScope(scopeEl.value) || 'teacher' : 'teacher';
            }
            cur[listKey] = list;
            const nextRef = nextEmail || nextName;
            patchSchulOrganigrammRefLabel(cur, nodeType, oldRef, nextName || nextEmail, nextRef);
            saveSettings(cur);
            row = Object.assign({}, found);
            oldRef = nextRef;
            toast(titleDefault + ' gespeichert.', 'success');
            notify();
            if (closeAfter) closeOverlay();
            else {
                nameEl.value = found.name || '';
                emailEl.value = found.email || '';
                updateTitle();
            }
        }

        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });
        overlay.querySelector('[data-schul-org-save-pers]').addEventListener('click', function () {
            persistPerson(true);
        });
        overlay.querySelector('[data-schul-org-pers-m365]').addEventListener('click', function () {
            void toggleInlineEntraUserPicker(inlineEupHost, inlineEupState, {
                hint: titleDefault + ' in Microsoft 365 suchen und antippen.',
                onPick: function (picked) {
                    const mail = normEmail(picked.mail || picked.userPrincipalName || picked.email);
                    const name = normStr(picked.displayName || picked.name);
                    if (mail && mail.indexOf('@') !== -1) emailEl.value = mail;
                    if (name) nameEl.value = name;
                    updateTitle();
                    toast('Aus Microsoft 365 übernommen.', 'success');
                }
            }).catch(function (e) {
                toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
            });
        });
        overlay.querySelector('[data-schul-org-del-pers]').addEventListener('click', function () {
            void confirmDlg('„' + (row.name || row.email || titleDefault) + '“ wirklich löschen?', {
                title: titleDefault + ' löschen',
                okText: 'Löschen',
                cancelText: 'Abbrechen'
            }).then(function (ok) {
                if (!ok) return;
                const cur = loadSettings();
                const mail = normEmail(row.email);
                const nameL = normStr(row.name).toLowerCase();
                cur[listKey] = (cur[listKey] || []).filter(function (s) {
                    if (mail && normEmail(s.email) === mail) return false;
                    if (
                        !mail &&
                        nameL &&
                        normStr(s.name).toLowerCase() === nameL &&
                        (!isCouncil || normStr(s.klasse).toLowerCase() === normStr(row.klasse).toLowerCase()) &&
                        (isCouncil || normalizeSgaScope(s.scope) === normalizeSgaScope(row.scope))
                    ) {
                        return false;
                    }
                    return true;
                });
                removeSchulOrganigrammByRef(cur, nodeType, oldRef);
                saveSettings(cur);
                toast(titleDefault + ' gelöscht.', 'success');
                closeOverlay();
                notify();
            });
        });

        nameEl.focus();
    }

    /**
     * Sammel-Editor: alle Personen der Schülervertretung bzw. SGA (nach Scope).
     * @param {'council'|'sga'} kind
     */
    function openBoardListEditor(kind) {
        closeOverlay();
        const isCouncil = kind === 'council';
        const listKey = isCouncil ? 'studentCouncil' : 'sga';
        const title = isCouncil ? 'Schülervertretung' : 'SGA';
        const settings = loadSettings();
        /** @type {object[]} */
        let rows = (Array.isArray(settings[listKey]) ? settings[listKey] : []).map(function (r) {
            return Object.assign({}, r);
        });

        const classOptions = (Array.isArray(settings.classes) ? settings.classes : [])
            .map(function (c) {
                return normStr(c.code).toUpperCase() || normStr(c.name);
            })
            .filter(Boolean)
            .sort(function (a, b) {
                return a.localeCompare(b, 'de');
            });

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: isCouncil ? 'bi-people' : 'bi-building',
            title: title,
            sub: isCouncil
                ? 'Alle Schülervertreter:innen – Name, E-Mail, Klasse.'
                : 'Schulgemeinschaftsausschuss nach Gruppen: Schüler, Lehrer, Eltern.',
            body:
                '<section class="schul-org-edit-panel">' +
                '<div class="schul-org-edit-people-actions" style="margin-bottom:10px">' +
                '<button type="button" class="btn btn-sm btn-success" data-schul-org-board-add>' +
                '<i class="bi bi-plus-lg" aria-hidden="true"></i> Person</button>' +
                '<button type="button" class="btn btn-sm btn-primary" data-schul-org-board-m365>' +
                '<i class="bi bi-microsoft" aria-hidden="true"></i> Aus Microsoft 365</button>' +
                (!isCouncil
                    ? '<select id="schulOrgBoardAddScope" class="schul-org-edit-inline-select" title="Gruppe">' +
                      '<option value="student">Schüler</option>' +
                      '<option value="teacher">Lehrer</option>' +
                      '<option value="external">Eltern</option>' +
                      '</select>'
                    : '') +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>' +
                '<div id="schulOrgBoardList" class="schul-org-edit-board-list"></div>' +
                '</section>',
            foot:
                '<button type="button" class="btn" data-schul-org-close>Schließen</button>' +
                '<button type="button" class="btn btn-primary" data-schul-org-save-board>Speichern</button>'
        });
        document.body.appendChild(overlay);
        activeOverlay = overlay;
        document.addEventListener('keydown', onEsc, true);

        const listHost = /** @type {HTMLElement} */ (overlay.querySelector('#schulOrgBoardList'));
        const scopeAddEl = /** @type {HTMLSelectElement|null} */ (
            overlay.querySelector('#schulOrgBoardAddScope')
        );
        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };

        function ensureBoardNodes(cur) {
            cur.schulOrganigramm = ensureSchulOrganigrammSgaStructure(cur.schulOrganigramm, cur);
        }

        function renderBoard() {
            listHost.replaceChildren();
            if (!rows.length) {
                const empty = document.createElement('p');
                empty.className = 'muted';
                empty.textContent = 'Noch keine Personen – „Person“ oder Microsoft 365 nutzen.';
                listHost.appendChild(empty);
                return;
            }
            /** @type {{ key: string, label: string, items: object[] }[]} */
            const groups = isCouncil
                ? [{ key: 'all', label: 'Schülervertretung (SGA)', items: rows }]
                : [
                      {
                          key: 'student',
                          label: 'Schülervertretung',
                          items: rows.filter(function (r) {
                              return (normalizeSgaScope(r.scope) || 'student') === 'student';
                          })
                      },
                      {
                          key: 'teacher',
                          label: 'Lehrer',
                          items: rows.filter(function (r) {
                              return normalizeSgaScope(r.scope) === 'teacher';
                          })
                      },
                      {
                          key: 'external',
                          label: 'Eltern',
                          items: rows.filter(function (r) {
                              return normalizeSgaScope(r.scope) === 'external';
                          })
                      }
                  ];
            groups.forEach(function (g) {
                const block = document.createElement('div');
                block.className = 'schul-org-edit-board-group';
                const h = document.createElement('h4');
                h.className = 'schul-org-edit-panel-title';
                h.textContent = g.label + ' (' + g.items.length + ')';
                block.appendChild(h);
                if (!g.items.length) {
                    const p = document.createElement('p');
                    p.className = 'muted';
                    p.textContent = 'leer';
                    block.appendChild(p);
                }
                g.items.forEach(function (row) {
                    const idx = rows.indexOf(row);
                    const seatHint =
                        isCouncil && SCHUL_ORG_SV_SEATS[idx]
                            ? SCHUL_ORG_SV_SEATS[idx].label
                            : g.key === 'student' && SCHUL_ORG_SV_SEATS[g.items.indexOf(row)]
                              ? SCHUL_ORG_SV_SEATS[g.items.indexOf(row)].label
                              : '';
                    const line = document.createElement('div');
                    line.className = 'schul-org-edit-board-row';
                    line.innerHTML =
                        (seatHint
                            ? '<span class="muted schul-org-edit-seat-label">' +
                              escapeHtml(seatHint) +
                              '</span>'
                            : '') +
                        '<input type="text" data-f="name" placeholder="Name" maxlength="120" />' +
                        '<input type="email" data-f="email" placeholder="E-Mail" maxlength="160" />' +
                        (isCouncil
                            ? '<input type="text" data-f="klasse" placeholder="Klasse" maxlength="24" list="schulOrgBoardKlassen" />'
                            : '<select data-f="scope">' +
                              '<option value="student">Schülervertretung</option>' +
                              '<option value="teacher">Lehrer</option>' +
                              '<option value="external">Eltern</option>' +
                              '</select>') +
                        '<button type="button" class="btn btn-sm" data-del title="Entfernen"><i class="bi bi-trash" aria-hidden="true"></i></button>';
                    const nameInp = /** @type {HTMLInputElement} */ (line.querySelector('[data-f="name"]'));
                    const emailInp = /** @type {HTMLInputElement} */ (line.querySelector('[data-f="email"]'));
                    nameInp.value = row.name || '';
                    emailInp.value = row.email || '';
                    nameInp.addEventListener('input', function () {
                        rows[idx].name = nameInp.value;
                    });
                    emailInp.addEventListener('input', function () {
                        rows[idx].email = emailInp.value;
                    });
                    if (isCouncil) {
                        const kInp = /** @type {HTMLInputElement} */ (line.querySelector('[data-f="klasse"]'));
                        kInp.value = row.klasse || '';
                        kInp.addEventListener('input', function () {
                            rows[idx].klasse = kInp.value;
                        });
                    } else {
                        const scEl = /** @type {HTMLSelectElement} */ (line.querySelector('[data-f="scope"]'));
                        scEl.value = normalizeSgaScope(row.scope) || 'student';
                        scEl.addEventListener('change', function () {
                            rows[idx].scope = scEl.value;
                            renderBoard();
                        });
                    }
                    line.querySelector('[data-del]').addEventListener('click', function () {
                        rows.splice(idx, 1);
                        renderBoard();
                    });
                    block.appendChild(line);
                });
                listHost.appendChild(block);
            });
            if (isCouncil && !overlay.querySelector('#schulOrgBoardKlassen')) {
                const dl = document.createElement('datalist');
                dl.id = 'schulOrgBoardKlassen';
                classOptions.forEach(function (c) {
                    const opt = document.createElement('option');
                    opt.value = c;
                    dl.appendChild(opt);
                });
                overlay.appendChild(dl);
            }
        }

        function addRow(preset) {
            const p = preset && typeof preset === 'object' ? preset : {};
            if (isCouncil) {
                rows.push({
                    name: normStr(p.name),
                    email: normEmail(p.email),
                    klasse: normStr(p.klasse).toUpperCase()
                });
            } else {
                rows.push({
                    name: normStr(p.name),
                    email: normEmail(p.email),
                    scope: normalizeSgaScope(p.scope) || (scopeAddEl && scopeAddEl.value) || 'student'
                });
            }
            renderBoard();
        }

        overlay.querySelector('[data-schul-org-board-add]').addEventListener('click', function () {
            addRow({});
        });
        overlay.querySelector('[data-schul-org-board-m365]').addEventListener('click', function () {
            void toggleInlineEntraUserPicker(inlineEupHost, inlineEupState, {
                hint: 'Person in Microsoft 365 suchen und antippen.',
                onPick: function (picked) {
                    addRow({
                        name: normStr(picked.displayName || picked.name),
                        email: normEmail(picked.mail || picked.userPrincipalName || picked.email),
                        scope: scopeAddEl ? scopeAddEl.value : 'student'
                    });
                    toast('Aus Microsoft 365 hinzugefügt.', 'success');
                }
            }).catch(function (e) {
                toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
            });
        });
        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });
        overlay.querySelector('[data-schul-org-save-board]').addEventListener('click', function () {
            const cleaned = [];
            const seen = new Set();
            let badMail = '';
            rows.forEach(function (r) {
                const name = normStr(r.name);
                const email = normEmail(r.email);
                if (!name && !email) return;
                if (email && email.indexOf('@') === -1) {
                    badMail = email;
                    return;
                }
                const key =
                    email ||
                    name.toLowerCase() +
                        '|' +
                        (isCouncil ? normStr(r.klasse) : normalizeSgaScope(r.scope));
                if (seen.has(key)) return;
                seen.add(key);
                if (isCouncil) {
                    cleaned.push({ name: name, email: email, klasse: normStr(r.klasse).toUpperCase() });
                } else {
                    cleaned.push({
                        name: name,
                        email: email,
                        scope: normalizeSgaScope(r.scope) || 'student'
                    });
                }
            });
            if (badMail) {
                toast('Ungültige E-Mail: ' + badMail, 'warning');
                return;
            }
            const cur = loadSettings();
            cur[listKey] = cleaned;
            ensureBoardNodes(cur);
            saveSettings(cur);
            toast(title + ' gespeichert (' + cleaned.length + ' Personen).', 'success');
            closeOverlay();
            notify();
        });

        renderBoard();
    }

    function openCouncilEditor(refKey) {
        if (!refKey || refKey === SCHUL_ORG_BOARD_REF || refKey === 'council_board') {
            openBoardListEditor('council');
            return;
        }
        const seat = svSeatForNode({ id: '', refKey: refKey, type: 'council' }) ||
            SCHUL_ORG_SV_SEATS.find(function (s) {
                return s.id === refKey || s.refKey === refKey;
            });
        if (seat) {
            // Einzelner Sitz → gesamtes SV-Board (drei Funktionen)
            openBoardListEditor('council');
            return;
        }
        openPersonListEditor('council', refKey);
    }

    function openSgaEditor(refKey) {
        if (!refKey || refKey === SCHUL_ORG_BOARD_REF || refKey === 'sga_board') {
            openBoardListEditor('sga');
            return;
        }
        openPersonListEditor('sga', refKey);
    }

    /**
     * @param {string} [parentSectionId]
     * @param {'council'|'sga'} kind
     */
    async function createPersonUnderSection(parentSectionId, kind) {
        openBoardListEditor(kind);
    }

    async function createCouncilUnderSection(parentSectionId) {
        return createPersonUnderSection(parentSectionId, 'council');
    }

    async function createSgaUnderSection(parentSectionId) {
        return createPersonUnderSection(parentSectionId, 'sga');
    }

    function movePersonToNode(person, targetNode) {
        const cur = loadSettings();
        const result = moveSchulOrgPerson(cur, person, targetNode);
        if (!result.ok || !result.settings) return result;
        saveSettings(result.settings);
        notify();
        return result;
    }

    /**
     * Neue Verwaltungsgruppe (Stammdaten) + Organigramm-Bereich.
     * z. B. „Erweiterte Schulleitung“, „Schulwarte“.
     */
    async function createAudienceSection() {
        let name = '';
        if (typeof window.ms365AppDialogPrompt === 'function') {
            name = await window.ms365AppDialogPrompt(
                'Name des neuen Bereichs (wird auch unter Stammdaten → Verwaltung gespeichert)',
                '',
                {
                    title: 'Verwaltungsgruppe anlegen',
                    inputLabel: 'z. B. Erweiterte Schulleitung, Schulwarte …'
                }
            );
        } else {
            name = window.prompt('Name der neuen Verwaltungsgruppe', '') || '';
        }
        const label = normStr(name);
        if (!label) return null;

        const cur = loadSettings();
        const groups = normalizeVerwaltungAudienceGroups(cur.verwaltungAudienceGroups);
        const dup = groups.some(function (g) {
            return normStr(g.label).toLowerCase() === label.toLowerCase();
        });
        if (dup) {
            toast('Diesen Bereich gibt es bereits.', 'warning');
            return null;
        }
        const added = addCustomVerwaltungAudienceGroup(label, groups);
        if (!added.id) {
            toast('Gruppe konnte nicht angelegt werden.', 'warning');
            return null;
        }
        cur.verwaltungAudienceGroups = added.groups;
        if (!Array.isArray(cur.adminAudienceMemberships)) cur.adminAudienceMemberships = [];

        const sid = audienceSectionId(added.id);
        const layout = normalizeSchulOrganigramm(
            cur.schulOrganigramm && Array.isArray(cur.schulOrganigramm.nodes)
                ? cur.schulOrganigramm
                : { version: 2, nodes: [] }
        );
        const maxOrder = layout.nodes.reduce(function (m, n) {
            return Math.max(m, typeof n.order === 'number' ? n.order : 0);
        }, 0);
        // Nach Verwaltung (Personal), vor ARGEs einfügen
        const staff = layout.nodes.find(function (n) {
            return n.id === 'sec_verwaltung';
        });
        const insertOrder =
            staff && typeof staff.order === 'number' ? staff.order + 0.5 : maxOrder + 1;
        layout.nodes.push({
            id: sid,
            type: 'area',
            label: label,
            parentId: null,
            order: insertOrder,
            tone: 'staff',
            refKind: 'audience',
            refKey: added.id
        });
        // Orders neu verdichten (nur Top-Level Areas)
        const roots = layout.nodes
            .filter(function (n) {
                return isAreaNode(n) && (n.parentId == null || n.parentId === '');
            })
            .sort(function (a, b) {
                return (a.order || 0) - (b.order || 0);
            });
        roots.forEach(function (n, i) {
            n.order = i;
        });
        cur.schulOrganigramm = normalizeSchulOrganigramm(layout);
        saveSettings(cur);
        toast('Bereich „' + label + '“ angelegt (auch in Stammdaten → Verwaltung).', 'success');
        notify();
        return { id: added.id, sectionId: sid, label: label };
    }

    /**
     * Bereich anlegen (Top-Level = Verwaltungsgruppe, sonst verschachtelter Layout-Bereich).
     * @param {{ parentId?: string|null, x?: number, y?: number, label?: string }} [opts]
     */
    async function createAreaAt(opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const parentId =
            o.parentId == null || o.parentId === '' ? null : String(o.parentId);
        if (!parentId) {
            const created = await createAudienceSection();
            if (!created || !created.sectionId) return null;
            if (Number.isFinite(Number(o.x)) && Number.isFinite(Number(o.y))) {
                const cur = loadSettings();
                cur.schulOrganigramm = setSchulOrganigrammNodePosition(
                    cur.schulOrganigramm,
                    created.sectionId,
                    Number(o.x),
                    Number(o.y)
                );
                saveSettings(cur);
                notify();
            }
            return { sectionId: created.sectionId, label: created.label, parentId: null };
        }

        let name = normStr(o.label);
        if (!name) {
            if (typeof window.ms365AppDialogPrompt === 'function') {
                name = await window.ms365AppDialogPrompt('Name des neuen Unterbereichs', '', {
                    title: 'Bereich anlegen',
                    inputLabel: 'Bezeichnung'
                });
            } else {
                name = window.prompt('Name des neuen Unterbereichs', '') || '';
            }
            name = normStr(name);
        }
        if (!name) return null;

        const cur = loadSettings();
        const layout = normalizeSchulOrganigramm(cur.schulOrganigramm);
        const parent = layout.nodes.find(function (n) {
            return n.id === parentId && isAreaNode(n);
        });
        if (!parent) {
            toast('Eltern-Bereich nicht gefunden.', 'warning');
            return null;
        }
        const id =
            'area_' +
            Math.random().toString(36).slice(2, 8) +
            Date.now().toString(36).slice(-4);
        const siblings = childrenOf(layout, parentId).filter(isAreaNode);
        // Neu angelegte Unterbereiche oben einordnen
        siblings.forEach(function (s) {
            s.order = (typeof s.order === 'number' ? s.order : 0) + 1;
        });
        const node = {
            id: id,
            type: 'area',
            label: name,
            parentId: parentId,
            order: 0,
            tone: parent.tone || 'staff',
            refKind: 'none',
            refKey: '',
            // Keine festen Relativ-Koordinaten – Canvas packt oben und vergrößert den Elternbereich
            x: undefined,
            y: undefined
        };
        layout.nodes.push(node);
        cur.schulOrganigramm = normalizeSchulOrganigramm(layout);
        saveSettings(cur);
        toast('Unterbereich „' + name + '“ angelegt.', 'success');
        notify();
        return { sectionId: id, label: name, parentId: parentId };
    }

    /**
     * Person aus Canvas-Zoom bearbeiten (Schüler/Lehrer/Admin/SGA …).
     * @param {{ email?: string, name?: string, kind?: string }} person
     * @param {object|null} [contextNode]
     */
    function openEditCanvasPerson(person, contextNode) {
        closeOverlay();
        const p = person && typeof person === 'object' ? person : {};
        const kind = normStr(p.kind).toLowerCase();
        const email0 = normEmail(p.email);
        const name0 = normStr(p.name);

        // Leitung/KV auch unbesetzt → Editor zum Zuordnen
        if (kind === 'head' && contextNode) {
            if (isClassNode(contextNode) && contextNode.refKey) {
                openClassEditor(contextNode.refKey);
                return;
            }
            if (isArgeNode(contextNode) && contextNode.refKey) {
                openArgeEditor(contextNode.refKey);
                return;
            }
        }

        if (!email0 && !name0) {
            toast('Person nicht gefunden.', 'warning');
            return;
        }
        if (kind === 'admin' && contextNode && contextNode.type === NODE.ROLE && contextNode.refKey) {
            openRoleEditor(contextNode.refKey);
            return;
        }
        if (kind === 'council') {
            openCouncilEditor(SCHUL_ORG_BOARD_REF);
            return;
        }
        if (kind === 'sga') {
            openSgaEditor(SCHUL_ORG_BOARD_REF);
            return;
        }

        const settings = loadSettings();
        let klasse0 = '';
        if (kind === 'student') {
            const hit = (Array.isArray(settings.students) ? settings.students : []).find(function (s) {
                return normEmail(s && s.email) === email0 || (!email0 && normStr(s && s.name) === name0);
            });
            if (hit) klasse0 = normStr(hit.klasse);
        }

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: 'bi-person-gear',
            title: 'Person bearbeiten',
            sub:
                kind === 'student'
                    ? 'Schüler:in in den Stammdaten'
                    : kind === 'teacher'
                      ? 'Lehrkraft in den Stammdaten'
                      : 'Person in den Stammdaten',
            body:
                '<div class="schul-org-edit-grid">' +
                '<div class="field"><label for="schulOrgEditPersonName">Name</label>' +
                '<input type="text" id="schulOrgEditPersonName" maxlength="120" autocomplete="name"></div>' +
                '<div class="field"><label for="schulOrgEditPersonEmail">E-Mail</label>' +
                '<input type="email" id="schulOrgEditPersonEmail" maxlength="160" autocomplete="email"></div>' +
                (kind === 'student'
                    ? '<div class="field schul-org-edit-field-full"><label for="schulOrgEditPersonKlasse">Klasse</label>' +
                      '<input type="text" id="schulOrgEditPersonKlasse" maxlength="24" autocomplete="off"></div>'
                    : '') +
                '</div>',
            foot:
                '<span class="schul-org-edit-foot-spacer"></span>' +
                '<button type="button" class="btn alt" data-schul-org-close>Abbrechen</button>' +
                '<button type="button" class="btn btn-success" data-schul-org-save>Speichern</button>'
        });
        document.body.appendChild(overlay);
        activeOverlay = overlay;
        document.addEventListener('keydown', onEsc, true);

        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgEditPersonName'));
        const emailEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgEditPersonEmail'));
        const klasseEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgEditPersonKlasse')
        );
        nameEl.value = name0;
        emailEl.value = email0;
        if (klasseEl) klasseEl.value = klasse0;
        nameEl.focus();

        overlay.querySelector('[data-schul-org-close]').addEventListener('click', closeOverlay);
        overlay.querySelector('[data-schul-org-save]').addEventListener('click', function () {
            const nextName = normStr(nameEl.value);
            const nextEmail = normEmail(emailEl.value);
            if (!nextName && !nextEmail) {
                toast('Name oder E-Mail angeben.', 'warning');
                return;
            }
            const cur = loadSettings();
            let ok = false;
            if (kind === 'student') {
                if (!Array.isArray(cur.students)) cur.students = [];
                const idx = findStudentIndex(cur.students, { email: email0, name: name0 }, klasse0);
                if (idx >= 0) {
                    cur.students[idx] = Object.assign({}, cur.students[idx], {
                        name: nextName || cur.students[idx].name,
                        email: nextEmail || cur.students[idx].email,
                        klasse: klasseEl ? normStr(klasseEl.value) || cur.students[idx].klasse : cur.students[idx].klasse
                    });
                    ok = true;
                }
            } else if (kind === 'teacher') {
                if (!Array.isArray(cur.teachers)) cur.teachers = [];
                const idx = cur.teachers.findIndex(function (t) {
                    return normEmail(t && t.email) === email0 || (!email0 && normStr(t && t.name) === name0);
                });
                if (idx >= 0) {
                    cur.teachers[idx] = Object.assign({}, cur.teachers[idx], {
                        name: nextName || cur.teachers[idx].name,
                        email: nextEmail || cur.teachers[idx].email
                    });
                    ok = true;
                }
            } else if (kind === 'areaPerson' || kind === 'person' || !kind) {
                const areaId =
                    (contextNode && (contextNode.id || contextNode.parentId)) || '';
                if (areaId) {
                    cur.schulOrganigramm = upsertSchulOrganigrammAreaPerson(
                        cur.schulOrganigramm,
                        areaId,
                        { email: nextEmail || email0, name: nextName || name0 }
                    );
                    if (email0 && nextEmail && email0 !== nextEmail) {
                        cur.schulOrganigramm = removeSchulOrganigrammAreaPerson(
                            cur.schulOrganigramm,
                            email0,
                            areaId
                        );
                    }
                    ok = true;
                }
            } else if (kind === 'admin') {
                if (!Array.isArray(cur.admin)) cur.admin = [];
                const idx = findAdminIndex(cur.admin, { email: email0, name: name0 });
                if (idx >= 0) {
                    cur.admin[idx] = Object.assign({}, cur.admin[idx], {
                        name: nextName || cur.admin[idx].name,
                        email: nextEmail || cur.admin[idx].email
                    });
                    syncAdministrationFromAdmin(cur);
                    ok = true;
                }
            }
            if (!ok) {
                toast('Person in den Stammdaten nicht gefunden.', 'warning');
                return;
            }
            saveSettings(cur);
            toast('Person gespeichert.', 'success');
            closeOverlay();
            notify();
        });
    }

    /**
     * Lokale Platzhalter-Mail, wenn nur ein Name eingegeben wird (ohne M365).
     * @param {string} name
     */
    function suggestLocalPersonEmail(name) {
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

    async function openAddCanvasPerson(preselectArea) {
        closeOverlay();
        const settings = loadSettings();
        const layout = normalizeSchulOrganigramm(settings.schulOrganigramm);
        const areas = layout.nodes
            .filter(function (n) {
                if (!isAreaNode(n)) return false;
                if (isSgaBoardArea(n)) return false;
                if (isSammelgruppeArea(n) && !isSgaScopeArea(n)) return false;
                return true;
            })
            .sort(function (a, b) {
                return normStr(a.label).localeCompare(normStr(b.label), 'de');
            });
        if (!areas.length) {
            toast('Kein geeigneter Bereich vorhanden – zuerst einen Bereich anlegen.', 'warning');
            return false;
        }
        const preId = preselectArea && preselectArea.id ? String(preselectArea.id) : '';
        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: 'bi-person-plus',
            title: 'Person anlegen',
            sub: 'Name eintragen und hinzufügen – Microsoft 365 ist optional.',
            body:
                '<div class="field"><label for="schulOrgAddPersonArea">Bereich</label>' +
                '<select id="schulOrgAddPersonArea">' +
                areas
                    .map(function (a) {
                        return (
                            '<option value="' +
                            escapeHtml(a.id) +
                            '"' +
                            (a.id === preId ? ' selected' : '') +
                            '>' +
                            escapeHtml(a.label || a.id) +
                            '</option>'
                        );
                    })
                    .join('') +
                '</select></div>' +
                '<div class="schul-org-edit-grid">' +
                '<div class="field"><label for="schulOrgAddPersonName">Name</label>' +
                '<input type="text" id="schulOrgAddPersonName" maxlength="120" autocomplete="name" placeholder="Vor- und Nachname"></div>' +
                '<div class="field"><label for="schulOrgAddPersonEmail">E-Mail <span class="muted">(optional)</span></label>' +
                '<input type="email" id="schulOrgAddPersonEmail" maxlength="160" autocomplete="email" placeholder="name@schule.at"></div>' +
                '</div>' +
                '<p class="muted schul-org-edit-hint">Ohne E-Mail wird eine lokale Kennung erzeugt – später jederzeit änderbar bzw. aus Microsoft 365 zuordenbar.</p>' +
                '<div class="schul-org-edit-people-actions">' +
                '<button type="button" class="btn btn-sm alt" data-add-person-m365>' +
                '<i class="bi bi-microsoft" aria-hidden="true"></i> Optional: aus Microsoft 365</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>',
            foot:
                '<span class="schul-org-edit-foot-spacer"></span>' +
                '<button type="button" class="btn alt" data-schul-org-close>Abbrechen</button>' +
                '<button type="button" class="btn btn-success" data-schul-org-save>' +
                '<i class="bi bi-check-lg" aria-hidden="true"></i> Person hinzufügen</button>'
        });
        document.body.appendChild(overlay);
        activeOverlay = overlay;
        document.addEventListener('keydown', onEsc, true);

        const areaEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgAddPersonArea'));
        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgAddPersonName'));
        const emailEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgAddPersonEmail'));
        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };
        nameEl.focus();

        function submitAddPerson() {
            const areaId = areaEl.value;
            const areaNode = layout.nodes.find(function (n) {
                return n.id === areaId;
            });
            if (!areaNode) {
                toast('Bereich nicht gefunden.', 'warning');
                return;
            }
            const name = normStr(nameEl.value);
            let email = normEmail(emailEl.value);
            if (!name && !email) {
                toast('Bitte mindestens einen Namen eingeben.', 'warning');
                nameEl.focus();
                return;
            }
            if (!email) {
                email = suggestLocalPersonEmail(name);
                emailEl.value = email;
            }
            const cur = loadSettings();
            const res = addSchulOrgPersonToArea(cur, areaNode, {
                name: name || email,
                email: email
            });
            if (!res.ok) {
                toast(res.message || 'Speichern fehlgeschlagen.', 'warning');
                return;
            }
            saveSettings(res.settings);
            toast('Person hinzugefügt.', 'success');
            closeOverlay();
            notify();
        }

        overlay.querySelector('[data-schul-org-close]').addEventListener('click', closeOverlay);
        overlay.querySelector('[data-add-person-m365]').addEventListener('click', function () {
            void toggleInlineEntraUserPicker(inlineEupHost, inlineEupState, {
                hint: 'Optional: Person in Microsoft 365 suchen – Name und E-Mail werden übernommen.',
                onPick: function (picked) {
                    nameEl.value = picked.displayName || nameEl.value;
                    emailEl.value = picked.mail || picked.userPrincipalName || emailEl.value;
                }
            }).catch(function (e) {
                toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
            });
        });
        overlay.querySelector('[data-schul-org-save]').addEventListener('click', submitAddPerson);
        nameEl.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                emailEl.focus();
            }
        });
        emailEl.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                submitAddPerson();
            }
        });
        return true;
    }

    /**
     * Person per Entra-Picker einem Bereich zuordnen.
     * @param {object} areaNode
     */
    async function pickAndAddPersonToArea(areaNode) {
        if (!areaNode || !isAreaNode(areaNode)) {
            toast('Bitte zuerst einen Bereich wählen.', 'warning');
            return false;
        }
        if (isSgaBoardArea(areaNode)) {
            toast('Rechtsklick auf Schüler, Lehrer oder Eltern unter SGA.', 'info');
            return false;
        }
        if (isSammelgruppeArea(areaNode) && !isSgaScopeArea(areaNode)) {
            toast(
                sammelgruppeKeyForArea(areaNode) === 'lehrer'
                    ? 'Lehrkräfte bitte über die Lehrerliste pflegen.'
                    : 'Schüler:innen bitte über die Klassen bzw. Schülerliste pflegen.',
                'warning'
            );
            return false;
        }
        try {
            const picked = await pickEntraUser({
                title: 'Person für „' + (areaNode.label || 'Bereich') + '“',
                hint: 'Wird dem Bereich zugeordnet.'
            });
            if (!picked) return false;
            const cur = loadSettings();
            const res = addSchulOrgPersonToArea(cur, areaNode, {
                email: picked.mail || picked.userPrincipalName || '',
                name: picked.displayName || ''
            });
            if (!res.ok) {
                toast(res.message || 'Hinzufügen fehlgeschlagen.', 'warning');
                return false;
            }
            saveSettings(res.settings);
            notify();
            toast('Person hinzugefügt.', 'success');
            return true;
        } catch (e) {
            toast('Entra-Auswahl fehlgeschlagen: ' + (e && e.message ? e.message : e), 'warning');
            return false;
        }
    }

    /**
     * Custom-Verwaltungsgruppe löschen (Stammdaten + Section + Kind-Rollen verschieben).
     * @param {string} audienceGroupId va-*
     * @param {{ skipConfirm?: boolean }} [opts]
     */
    async function removeAudienceSection(audienceGroupId, opts) {
        const id = normStr(audienceGroupId);
        if (!id || id.indexOf('va-') !== 0) {
            toast('Nur selbst angelegte Bereiche können entfernt werden.', 'warning');
            return false;
        }
        if (!(opts && opts.skipConfirm)) {
            let ok = true;
            if (typeof window.ms365AppDialogConfirm === 'function') {
                ok = await window.ms365AppDialogConfirm(
                    'Bereich und Verwaltungsgruppe wirklich entfernen? Rollen bleiben erhalten und wandern nach „Verwaltung (Personal)“.',
                    { title: 'Bereich entfernen', okText: 'Entfernen', cancelText: 'Abbrechen' }
                );
            } else {
                ok = window.confirm('Bereich entfernen?');
            }
            if (!ok) return false;
        }

        const cur = loadSettings();
        const removed = removeCustomVerwaltungAudienceGroup(
            id,
            cur.verwaltungAudienceGroups,
            cur.adminAudienceMemberships
        );
        cur.verwaltungAudienceGroups = removed.groups;
        cur.adminAudienceMemberships = removed.memberships;
        const sid = audienceSectionId(id);
        if (cur.schulOrganigramm && Array.isArray(cur.schulOrganigramm.nodes)) {
            cur.schulOrganigramm = normalizeSchulOrganigramm({
                version: cur.schulOrganigramm.version,
                view: cur.schulOrganigramm.view,
                nodes: cur.schulOrganigramm.nodes
                    .filter(function (n) {
                        return n.id !== sid;
                    })
                    .map(function (n) {
                        if (n.parentId === sid) {
                            const hasVerwaltung = (cur.schulOrganigramm.nodes || []).some(
                                function (x) {
                                    return x && x.id === 'sec_verwaltung';
                                }
                            );
                            return Object.assign({}, n, {
                                parentId: hasVerwaltung ? 'sec_verwaltung' : null
                            });
                        }
                        return n;
                    })
            });
        }
        saveSettings(cur);
        toast('Bereich entfernt.', 'success');
        notify();
        return true;
    }

    /**
     * Bereich aus dem Organigramm entfernen (Custom-va-* inkl. Stammdaten, sonst Layout-Knoten).
     * @param {string|object} areaIdOrNode
     */
    async function removeArea(areaIdOrNode) {
        const settings = loadSettings();
        const layout = normalizeSchulOrganigramm(settings.schulOrganigramm);
        const areaId =
            typeof areaIdOrNode === 'string'
                ? normStr(areaIdOrNode)
                : normStr(areaIdOrNode && areaIdOrNode.id);
        const area = layout.nodes.find(function (n) {
            return n.id === areaId && isAreaNode(n);
        });
        if (!area || !canDeleteSchulOrganigrammArea(area)) {
            toast(
                'Dieser Standard-Bereich (z. B. Lehrer, Schüler, ARGE) kann nicht gelöscht werden.',
                'warning'
            );
            return false;
        }
        const audId = audienceIdForArea(area);
        const isCustomAud = audId && String(audId).indexOf('va-') === 0;
        const isOptionalBuiltin = isOptionalDeletableBuiltinArea(area);
        let ok = true;
        if (typeof window.ms365AppDialogConfirm === 'function') {
            ok = await window.ms365AppDialogConfirm(
                isCustomAud
                    ? 'Bereich und Verwaltungsgruppe wirklich entfernen? Rollen bleiben erhalten und wandern nach „Verwaltung (Personal)“ (falls vorhanden) bzw. eine Ebene höher.'
                    : isOptionalBuiltin
                      ? '„' +
                        (area.label || 'Bereich') +
                        '“ von der Landkarte entfernen? Untergeordnete Einträge wandern eine Ebene höher. Der Bereich wird nicht automatisch wieder angelegt.'
                      : 'Bereich „' +
                        (area.label || 'Bereich') +
                        '“ wirklich entfernen? Untergeordnete Einträge wandern eine Ebene höher.',
                { title: 'Bereich entfernen', okText: 'Entfernen', cancelText: 'Abbrechen' }
            );
        } else {
            ok = window.confirm('Bereich entfernen?');
        }
        if (!ok) return false;

        if (isCustomAud) {
            return removeAudienceSection(audId, { skipConfirm: true });
        }

        const cur = loadSettings();
        const fallbackParent =
            area.parentId == null || area.parentId === '' ? null : String(area.parentId);
        const nextNodes = (cur.schulOrganigramm.nodes || [])
            .filter(function (n) {
                return n.id !== area.id;
            })
            .map(function (n) {
                if (n.parentId === area.id) {
                    return Object.assign({}, n, { parentId: fallbackParent });
                }
                return n;
            });
        let nextLayout = normalizeSchulOrganigramm({
            version: cur.schulOrganigramm.version,
            view: cur.schulOrganigramm.view,
            nodes: nextNodes,
            removedAreaIds: cur.schulOrganigramm.removedAreaIds
        });
        if (isOptionalBuiltin) {
            nextLayout = withRemovedSchulOrganigrammArea(nextLayout, area.id);
        }
        cur.schulOrganigramm = nextLayout;
        saveSettings(cur);
        toast('Bereich entfernt.', 'success');
        notify();
        return true;
    }

    /**
     * Area: Personen + M365-Gruppe / Postfach.
     * @param {string|object} areaIdOrNode
     */
    function openAreaEditor(areaIdOrNode) {
        closeOverlay();
        const settings = loadSettings();
        const layout = normalizeSchulOrganigramm(settings.schulOrganigramm);
        const areaId =
            typeof areaIdOrNode === 'string'
                ? normStr(areaIdOrNode)
                : normStr(areaIdOrNode && areaIdOrNode.id);
        let area = layout.nodes.find(function (n) {
            return n.id === areaId && isAreaNode(n);
        });
        if (!area) {
            toast('Bereich nicht gefunden.', 'warning');
            return;
        }
        // Explizites groupMode=none / gelöste Verknüpfung nicht wieder überschreiben
        const hadStoredM365 = !!(area.m365 && typeof area.m365 === 'object');
        area = Object.assign({}, area, {
            m365: area.m365
                ? Object.assign(
                      {
                          groupMode: 'none',
                          groupKind: 'unified',
                          groupId: null,
                          mailboxMode: 'none',
                          mailboxUpn: null
                      },
                      area.m365
                  )
                : {
                      groupMode: 'none',
                      groupKind: 'unified',
                      groupId: null,
                      mailboxMode: 'none',
                      mailboxUpn: null
                  }
        });
        area.m365.groupKind = normalizeSchulOrgGroupKind(area.m365.groupKind || 'unified');
        const sammelKey = sammelgruppeKeyForArea(area);
        // Nur Erstvorschlag, wenn noch nie M365 am Bereich gespeichert wurde
        if (!hadStoredM365) {
            try {
                const container =
                    typeof window !== 'undefined' &&
                    window.ms365AppDataV2 &&
                    typeof window.ms365AppDataV2.getContainer === 'function'
                        ? window.ms365AppDataV2.getContainer()
                        : null;
                const setup = container && container.setup ? container.setup : {};
                const matched = setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
                if (sammelKey === 'lehrer' || sammelKey === 'schueler') {
                    const gid =
                        sammelKey === 'lehrer'
                            ? String(matched.lehrerGroupId || '').trim()
                            : String(matched.schuelerGroupId || '').trim();
                    if (gid) {
                        area.m365.groupMode = 'match';
                        area.m365.groupKind = 'unified';
                        area.m365.groupId = gid;
                        area.m365.groupLabel =
                            sammelKey === 'lehrer' ? 'Lehrer-Sammelgruppe' : 'Schüler-Sammelgruppe';
                    }
                } else if (isClassAreaNode(area) && area.refKey) {
                    // Geerbte Klassengruppe anzeigen, damit „Gruppe lösen“ sie wirklich abschalten kann
                    const cls = findClass(settings, area.refKey);
                    const classTeams =
                        container && container.core && Array.isArray(container.core.classTeams)
                            ? container.core.classTeams
                            : [];
                    const matchMap =
                        setup.classGroupMatchByKey && typeof setup.classGroupMatchByKey === 'object'
                            ? setup.classGroupMatchByKey
                            : {};
                    const gid = cls ? resolveClassGraphGroupId(cls, classTeams, matchMap) : '';
                    if (gid) {
                        area.m365.groupMode = 'match';
                        area.m365.groupKind = 'unified';
                        area.m365.groupId = gid;
                        area.m365.groupLabel = 'Klassengruppe (Stammdaten)';
                    }
                }
            } catch {
                /* ignore */
            }
        }

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        const audId = audienceIdForArea(area);
        const jgKey = jahrgangKeyForArea(area);
        const peopleSub = sammelKey
            ? sammelKey === 'lehrer'
                ? 'Mitglieder aus der Lehrerliste · Standard: M365-Sammelgruppe (Einrichtung)'
                : 'Jahrgänge & Klassen · Standard: M365-Sammelgruppe (Einrichtung)'
            : jgKey
              ? 'Klassen dieses Jahrgangs · optional Gruppe/Verteiler/Postfach zuordnen oder anlegen'
              : audId
                ? 'Personen · M365-Objekt matchen oder neu anlegen'
                : 'Struktur zuerst – dann M365-Gruppe, Sicherheitsgruppe, Verteiler oder Postfach';
        overlay.innerHTML = editDialogHtml({
            icon: 'bi-pencil-square',
            title: 'Bereich bearbeiten',
            sub: peopleSub,
            body:
                '<div class="field"><label for="schulOrgAreaLabel">Bezeichnung</label>' +
                '<input type="text" id="schulOrgAreaLabel" value="' +
                escapeHtml(area.label) +
                '"></div>' +
                '<div class="schul-org-area-people">' +
                '<div class="schul-org-area-people__head">' +
                '<strong>Personen</strong>' +
                (sammelKey
                    ? '<span class="muted">aus Stammdaten</span>'
                    : '<div class="schul-org-edit-people-actions">' +
                      '<button type="button" class="btn btn-sm btn-success" data-area-add-person-manual>' +
                      '<i class="bi bi-person-plus" aria-hidden="true"></i> Person</button>' +
                      '<button type="button" class="btn btn-sm btn-primary" data-area-add-person>' +
                      '<i class="bi bi-microsoft" aria-hidden="true"></i> Aus Microsoft 365</button>' +
                      '</div>') +
                '</div>' +
                (sammelKey
                    ? ''
                    : '<div class="schul-org-area-people__manual" data-area-manual-form hidden>' +
                      '<div class="schul-org-edit-grid">' +
                      '<div class="field"><label for="schulOrgAreaPersonName">Name</label>' +
                      '<input type="text" id="schulOrgAreaPersonName" maxlength="120" autocomplete="name" placeholder="Vor- und Nachname"></div>' +
                      '<div class="field"><label for="schulOrgAreaPersonEmail">E-Mail <span class="muted">(optional)</span></label>' +
                      '<input type="email" id="schulOrgAreaPersonEmail" maxlength="160" autocomplete="email" placeholder="name@schule.at"></div>' +
                      '</div>' +
                      '<p class="muted schul-org-edit-hint">Ohne E-Mail wird eine lokale Kennung erzeugt.</p>' +
                      '<button type="button" class="btn btn-sm btn-success" data-area-manual-submit>' +
                      '<i class="bi bi-check-lg" aria-hidden="true"></i> Hinzufügen</button>' +
                      '</div>') +
                '<ul class="schul-org-area-people__list" data-area-people-list></ul>' +
                (sammelKey
                    ? ''
                    : '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>') +
                '</div>' +
                '<fieldset class="schul-org-area-m365">' +
                '<legend>Microsoft 365</legend>' +
                '<p class="muted schul-org-area-m365__hint">Beliebigen Bereich mit einem Entra-/Exchange-Objekt verbinden – zuerst Struktur, dann matchen oder anlegen. Lehrer, Schüler, ARGEs und Klassen nutzen die bestehenden Standardgruppen.</p>' +
                '<div class="schul-org-area-m365__grid">' +
                '<div class="field"><label for="schulOrgAreaGroupKind">Objekttyp</label>' +
                '<select id="schulOrgAreaGroupKind">' +
                '<option value="unified">Microsoft 365-Gruppe</option>' +
                '<option value="security">Sicherheitsgruppe</option>' +
                '<option value="mailSecurity">E-Mail-Sicherheitsgruppe</option>' +
                '<option value="distribution">Verteilerliste</option>' +
                '</select></div>' +
                '<div class="field"><label for="schulOrgAreaGroupMode">Aktion</label>' +
                '<select id="schulOrgAreaGroupMode">' +
                '<option value="none">Keine Gruppe</option>' +
                '<option value="match">Bestehende zuordnen</option>' +
                '<option value="create">Neu in Microsoft 365 anlegen</option>' +
                '</select></div></div>' +
                '<p class="muted schul-org-area-m365__group-sum" data-area-group-sum></p>' +
                '<div class="schul-org-area-m365__row" data-area-group-match-row>' +
                '<button type="button" class="btn btn-sm" data-area-pick-group><i class="bi bi-search" aria-hidden="true"></i> In Entra wählen</button>' +
                '<button type="button" class="btn btn-sm alt" data-area-clear-group>Gruppe lösen</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-egp hidden></div>' +
                '<div class="schul-org-area-m365__create" data-area-group-create-row hidden>' +
                '<p class="muted schul-org-area-m365__group-people" data-area-group-people></p>' +
                '<div class="field"><label for="schulOrgAreaCreateName">Anzeigename</label>' +
                '<input type="text" id="schulOrgAreaCreateName" value="' +
                escapeHtml(area.label) +
                '"></div>' +
                '<div class="field"><label for="schulOrgAreaCreateNick">Mail-Nickname / Alias</label>' +
                '<input type="text" id="schulOrgAreaCreateNick" value="' +
                escapeHtml(suggestSchulOrgMailNickname(area.label)) +
                '" autocomplete="off"></div>' +
                '<div class="field"><label for="schulOrgAreaGroupOwner">Besitzer</label>' +
                '<select id="schulOrgAreaGroupOwner" data-area-group-owner>' +
                '<option value="">— aus Bereich wählen —</option>' +
                '</select>' +
                '<p class="muted schul-org-edit-hint">Personen aus diesem Bereich, Rollen und Unterbereichen werden flach als Mitglieder hinzugefügt (keine verschachtelten M365-Gruppen). Ein Besitzer ist wählbar.</p></div>' +
                '<button type="button" class="btn btn-sm btn-primary" data-area-create-group>' +
                '<i class="bi bi-cloud-plus" aria-hidden="true"></i> Jetzt in Microsoft 365 anlegen</button>' +
                '<p class="muted schul-org-area-m365__create-hint" data-area-create-hint></p>' +
                '</div>' +
                '<div class="field"><label for="schulOrgAreaMailboxMode">Freigegebenes Postfach</label>' +
                '<select id="schulOrgAreaMailboxMode">' +
                '<option value="none">Keins</option>' +
                '<option value="match">Bestehendes zuordnen</option>' +
                '<option value="create">Neu in Exchange anlegen</option>' +
                '</select></div>' +
                '<p class="muted schul-org-area-m365__mail-sum" data-area-mail-sum></p>' +
                '<div class="schul-org-area-m365__row" data-area-mail-match-row>' +
                '<button type="button" class="btn btn-sm" data-area-pick-mail>In Entra wählen</button>' +
                '<button type="button" class="btn btn-sm alt" data-area-clear-mail>Postfach lösen</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup-mail hidden></div>' +
                '<div class="schul-org-area-m365__create" data-area-mail-create-row hidden>' +
                '<p class="muted schul-org-area-m365__mail-hint">Freigegebene Postfächer legt Exchange Online an (nicht Graph). Es entsteht eine doppelklickbare <strong>.cmd</strong>-Datei mit geführtem Assistenten (Admin-Konto / Tenant wählen, bestätigen, anlegen). Personen aus Bereich, Rollen und Unterbereichen erhalten FullAccess und SendAs.</p>' +
                '<p class="muted schul-org-area-m365__mail-delegates" data-area-mail-delegates></p>' +
                '<div class="field"><label for="schulOrgAreaMailCreateName">Anzeigename</label>' +
                '<input type="text" id="schulOrgAreaMailCreateName" value="' +
                escapeHtml(area.label) +
                '" autocomplete="off"></div>' +
                '<div class="field"><label for="schulOrgAreaMailCreateAlias">Alias</label>' +
                '<input type="text" id="schulOrgAreaMailCreateAlias" value="' +
                escapeHtml(suggestSchulOrgMailNickname(area.label)) +
                '" autocomplete="off"></div>' +
                '<div class="field"><label for="schulOrgAreaMailCreateSmtp">Primäre E-Mail</label>' +
                '<input type="email" id="schulOrgAreaMailCreateSmtp" value="" autocomplete="off" placeholder="alias@schule.at"></div>' +
                '<div class="schul-org-edit-people-actions">' +
                '<button type="button" class="btn btn-sm btn-primary" data-area-create-mail>' +
                '<i class="bi bi-terminal" aria-hidden="true"></i> Assistent erzeugen &amp; vormerken</button>' +
                '<button type="button" class="btn btn-sm" data-area-mail-script-copy hidden>' +
                '<i class="bi bi-clipboard" aria-hidden="true"></i> PowerShell kopieren</button>' +
                '<button type="button" class="btn btn-sm btn-success" data-area-mail-script-dl hidden>' +
                '<i class="bi bi-download" aria-hidden="true"></i> .cmd herunterladen</button>' +
                '</div>' +
                '<textarea class="schul-org-area-m365__script" data-area-mail-script hidden readonly rows="10" aria-label="Exchange-PowerShell-Assistent"></textarea>' +
                '</div>' +
                '</fieldset>',
            foot: canDeleteSchulOrganigrammArea(area)
                ? editFootHtml('data-schul-org-delete-area', 'Bereich löschen', 'data-schul-org-save')
                : '<span class="schul-org-edit-foot-spacer"></span>' +
                  '<button type="button" class="btn alt" data-schul-org-close>Abbrechen</button>' +
                  '<button type="button" class="btn btn-success" data-schul-org-save>Speichern</button>'
        });
        document.body.appendChild(overlay);
        activeOverlay = overlay;
        document.addEventListener('keydown', onEsc, true);

        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        const inlineEgpHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-egp]')
        );
        const inlineMailHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup-mail]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEgpState = { api: null };
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineMailState = { api: null };
        const inlinePickers = [inlineEupState, inlineEgpState, inlineMailState];

        const delAreaBtn = overlay.querySelector('[data-schul-org-delete-area]');
        if (delAreaBtn) {
            delAreaBtn.addEventListener('click', function () {
                void removeArea(area).then(function (ok) {
                    if (ok) closeOverlay();
                });
            });
        }

        const groupModeEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgAreaGroupMode'));
        const groupKindEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgAreaGroupKind'));
        const mailboxModeEl = /** @type {HTMLSelectElement} */ (
            overlay.querySelector('#schulOrgAreaMailboxMode')
        );
        const createNameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgAreaCreateName'));
        const createNickEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgAreaCreateNick'));
        const mailCreateNameEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgAreaMailCreateName')
        );
        const mailCreateAliasEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgAreaMailCreateAlias')
        );
        const mailCreateSmtpEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgAreaMailCreateSmtp')
        );
        const mailScriptEl = /** @type {HTMLTextAreaElement|null} */ (
            overlay.querySelector('[data-area-mail-script]')
        );
        const mailScriptCopyBtn = /** @type {HTMLButtonElement|null} */ (
            overlay.querySelector('[data-area-mail-script-copy]')
        );
        const mailScriptDlBtn = /** @type {HTMLButtonElement|null} */ (
            overlay.querySelector('[data-area-mail-script-dl]')
        );
        const mailDelegatesEl = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-area-mail-delegates]')
        );
        const groupPeopleEl = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-area-group-people]')
        );
        const groupOwnerEl = /** @type {HTMLSelectElement|null} */ (
            overlay.querySelector('[data-area-group-owner]')
        );
        /** @type {string} */
        let mailCmdDownloadText = '';

        function suggestedExchangeAdminUpn() {
            try {
                if (typeof window.ms365AuthGetAccountInfo === 'function') {
                    const info = window.ms365AuthGetAccountInfo();
                    const upn = normEmail((info && (info.upn || info.username)) || '');
                    if (upn && upn.indexOf('@') !== -1) return upn;
                }
            } catch {
                /* ignore */
            }
            return '';
        }
        groupModeEl.value = area.m365.groupMode || 'none';
        groupKindEl.value = normalizeSchulOrgGroupKind(area.m365.groupKind || 'unified');
        mailboxModeEl.value =
            area.m365.mailboxMode === 'create' ||
            area.m365.mailboxMode === 'match' ||
            area.m365.mailboxMode === 'none'
                ? area.m365.mailboxMode
                : 'none';

        function schoolDomain() {
            const cur = loadSettings();
            return normStr(cur && cur.domain).replace(/^@/, '').toLowerCase();
        }

        function areaM365People() {
            return listM365PeopleForArea(loadSettings(), area);
        }

        function refreshGroupPeopleUi() {
            const people = areaM365People();
            if (groupPeopleEl) {
                if (!people.length) {
                    groupPeopleEl.textContent =
                        'Noch keine Personen mit M365-E-Mail in diesem Bereich, seinen Rollen oder Unterbereichen – die Gruppe wird ohne Mitglieder angelegt.';
                } else {
                    const nested = people.filter(function (p) {
                        return p.source === 'nested';
                    }).length;
                    groupPeopleEl.textContent =
                        'Als Mitglieder vorgesehen (' +
                        people.length +
                        (nested
                            ? ', davon ' + nested + ' aus Unterbereichen'
                            : '') +
                        '): ' +
                        people
                            .map(function (p) {
                                return p.label + (p.email ? ' <' + p.email + '>' : '');
                            })
                            .join(', ');
                }
            }
            if (groupOwnerEl) {
                const prev = normEmail(groupOwnerEl.value);
                groupOwnerEl.innerHTML =
                    '<option value="">— aus Bereich wählen —</option>' +
                    people
                        .map(function (p) {
                            return (
                                '<option value="' +
                                escapeHtml(p.email) +
                                '"' +
                                (prev && prev === p.email ? ' selected' : '') +
                                '>' +
                                escapeHtml(p.label + ' (' + p.email + ')') +
                                '</option>'
                            );
                        })
                        .join('');
                if (!prev && people.length === 1) {
                    groupOwnerEl.value = people[0].email;
                } else if (prev && people.some(function (p) {
                    return p.email === prev;
                })) {
                    groupOwnerEl.value = prev;
                } else if (!prev && people.length > 1) {
                    groupOwnerEl.value = people[0].email;
                }
            }
        }

        function refreshMailDelegatesHint() {
            if (!mailDelegatesEl) return;
            const dels = areaM365People();
            if (!dels.length) {
                mailDelegatesEl.textContent =
                    'Noch keine Personen mit M365-E-Mail in diesem Bereich, seinen Rollen oder Unterbereichen – der Befehl legt nur das Postfach an.';
                return;
            }
            const nestedMail = dels.filter(function (p) {
                return p.source === 'nested';
            }).length;
            mailDelegatesEl.textContent =
                'Berechtigungen (FullAccess + SendAs) für ' +
                dels.length +
                (dels.length === 1 ? ' Person' : ' Personen') +
                (nestedMail
                    ? ' (davon ' + nestedMail + ' aus Unterbereichen)'
                    : '') +
                ': ' +
                dels
                    .map(function (p) {
                        return p.label + (p.email ? ' <' + p.email + '>' : '');
                    })
                    .join(', ');
        }

        function ensureMailCreateDefaults() {
            if (!mailCreateNameEl || !mailCreateAliasEl || !mailCreateSmtpEl) return;
            if (!normStr(mailCreateNameEl.value)) mailCreateNameEl.value = area.label || '';
            refreshMailDelegatesHint();
            void import('../../shared/exchange-shared-mailbox-script.js').then(function (mod) {
                const sug = mod.suggestSharedMailboxIdentity(
                    mailCreateNameEl.value || area.label,
                    schoolDomain()
                );
                if (!normStr(mailCreateAliasEl.value)) mailCreateAliasEl.value = sug.alias;
                if (!normStr(mailCreateSmtpEl.value) && sug.primarySmtp) {
                    mailCreateSmtpEl.value = sug.primarySmtp;
                }
            });
        }

        function syncM365ModeUi() {
            const mode = groupModeEl.value || 'none';
            const matchRow = overlay.querySelector('[data-area-group-match-row]');
            const createRow = overlay.querySelector('[data-area-group-create-row]');
            const hint = overlay.querySelector('[data-area-create-hint]');
            if (matchRow) matchRow.hidden = mode === 'create';
            if (createRow) createRow.hidden = mode !== 'create';
            if (mode === 'create' || mode === 'none') {
                if (inlineEgpState.api) {
                    inlineEgpState.api.destroy();
                    inlineEgpState.api = null;
                }
                if (inlineEgpHost) inlineEgpHost.hidden = true;
            }
            if (hint) {
                const kind = normalizeSchulOrgGroupKind(groupKindEl.value);
                const memberNote =
                    ' Personen aus dem Bereich, Rollen und Unterbereichen werden flach als Mitglieder gesetzt; Besitzer ist wählbar.';
                if (kind === 'distribution') {
                    hint.textContent =
                        'Verteilerlisten: Graph legt eine mailaktivierte Cloud-Gruppe an (ohne Team). Bei Tenant-Einschränkungen ggf. Exchange nutzen.' +
                        memberNote;
                } else if (kind === 'security') {
                    hint.textContent = 'Reine Sicherheitsgruppe (ohne E-Mail) in Entra.' + memberNote;
                } else if (kind === 'mailSecurity') {
                    hint.textContent = 'E-Mail-aktivierte Sicherheitsgruppe in Entra.' + memberNote;
                } else {
                    hint.textContent =
                        'Microsoft 365-Gruppe (Unified) – optional später Team möglich.' + memberNote;
                }
            }
            const mailMode = mailboxModeEl.value || 'none';
            const mailMatchRow = overlay.querySelector('[data-area-mail-match-row]');
            const mailCreateRow = overlay.querySelector('[data-area-mail-create-row]');
            if (mailMatchRow) mailMatchRow.hidden = mailMode === 'create';
            if (mailCreateRow) mailCreateRow.hidden = mailMode !== 'create';
            if (mode === 'create') {
                refreshGroupPeopleUi();
            }
            if (mailMode === 'create') {
                ensureMailCreateDefaults();
                if (inlineMailState.api) {
                    inlineMailState.api.destroy();
                    inlineMailState.api = null;
                }
                if (inlineMailHost) inlineMailHost.hidden = true;
            } else if (mailMode === 'none') {
                if (inlineMailState.api) {
                    inlineMailState.api.destroy();
                    inlineMailState.api = null;
                }
                if (inlineMailHost) inlineMailHost.hidden = true;
            }
        }

        function refreshM365Sums() {
            const gSum = overlay.querySelector('[data-area-group-sum]');
            const mSum = overlay.querySelector('[data-area-mail-sum]');
            if (gSum) {
                if (area.m365.groupMode === 'none' || !area.m365.groupId) {
                    gSum.textContent =
                        groupModeEl.value === 'create'
                            ? 'Noch nicht angelegt – Name/Alias prüfen und anlegen.'
                            : 'Keine Gruppe verknüpft.';
                } else {
                    gSum.textContent = [
                        schulOrgGroupKindLabel(area.m365.groupKind),
                        area.m365.groupLabel,
                        area.m365.groupMail,
                        area.m365.groupId
                    ]
                        .filter(Boolean)
                        .join(' · ');
                }
            }
            if (mSum) {
                if (mailboxModeEl.value === 'create' && area.m365.mailboxUpn) {
                    mSum.textContent =
                        'Zum Anlegen vorgemerkt: ' +
                        [area.m365.mailboxLabel, area.m365.mailboxUpn].filter(Boolean).join(' · ') +
                        ' (.cmd-Assistent in Exchange ausführen)';
                } else if (area.m365.mailboxMode === 'none' || !area.m365.mailboxUpn) {
                    mSum.textContent =
                        mailboxModeEl.value === 'create'
                            ? 'Noch nicht vorgemerkt – Name/Alias/E-Mail prüfen und Assistent erzeugen.'
                            : 'Kein Postfach verknüpft.';
                } else {
                    mSum.textContent = [area.m365.mailboxLabel, area.m365.mailboxUpn]
                        .filter(Boolean)
                        .join(' · ');
                }
            }
            syncM365ModeUi();
        }

        function renderPeopleList() {
            const list = overlay.querySelector('[data-area-people-list]');
            if (!list) return;
            const cur = loadSettings();
            const people = listPeopleForSchulOrgNode(cur, area);
            if (!people.length) {
                list.innerHTML =
                    '<li class="muted">' +
                    (sammelKey === 'lehrer'
                        ? 'Noch keine Lehrkräfte in den Stammdaten.'
                        : sammelKey === 'schueler'
                          ? 'Keine direkten Personen – Klassen liegen in den Jahrgangs-Bereichen.'
                          : 'Noch keine Personen – „Person“ oder Microsoft 365 nutzen.') +
                    '</li>';
                return;
            }
            list.innerHTML = people
                .slice(0, sammelKey ? 80 : people.length)
                .map(function (p) {
                    const em = normEmail(p.email);
                    const localMail = /@local\.schul-org$/i.test(em);
                    return (
                        '<li class="schul-org-area-people__item">' +
                        '<span><strong>' +
                        escapeHtml(p.name || p.email) +
                        '</strong>' +
                        (em && !localMail
                            ? '<br/><span class="muted">' + escapeHtml(em) + '</span>'
                            : localMail
                              ? '<br/><span class="muted">ohne E-Mail</span>'
                              : '') +
                        '</span>' +
                        (sammelKey
                            ? ''
                            : '<button type="button" class="btn btn-sm alt" data-area-remove-email="' +
                              escapeHtml(p.email) +
                              '">Entfernen</button>') +
                        '</li>'
                    );
                })
                .join('');
            if (sammelKey && people.length > 80) {
                list.innerHTML +=
                    '<li class="muted">… und ' + (people.length - 80) + ' weitere in den Stammdaten</li>';
            }
            list.querySelectorAll('[data-area-remove-email]').forEach(function (btn) {
                btn.addEventListener('click', function () {
                    const em = btn.getAttribute('data-area-remove-email') || '';
                    const st = loadSettings();
                    const res = removeSchulOrgPersonFromArea(st, area, em);
                    if (!res.ok) {
                        toast(res.message || 'Entfernen fehlgeschlagen.', 'warning');
                        return;
                    }
                    saveSettings(res.settings);
                    renderPeopleList();
                    refreshGroupPeopleUi();
                    refreshMailDelegatesHint();
                    notify();
                });
            });
        }

        refreshM365Sums();
        renderPeopleList();
        refreshGroupPeopleUi();
        refreshMailDelegatesHint();

        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });

        const manualForm = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-area-manual-form]')
        );
        const manualNameEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgAreaPersonName')
        );
        const manualEmailEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgAreaPersonEmail')
        );

        function submitManualPerson() {
            if (!manualNameEl || !manualEmailEl) return;
            const name = normStr(manualNameEl.value);
            let email = normEmail(manualEmailEl.value);
            if (!name && !email) {
                toast('Bitte mindestens einen Namen eingeben.', 'warning');
                manualNameEl.focus();
                return;
            }
            if (!email) {
                email = suggestLocalPersonEmail(name);
                manualEmailEl.value = email;
            }
            const st = loadSettings();
            const res = addSchulOrgPersonToArea(st, area, {
                name: name || email,
                email: email
            });
            if (!res.ok) {
                toast(res.message || 'Hinzufügen fehlgeschlagen.', 'warning');
                return;
            }
            saveSettings(res.settings);
            manualNameEl.value = '';
            manualEmailEl.value = '';
            if (manualForm) manualForm.hidden = true;
            renderPeopleList();
            refreshGroupPeopleUi();
            refreshMailDelegatesHint();
            notify();
            toast('Person hinzugefügt.', 'success');
        }

        const addManualBtn = overlay.querySelector('[data-area-add-person-manual]');
        if (addManualBtn && manualForm) {
            addManualBtn.addEventListener('click', function () {
                if (inlineEupState.api) {
                    inlineEupState.api.destroy();
                    inlineEupState.api = null;
                }
                if (inlineEupHost) inlineEupHost.hidden = true;
                manualForm.hidden = !manualForm.hidden;
                if (!manualForm.hidden && manualNameEl) manualNameEl.focus();
            });
        }
        const manualSubmitBtn = overlay.querySelector('[data-area-manual-submit]');
        if (manualSubmitBtn) {
            manualSubmitBtn.addEventListener('click', submitManualPerson);
        }
        if (manualEmailEl) {
            manualEmailEl.addEventListener('keydown', function (ev) {
                if (ev.key === 'Enter') {
                    ev.preventDefault();
                    submitManualPerson();
                }
            });
        }
        if (manualNameEl) {
            manualNameEl.addEventListener('keydown', function (ev) {
                if (ev.key === 'Enter') {
                    ev.preventDefault();
                    if (manualEmailEl) manualEmailEl.focus();
                }
            });
        }

        const addPersonBtn = overlay.querySelector('[data-area-add-person]');
        if (addPersonBtn) {
            addPersonBtn.addEventListener('click', function () {
                if (manualForm) manualForm.hidden = true;
                void toggleInlineEntraUserPicker(
                    inlineEupHost,
                    inlineEupState,
                    {
                        hint: 'Person für „' + area.label + '“ in Microsoft 365 suchen und antippen.',
                        onPick: function (picked) {
                            const st = loadSettings();
                            const res = addSchulOrgPersonToArea(st, area, {
                                email: picked.mail || picked.userPrincipalName || '',
                                name: picked.displayName || ''
                            });
                            if (!res.ok) {
                                toast(res.message || 'Hinzufügen fehlgeschlagen.', 'warning');
                                return;
                            }
                            saveSettings(res.settings);
                            renderPeopleList();
                            refreshGroupPeopleUi();
                            refreshMailDelegatesHint();
                            notify();
                            toast('Person hinzugefügt.', 'success');
                        }
                    },
                    inlinePickers
                ).catch(function (e) {
                    toast(
                        'Entra-Auswahl fehlgeschlagen: ' + (e && e.message ? e.message : e),
                        'warning'
                    );
                });
            });
        }

        groupModeEl.addEventListener('change', function () {
            area.m365.groupMode = groupModeEl.value || 'none';
            if (area.m365.groupMode === 'none') {
                area.m365.groupId = null;
                area.m365.groupLabel = null;
                area.m365.groupMail = null;
            }
            if (groupModeEl.value === 'create' && createNickEl && !normStr(createNickEl.value)) {
                createNickEl.value = suggestSchulOrgMailNickname(
                    (createNameEl && createNameEl.value) || area.label
                );
            }
            refreshM365Sums();
        });
        groupKindEl.addEventListener('change', function () {
            area.m365.groupKind = normalizeSchulOrgGroupKind(groupKindEl.value);
            refreshM365Sums();
        });

        overlay.querySelector('[data-area-pick-group]').addEventListener('click', function () {
            if (groupModeEl.value === 'none') groupModeEl.value = 'match';
            void toggleInlineEntraGroupPicker(
                inlineEgpHost,
                inlineEgpState,
                {
                    hint:
                        'Gruppe für „' +
                        area.label +
                        '“ suchen (M365, Sicherheit, Verteiler) und antippen.',
                    onPick: function (picked) {
                        area.m365.groupMode = 'match';
                        groupModeEl.value = 'match';
                        area.m365.groupKind = inferSchulOrgGroupKind(picked);
                        groupKindEl.value = area.m365.groupKind;
                        area.m365.groupId = picked.id || null;
                        area.m365.groupLabel = picked.displayName || picked.name || '';
                        area.m365.groupMail = normEmail(picked.mail || picked.email || '');
                        refreshM365Sums();
                        toast(schulOrgGroupKindLabel(area.m365.groupKind) + ' zugeordnet.', 'success');
                    }
                },
                inlinePickers
            ).catch(function (e) {
                toast('Gruppenauswahl fehlgeschlagen: ' + (e && e.message ? e.message : e), 'warning');
            });
        });
        overlay.querySelector('[data-area-clear-group]').addEventListener('click', function () {
            groupModeEl.value = 'none';
            area.m365.groupMode = 'none';
            area.m365.groupId = null;
            area.m365.groupLabel = null;
            area.m365.groupMail = null;
            refreshM365Sums();
            // Sofort speichern – sonst bleibt die falsche Zuordnung auf der Landkarte
            const cur = loadSettings();
            let nextLayout = normalizeSchulOrganigramm(cur.schulOrganigramm);
            const labelEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgAreaLabel'));
            const nextLabel = normStr(labelEl && labelEl.value) || area.label;
            const node = nextLayout.nodes.find(function (n) {
                return n.id === area.id;
            });
            if (node) node.label = nextLabel;
            nextLayout = setSchulOrganigrammAreaM365(nextLayout, area.id, {
                groupMode: 'none',
                groupKind: normalizeSchulOrgGroupKind(groupKindEl.value || area.m365.groupKind),
                groupId: null,
                groupLabel: null,
                groupMail: null,
                mailboxMode: area.m365.mailboxMode || 'none',
                mailboxUpn: area.m365.mailboxUpn || null,
                mailboxLabel: area.m365.mailboxLabel || null,
                mailboxId: area.m365.mailboxId || null
            });
            cur.schulOrganigramm = nextLayout;
            try {
                window.__ms365SchulOrgSkipSyncOnce = true;
            } catch {
                /* ignore */
            }
            saveSettings(cur);
            // Arbeitskopie an gespeicherten Stand anbinden
            const savedNode = nextLayout.nodes.find(function (n) {
                return n.id === area.id;
            });
            if (savedNode && savedNode.m365) {
                area.m365 = Object.assign({}, savedNode.m365);
            }
            refreshM365Sums();
            toast('Gruppe gelöst und gespeichert.', 'success');
            notify();
        });

        const createBtn = overlay.querySelector('[data-area-create-group]');
        if (createBtn) {
            createBtn.addEventListener('click', function () {
                const kind = normalizeSchulOrgGroupKind(groupKindEl.value);
                const displayName =
                    normStr(createNameEl && createNameEl.value) || normStr(area.label) || 'Bereich';
                const nick =
                    suggestSchulOrgMailNickname(
                        normStr(createNickEl && createNickEl.value) || displayName
                    );
                if (createNickEl) createNickEl.value = nick;
                const gug = window.ms365GraphUnifiedGroups;
                if (!gug || typeof gug.createGroupByKind !== 'function') {
                    toast('Graph-Modul für Gruppenanlage nicht geladen.', 'warning');
                    return;
                }
                createBtn.disabled = true;
                void (async function () {
                    try {
                        let token = '';
                        if (typeof window.ms365AuthAcquireTokenSilent === 'function') {
                            try {
                                token = (await window.ms365AuthAcquireTokenSilent()) || '';
                            } catch {
                                token = '';
                            }
                        }
                        if (!token && typeof gug.getGraphToken === 'function') {
                            token = await gug.getGraphToken();
                        }
                        if (!token && typeof window.ms365AuthAcquireToken === 'function') {
                            token = await window.ms365AuthAcquireToken();
                        }
                        if (!token) throw new Error('Anmeldung für Microsoft Graph fehlgeschlagen.');
                        const people = areaM365People();
                        const memberEmails = people.map(function (p) {
                            return p.email;
                        });
                        let ownerEmail = normEmail(groupOwnerEl && groupOwnerEl.value);
                        if (!ownerEmail && memberEmails.length) ownerEmail = memberEmails[0];
                        const created = await gug.createGroupByKind(
                            token,
                            kind,
                            displayName,
                            nick,
                            'Schul-Organigramm: ' + displayName
                        );
                        const gid = created && created.id ? String(created.id) : '';
                        if (!gid) throw new Error('Gruppe angelegt, aber keine ID zurückgegeben.');
                        let memberOk = 0;
                        let ownerOk = false;
                        if (typeof gug.sleep === 'function') await gug.sleep(1200);
                        if (ownerEmail) {
                            try {
                                if (typeof gug.resolveUserByEmail === 'function') {
                                    const ownerUser = await gug.resolveUserByEmail(token, ownerEmail);
                                    if (ownerUser && ownerUser.id) {
                                        if (typeof gug.addOwnerWithMemberFallback === 'function') {
                                            await gug.addOwnerWithMemberFallback(
                                                token,
                                                gid,
                                                ownerUser.id
                                            );
                                            ownerOk = true;
                                        } else if (typeof gug.ensureOwners === 'function') {
                                            await gug.ensureOwners(token, gid, [ownerEmail]);
                                            ownerOk = true;
                                        }
                                    }
                                } else if (typeof gug.ensureOwners === 'function') {
                                    await gug.ensureOwners(token, gid, [ownerEmail]);
                                    ownerOk = true;
                                }
                            } catch (ownerErr) {
                                console.warn(
                                    '[schul-org] Besitzer setzen fehlgeschlagen',
                                    ownerErr
                                );
                            }
                        } else if (typeof gug.ensureOwners === 'function') {
                            await gug.ensureOwners(token, gid, []);
                        }
                        if (memberEmails.length && typeof gug.syncEmailsToGroup === 'function') {
                            const sync = await gug.syncEmailsToGroup(
                                token,
                                gid,
                                memberEmails,
                                displayName
                            );
                            memberOk = (sync && sync.ok) || 0;
                        }
                        area.m365.groupMode = 'match';
                        groupModeEl.value = 'match';
                        area.m365.groupKind = kind;
                        area.m365.groupId = gid;
                        area.m365.groupLabel =
                            (created && created.displayName) || displayName;
                        area.m365.groupMail = normEmail(
                            (created && created.mail) || nick
                        );
                        refreshM365Sums();
                        const bits = [schulOrgGroupKindLabel(kind) + ' angelegt und verknüpft'];
                        if (memberEmails.length) {
                            bits.push(
                                memberOk +
                                    '/' +
                                    memberEmails.length +
                                    ' Mitglied(er)'
                            );
                        }
                        if (ownerEmail) {
                            bits.push(
                                ownerOk
                                    ? 'Besitzer: ' + ownerEmail
                                    : 'Besitzer konnte nicht gesetzt werden'
                            );
                        }
                        toast(bits.join(' · '), ownerOk || !ownerEmail ? 'success' : 'warning');
                    } catch (e) {
                        toast(
                            'Anlegen fehlgeschlagen: ' + (e && e.message ? e.message : e),
                            'warning'
                        );
                    } finally {
                        createBtn.disabled = false;
                    }
                })();
            });
        }

        mailboxModeEl.addEventListener('change', function () {
            const mode = mailboxModeEl.value || 'none';
            area.m365.mailboxMode = mode === 'create' || mode === 'match' ? mode : 'none';
            if (mode === 'none') {
                area.m365.mailboxUpn = null;
                area.m365.mailboxLabel = null;
                area.m365.mailboxId = null;
                if (mailScriptEl) {
                    mailScriptEl.value = '';
                    mailScriptEl.hidden = true;
                }
                mailCmdDownloadText = '';
                if (mailScriptCopyBtn) mailScriptCopyBtn.hidden = true;
                if (mailScriptDlBtn) mailScriptDlBtn.hidden = true;
            }
            refreshM365Sums();
        });

        overlay.querySelector('[data-area-pick-mail]').addEventListener('click', function () {
            if (mailboxModeEl.value === 'none') mailboxModeEl.value = 'match';
            void toggleInlineEntraUserPicker(
                inlineMailHost,
                inlineMailState,
                {
                    hint:
                        'Postfach für „' +
                        area.label +
                        '“ suchen (freigegebene Postfächer oft als Benutzer) und antippen.',
                    onPick: function (picked) {
                        area.m365.mailboxMode = 'match';
                        mailboxModeEl.value = 'match';
                        area.m365.mailboxUpn = normEmail(
                            picked.mail || picked.userPrincipalName || ''
                        );
                        area.m365.mailboxLabel = picked.displayName || '';
                        area.m365.mailboxId = picked.id || null;
                        refreshM365Sums();
                        toast('Postfach zugeordnet.', 'success');
                    }
                },
                inlinePickers
            ).catch(function (e) {
                toast('Postfachauswahl fehlgeschlagen: ' + (e && e.message ? e.message : e), 'warning');
            });
        });
        overlay.querySelector('[data-area-clear-mail]').addEventListener('click', function () {
            mailboxModeEl.value = 'none';
            area.m365.mailboxMode = 'none';
            area.m365.mailboxUpn = null;
            area.m365.mailboxLabel = null;
            area.m365.mailboxId = null;
            if (mailScriptEl) {
                mailScriptEl.value = '';
                mailScriptEl.hidden = true;
            }
            mailCmdDownloadText = '';
            if (mailScriptCopyBtn) mailScriptCopyBtn.hidden = true;
            if (mailScriptDlBtn) mailScriptDlBtn.hidden = true;
            refreshM365Sums();
        });

        const createMailBtn = overlay.querySelector('[data-area-create-mail]');
        if (createMailBtn) {
            createMailBtn.addEventListener('click', function () {
                void import('../../shared/exchange-shared-mailbox-script.js')
                    .then(function (mod) {
                        const displayName =
                            normStr(mailCreateNameEl && mailCreateNameEl.value) ||
                            normStr(area.label) ||
                            'Postfach';
                        let alias = suggestSchulOrgMailNickname(
                            normStr(mailCreateAliasEl && mailCreateAliasEl.value) || displayName
                        );
                        if (mailCreateAliasEl) mailCreateAliasEl.value = alias;
                        let smtp = normEmail(mailCreateSmtpEl && mailCreateSmtpEl.value);
                        if (!smtp) {
                            const sug = mod.suggestSharedMailboxIdentity(displayName, schoolDomain());
                            smtp = sug.primarySmtp;
                            if (mailCreateSmtpEl && smtp) mailCreateSmtpEl.value = smtp;
                        }
                        if (!smtp || smtp.indexOf('@') === -1) {
                            toast(
                                'Bitte eine gültige E-Mail angeben (z. B. alias@' +
                                    (schoolDomain() || 'schule.at') +
                                    ').',
                                'warning'
                            );
                            return;
                        }
                        const delegates = areaM365People();
                        const mailOpts = {
                            displayName: displayName,
                            alias: alias,
                            primarySmtp: smtp,
                            description: 'Schul-Organigramm: ' + displayName,
                            delegates: delegates.map(function (p) {
                                return p.email;
                            }),
                            rights: ['FullAccess', 'SendAs'],
                            autoMapping: true,
                            suggestedAdminUpn: suggestedExchangeAdminUpn()
                        };
                        const script = mod.buildSharedMailboxCreateScript(mailOpts);
                        mailCmdDownloadText =
                            typeof mod.buildSharedMailboxCreateCmd === 'function'
                                ? mod.buildSharedMailboxCreateCmd(mailOpts)
                                : script;
                        area.m365.mailboxMode = 'create';
                        mailboxModeEl.value = 'create';
                        area.m365.mailboxUpn = smtp;
                        area.m365.mailboxLabel = displayName;
                        area.m365.mailboxId = null;
                        refreshMailDelegatesHint();
                        if (mailScriptEl) {
                            mailScriptEl.value = script;
                            mailScriptEl.hidden = false;
                        }
                        if (mailScriptCopyBtn) mailScriptCopyBtn.hidden = false;
                        if (mailScriptDlBtn) mailScriptDlBtn.hidden = false;
                        refreshM365Sums();
                        toast(
                            delegates.length
                                ? 'Postfach vorgemerkt. .cmd herunterladen und per Doppelklick starten – inkl. Rechte für ' +
                                      delegates.length +
                                      (delegates.length === 1 ? ' Person.' : ' Personen.')
                                : 'Postfach vorgemerkt. .cmd herunterladen und per Doppelklick starten.',
                            'success'
                        );
                    })
                    .catch(function (e) {
                        toast(
                            'Assistent fehlgeschlagen: ' + (e && e.message ? e.message : e),
                            'warning'
                        );
                    });
            });
        }
        if (mailScriptCopyBtn) {
            mailScriptCopyBtn.addEventListener('click', function () {
                const text = mailScriptEl ? mailScriptEl.value : '';
                if (!text) return;
                void navigator.clipboard.writeText(text).then(
                    function () {
                        toast('Geführten PowerShell-Assistenten kopiert.', 'success');
                    },
                    function () {
                        if (mailScriptEl) {
                            mailScriptEl.focus();
                            mailScriptEl.select();
                        }
                        toast('Kopieren manuell: Text markieren und Strg+C.', 'info');
                    }
                );
            });
        }
        if (mailScriptDlBtn) {
            mailScriptDlBtn.addEventListener('click', function () {
                const text = mailCmdDownloadText || (mailScriptEl ? mailScriptEl.value : '');
                if (!text) return;
                const alias =
                    suggestSchulOrgMailNickname(
                        (mailCreateAliasEl && mailCreateAliasEl.value) || area.label || 'postfach'
                    ) || 'postfach';
                // UTF-8 BOM: Windows PowerShell 5.1 liest -Encoding UTF8 sonst falsch
                const blob = new Blob(['\uFEFF' + text], {
                    type: 'application/x-bat;charset=utf-8'
                });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'shared-mailbox-' + alias + '.cmd';
                document.body.appendChild(a);
                a.click();
                a.remove();
                URL.revokeObjectURL(url);
                toast('.cmd gespeichert – per Doppelklick starten und Admin-Konto wählen.', 'success');
            });
        }

        overlay.querySelector('[data-schul-org-save]').addEventListener('click', function () {
            const cur = loadSettings();
            const labelEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgAreaLabel'));
            const nextLabel = normStr(labelEl && labelEl.value) || area.label;
            let nextLayout = normalizeSchulOrganigramm(cur.schulOrganigramm);
            const node = nextLayout.nodes.find(function (n) {
                return n.id === area.id;
            });
            if (node) node.label = nextLabel;
            area.m365.groupMode = groupModeEl.value || 'none';
            area.m365.groupKind = normalizeSchulOrgGroupKind(groupKindEl.value);
            area.m365.mailboxMode = mailboxModeEl.value || 'none';
            if (area.m365.groupMode === 'none') {
                area.m365.groupId = null;
                area.m365.groupLabel = null;
                area.m365.groupMail = null;
            }
            if (area.m365.groupMode === 'create' && !area.m365.groupId) {
                toast('Bitte die Gruppe zuerst anlegen oder „Bestehende zuordnen“ wählen.', 'warning');
                return;
            }
            if (area.m365.groupMode === 'create' && area.m365.groupId) {
                area.m365.groupMode = 'match';
            }
            if (area.m365.mailboxMode === 'none') {
                area.m365.mailboxUpn = null;
                area.m365.mailboxLabel = null;
                area.m365.mailboxId = null;
            }
            if (area.m365.mailboxMode === 'create' && !area.m365.mailboxUpn) {
                toast(
                    'Bitte zuerst „Assistent erzeugen & vormerken“ ausführen oder ein bestehendes Postfach zuordnen.',
                    'warning'
                );
                return;
            }
            // Nach erfolgreichem Exchange-Anlegen bleibt create bis zum Match – UPN/Label behalten
            nextLayout = setSchulOrganigrammAreaM365(nextLayout, area.id, area.m365);
            // Audience-Label spiegeln
            const aid = audienceIdForArea(area);
            if (aid && Array.isArray(cur.verwaltungAudienceGroups)) {
                cur.verwaltungAudienceGroups = normalizeVerwaltungAudienceGroups(
                    cur.verwaltungAudienceGroups
                ).map(function (g) {
                    if (g.id === aid) return Object.assign({}, g, { label: nextLabel });
                    return g;
                });
            }
            cur.schulOrganigramm = nextLayout;
            try {
                window.__ms365SchulOrgSkipSyncOnce = true;
            } catch {
                /* ignore */
            }
            saveSettings(cur);
            const unlinked =
                area.m365.groupMode === 'none' &&
                !(area.m365.mailboxMode === 'match' || area.m365.mailboxMode === 'create');
            toast(
                unlinked
                    ? 'Bereich gespeichert – M365-Gruppe ist gelöst.'
                    : 'Bereich gespeichert.',
                'success'
            );
            closeOverlay();
            notify();
        });
    }

    return {
        openRoleEditor: openRoleEditor,
        openAreaEditor: openAreaEditor,
        createRoleUnderSection: createRoleUnderSection,
        rememberRoleParent: rememberSchulOrgRoleParent,
        getLastRoleParent: getLastSchulOrgRoleParent,
        createAudienceSection: createAudienceSection,
        createAreaAt: createAreaAt,
        openAddCanvasPerson: openAddCanvasPerson,
        openEditCanvasPerson: openEditCanvasPerson,
        pickAndAddPersonToArea: pickAndAddPersonToArea,
        removeAudienceSection: removeAudienceSection,
        removeArea: removeArea,
        openArgeEditor: openArgeEditor,
        createArgeUnderSection: createArgeUnderSection,
        openClassEditor: openClassEditor,
        createClassUnderSection: createClassUnderSection,
        openCouncilEditor: openCouncilEditor,
        createCouncilUnderSection: createCouncilUnderSection,
        openSgaEditor: openSgaEditor,
        createSgaUnderSection: createSgaUnderSection,
        close: closeOverlay,
        movePersonToNode: movePersonToNode,
        addPersonToArea: function (areaNode, person) {
            const cur = loadSettings();
            const res = addSchulOrgPersonToArea(cur, areaNode, person);
            if (res.ok && res.settings) {
                saveSettings(res.settings);
                notify();
            }
            return res;
        },
        removePersonFromArea: function (areaNode, email) {
            const cur = loadSettings();
            const res = removeSchulOrgPersonFromArea(cur, areaNode, email);
            if (res.ok && res.settings) {
                saveSettings(res.settings);
                notify();
            }
            return res;
        },
        listPeopleForNode: function (node) {
            return listPeopleForSchulOrgNode(loadSettings(), node);
        },
        findAdminRole: findAdminRole,
        peopleForAdminRole: peopleForAdminRole,
        findArge: findArge,
        findTeacherByEmail: findTeacherByEmail,
        findClass: findClass,
        findStudentCouncil: findStudentCouncil,
        findSga: findSga
    };
}

if (typeof window !== 'undefined') {
    window.ms365SchulOrganigrammEdit = { createSchulOrganigrammEditor: createSchulOrganigrammEditor };
}

export default {
    createSchulOrganigrammEditor: createSchulOrganigrammEditor,
    findAdminRole: findAdminRole,
    peopleForAdminRole: peopleForAdminRole,
    studentsForClass: studentsForClass,
    guardiansForClass: guardiansForClass,
    teachersForClassChat: teachersForClassChat,
    teachersForArge: teachersForArge,
    listPeopleForSchulOrgNode: listPeopleForSchulOrgNode,
    listM365PeopleForArea: listM365PeopleForArea,
    listMailboxDelegatePeopleForArea: listMailboxDelegatePeopleForArea,
    previewPeopleForSchulOrgNode: previewPeopleForSchulOrgNode,
    moveSchulOrgPerson: moveSchulOrgPerson,
    addSchulOrgPersonToArea: addSchulOrgPersonToArea,
    removeSchulOrgPersonFromArea: removeSchulOrgPersonFromArea,
    SCHUL_ORG_PEOPLE_PREVIEW_LIMIT: SCHUL_ORG_PEOPLE_PREVIEW_LIMIT,
    findArge: findArge,
    findTeacherByEmail: findTeacherByEmail,
    findClass: findClass,
    findStudentCouncil: findStudentCouncil,
    findSga: findSga,
    sgaScopeLabel: sgaScopeLabel
};
