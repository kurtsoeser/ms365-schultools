/**
 * Move ops for Schul-Organigramm edit.
 */
import {
    normStr,
    normEmail,
    normCode,
    isBlankPerson,
    ensureTeacherInList
} from './schul-organigramm-edit-utils.js';
import {
    findAdminRole,
    peopleForAdminRole,
    syncAdministrationFromAdmin
} from './schul-organigramm-edit-people.js';
import { findArge, findClass } from './schul-organigramm-edit-find.js';
import { normalizeAdminAudienceMemberships } from '../../shared/administration-audience-groups.js';
import {
    placePersonOnArea,
    removePersonFromAreaSettings,
    syncSeatMode
} from './schul-organigramm-edit-ops-place.js';
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
    normalizeClassFunctionsMap,
    sgaScopeForArea,
    isSgaScopeArea,
    isSgaBoardArea,
    svSeatForNode,
    SCHUL_ORG_SV_SEATS,
    SCHUL_ORG_SGA_SV_AREA_ID,
    SCHUL_ORG_BOARD_REF,
    NODE,
    sammelgruppeKeyForArea,
    areaAcceptsDirectPeople,
    isRoleCapableSection,
    canPlaceSchulOrgNodeInSection,
    moveSchulOrganigrammNode,
    applySchulOrganigrammToStammdaten,
    isAreaNode,
    normalizeSchulOrganigramm,
    upsertSchulOrganigrammAreaPerson,
    removeSchulOrganigrammAreaPerson,
    childrenOf,
    ensureSchulOrganigrammSgaStructure,
    audienceIdForArea
} from './schul-organigramm-logic.js';

export function clearHeadOnRecord(rec) {
    if (!rec) return;
    rec.headName = '';
    rec.headEmail = '';
}

export function setHeadOnRecord(rec, person) {
    if (!rec || !person) return;
    rec.headName = normStr(person.name) || normStr(person.label);
    rec.headEmail = normEmail(person.email);
}

export function findAdminIndex(admin, person) {
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

export function findStudentIndex(students, person, preferKlasse) {
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
 * Rolle im Layout verschieben und alle Rollen-Personen in den neuen Bereich mitnehmen
 * (Audience-Mitgliedschaft / Area-Person-Slot: alt entfernen, neu setzen).
 * @param {object} settings
 * @param {string} roleNodeId
 * @param {string} newParentId
 * @param {string} [beforeNodeId]
 * @returns {{ ok: boolean, message?: string, settings?: object, layout?: object }}
 */
export function moveSchulOrgRole(settings, roleNodeId, newParentId, beforeNodeId) {
    if (!settings || !roleNodeId || !newParentId) {
        return { ok: false, message: 'Verschieben nicht möglich.' };
    }
    const next = Object.assign({}, settings);
    next.admin = Array.isArray(settings.admin)
        ? settings.admin.map(function (p) {
              return Object.assign({}, p);
          })
        : [];
    next.adminRoles = Array.isArray(settings.adminRoles) ? settings.adminRoles.slice() : [];
    next.adminAudienceMemberships = normalizeAdminAudienceMemberships(
        settings.adminAudienceMemberships
    ).map(function (m) {
        return Object.assign({}, m);
    });
    next.schulOrganigramm = normalizeSchulOrganigramm(settings.schulOrganigramm);

    const before = next.schulOrganigramm;
    const roleNode = before.nodes.find(function (n) {
        return n.id === roleNodeId;
    });
    if (!roleNode || (roleNode.type !== NODE.ROLE && roleNode.type !== 'role')) {
        return { ok: false, message: 'Keine Funktion.' };
    }
    const oldParentId = normStr(roleNode.parentId);
    if (oldParentId === normStr(newParentId)) {
        const resorted = moveSchulOrganigrammNode(before, roleNodeId, newParentId, beforeNodeId);
        next.schulOrganigramm = resorted;
        return { ok: true, settings: next, layout: resorted, message: 'Reihenfolge angepasst.' };
    }
    if (!canPlaceSchulOrgNodeInSection(roleNode, newParentId, before)) {
        return { ok: false, message: 'Dieser Bereich nimmt keine Funktionen auf.' };
    }

    const role = findAdminRole(next, roleNode.refKey);
    const people = role
        ? peopleForAdminRole(next, role).filter(function (p) {
              return !isBlankPerson(p);
          })
        : [];

    const movedLayout = moveSchulOrganigrammNode(before, roleNodeId, newParentId, beforeNodeId);
    const afterNode = movedLayout.nodes.find(function (n) {
        return n.id === roleNodeId;
    });
    if (!afterNode || normStr(afterNode.parentId) !== normStr(newParentId)) {
        return { ok: false, message: 'Hier kann diese Funktion nicht abgelegt werden.' };
    }
    afterNode.tone =
        newParentId === 'sec_schulleitung' ||
        (role && normStr(role.tier).toLowerCase() === 'schulleitung')
            ? 'lead'
            : 'staff';
    next.schulOrganigramm = movedLayout;

    const newArea = movedLayout.nodes.find(function (n) {
        return n.id === newParentId && isAreaNode(n);
    });
    people.forEach(function (p) {
        const email = normEmail(p.email);
        if (!email) return;
        if (oldParentId) removePersonFromAreaSettings(next, oldParentId, email);
        if (newArea) ensurePersonInAreaSettings(next, newArea, { email: email, name: p.name });
    });

    const applied = applySchulOrganigrammToStammdaten(next, next.schulOrganigramm);
    return {
        ok: true,
        settings: applied.settings,
        layout: applied.layout,
        message:
            people.length === 1
                ? 'Funktion inkl. Person verschoben.'
                : people.length > 1
                  ? 'Funktion inkl. ' + people.length + ' Personen verschoben.'
                  : 'Funktion verschoben.'
    };
}

/**
 * @param {object} settings
 * @param {object} areaNode
 * @param {{ email?: string, name?: string }} person
 */
export function ensurePersonInAreaSettings(settings, areaNode, person) {
    placePersonOnArea(settings, areaNode, person);
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

export {
    normalizeSchulOrgAreaPersonRole,
    ensureClassHeadsNamedFromTeachers,
    healClassHeadsFromAreaPeople,
    placePersonOnArea,
    removePersonFromAreaSettings,
    addSchulOrgPersonToArea,
    removeSchulOrgPersonFromArea,
    syncSeatMode,
    patchSchulOrganigrammRefLabel,
    removeSchulOrganigrammByRef,
    setClassFunctionOnRecord,
    clearClassFunctionSeatByEmail
} from './schul-organigramm-edit-ops-place.js';
