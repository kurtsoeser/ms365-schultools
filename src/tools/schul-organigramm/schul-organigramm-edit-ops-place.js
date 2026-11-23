/**
 * Place / add / remove / heal ops for Schul-Organigramm edit.
 */
import {
    normStr,
    normEmail,
    normCode,
    isBlankPerson,
    ensureTeacherInList
} from './schul-organigramm-edit-utils.js';
import {
    clearHeadOnRecord,
    setHeadOnRecord,
    findAdminIndex,
    findStudentIndex,
    ensurePersonInAreaSettings
} from './schul-organigramm-edit-ops.js';
import { findArge, findClass, normalizeSgaScope } from './schul-organigramm-edit-find.js';
import { studentMatchesClass } from './schul-organigramm-edit-people.js';
import { normalizeAdminAudienceMemberships } from '../../shared/administration-audience-groups.js';
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
    isAreaNode,
    normalizeSchulOrganigramm,
    upsertSchulOrganigrammAreaPerson,
    removeSchulOrganigrammAreaPerson,
    childrenOf,
    removeClassEntityFromSettings,
    renameClassEntityInSettings,
    removeArgeEntityFromSettings,
    ensureSchulOrganigrammSgaStructure,
    audienceIdForArea
} from './schul-organigramm-logic.js';

export function normalizeSchulOrgAreaPersonRole(raw, areaNode) {
    const r = normStr(raw).toLowerCase();
    if (isClassAreaNode(areaNode) || isClassNode(areaNode)) {
        if (
            r === 'kv' ||
            r === 'klassenvorstand' ||
            r === 'head' ||
            r === 'leitung' ||
            r === 'classhead'
        ) {
            return 'kv';
        }
        if (r.indexOf('fn:') === 0) {
            const id = r.slice(3).replace(/[^a-z0-9_-]+/g, '');
            if (id) return 'fn:' + id;
        }
        const knownFn = listSchulOrgClassFunctions().find(function (d) {
            return (
                d.id === r ||
                normStr(d.label).toLowerCase() === r ||
                normStr(d.labelPlural).toLowerCase() === r
            );
        });
        if (knownFn) return 'fn:' + knownFn.id;
        return 'student';
    }
    if (isArgeAreaNode(areaNode) || isArgeNode(areaNode)) {
        if (
            r === 'head' ||
            r === 'leitung' ||
            r === 'argehead' ||
            r === 'arge-leitung' ||
            r === 'arge_leitung' ||
            r === 'kv'
        ) {
            return 'argeHead';
        }
        if (r.indexOf('fn:') === 0) {
            const id = r.slice(3).replace(/[^a-z0-9_-]+/g, '');
            if (id) return 'fn:' + id;
        }
        const knownFn = listSchulOrgArgeFunctions().find(function (d) {
            return (
                d.id === r ||
                normStr(d.label).toLowerCase() === r ||
                normStr(d.labelPlural).toLowerCase() === r
            );
        });
        if (knownFn) return 'fn:' + knownFn.id;
        return 'argeMember';
    }
    return '';
}

/**
 * @param {object} cls
 * @param {string} functionId
 * @param {{ name?: string, email?: string }|null} person
 */
export function setClassFunctionOnRecord(cls, functionId, person) {
    if (!cls || typeof cls !== 'object') return;
    const id = normStr(functionId)
        .toLowerCase()
        .replace(/[^a-z0-9_-]+/g, '');
    if (!id) return;
    const seats = normalizeClassFunctionsMap(cls.functions);
    if (!person) {
        delete seats[id];
    } else {
        const email = normEmail(person.email);
        const name = normStr(person.name);
        if (!email && !name) delete seats[id];
        else seats[id] = { name: name || email, email: email };
    }
    if (Object.keys(seats).length) cls.functions = seats;
    else delete cls.functions;
}

/**
 * @param {object} cls
 * @param {string} email
 * @returns {string} gelöschte Funktions-ID oder ''
 */
export function clearClassFunctionSeatByEmail(cls, email) {
    const em = normEmail(email);
    if (!cls || !em) return '';
    const seats = normalizeClassFunctionsMap(cls.functions);
    let cleared = '';
    Object.keys(seats).forEach(function (id) {
        if (normEmail(seats[id].email) === em) {
            delete seats[id];
            cleared = id;
        }
    });
    if (Object.keys(seats).length) cls.functions = seats;
    else delete cls.functions;
    return cleared;
}

/**
 * headName aus Lehrerliste nachziehen, wenn headEmail gesetzt ist.
 * @param {object} settings
 * @returns {{ settings: object, changed: number }}
 */

export {
    ensureClassHeadsNamedFromTeachers,
    healClassHeadsFromAreaPeople
} from './schul-organigramm-edit-ops-heal.js';

export function placePersonOnArea(settings, areaNode, person) {
    if (!areaAcceptsDirectPeople(areaNode)) {
        return { ok: false, message: 'Bereich nimmt keine Personen auf.' };
    }
    if (isSammelgruppeArea(areaNode)) {
        const sk = sammelgruppeKeyForArea(areaNode);
        if (sk === 'lehrer') {
            const email = normEmail(person && person.email);
            const name = normStr(person && person.name);
            if (!email) return { ok: false, message: 'E-Mail fehlt.' };
            ensureTeacherInList(settings, { name: name || email, email: email });
            return { ok: true, role: 'teacher' };
        }
        return {
            ok: false,
            message:
                sk === 'schueler'
                    ? 'Schüler:innen bitte über die Klassen bzw. Schülerliste in den Stammdaten pflegen.'
                    : 'Personen bitte über die zugehörigen Unterbereiche bzw. Stammdatenlisten pflegen.'
        };
    }
    if (isSgaBoardArea(areaNode)) {
        return {
            ok: false,
            message: 'Personen bitte in Schüler, Lehrer oder Eltern (unter SGA) hinzufügen.'
        };
    }
    if (isFunctionAggregateArea(areaNode)) {
        const agg = functionAggregateDefForArea(areaNode);
        return {
            ok: false,
            message:
                ((agg && agg.labelPlural) || 'Funktions-Sammelbereiche') +
                ' ergeben sich aus den Gruppen-Sitze (Klasse/ARGE). Hier nur Microsoft 365 zuordnen.'
        };
    }
    const email = normEmail(person && person.email);
    const name = normStr(person && person.name);
    if (!email) return { ok: false, message: 'E-Mail fehlt.' };

    // Klasse: Funktion Klassenvorstand, Klassenfunktion oder Schüler:in → Stammdaten
    if ((isClassAreaNode(areaNode) || isClassNode(areaNode)) && areaNode.refKey) {
        const role = normalizeSchulOrgAreaPersonRole(person && person.role, areaNode);
        if (!Array.isArray(settings.classes)) settings.classes = [];
        const cls = findClass(settings, areaNode.refKey);
        if (!cls) return { ok: false, message: 'Klasse in den Stammdaten nicht gefunden.' };
        if (role === 'kv') {
            setHeadOnRecord(cls, { name: name || email, email: email });
            ensureTeacherInList(settings, { name: name || email, email: email });
            // Falls fälschlich als Schüler geführt: Klassen-Zuordnung lösen
            if (Array.isArray(settings.students)) {
                settings.students = settings.students.map(function (st) {
                    if (normEmail(st && st.email) !== email) return st;
                    if (!studentMatchesClass(cls, st)) return st;
                    return Object.assign({}, st, { klasse: '' });
                });
            }
            return { ok: true, role: 'kv' };
        }
        if (!Array.isArray(settings.students)) settings.students = [];
        const klasseCode = normStr(cls.code) || normStr(cls.name);
        const ensureStudent = function () {
            const idx = findStudentIndex(settings.students, { email: email, name: name }, klasseCode);
            if (idx < 0) {
                settings.students.push({
                    name: name || email,
                    email: email,
                    klasse: klasseCode
                });
            } else {
                settings.students[idx] = Object.assign({}, settings.students[idx], {
                    name: name || settings.students[idx].name || email,
                    email: email,
                    klasse: klasseCode
                });
            }
        };
        if (role.indexOf('fn:') === 0) {
            const fnId = role.slice(3);
            const def = listSchulOrgClassFunctions(settings).find(function (d) {
                return d.id === fnId;
            });
            if (!def) return { ok: false, message: 'Unbekannte Klassenfunktion.' };
            ensureStudent();
            setClassFunctionOnRecord(cls, fnId, { name: name || email, email: email });
            return { ok: true, role: role };
        }
        ensureStudent();
        return { ok: true, role: 'student' };
    }

    // ARGE: Leitung, Funktion oder Mitglied
    if ((isArgeAreaNode(areaNode) || isArgeNode(areaNode)) && areaNode.refKey) {
        const role = normalizeSchulOrgAreaPersonRole(person && person.role, areaNode);
        if (!Array.isArray(settings.arges)) settings.arges = [];
        const arge = findArge(settings, areaNode.refKey);
        if (!arge) return { ok: false, message: 'ARGE in den Stammdaten nicht gefunden.' };
        if (role === 'argeHead') {
            setHeadOnRecord(arge, { name: name || email, email: email });
            ensureTeacherInList(settings, { name: name || email, email: email });
            return { ok: true, role: 'argeHead' };
        }
        if (role.indexOf('fn:') === 0) {
            const fnId = role.slice(3);
            const def = listSchulOrgArgeFunctions(settings).find(function (d) {
                return d.id === fnId;
            });
            if (!def) return { ok: false, message: 'Unbekannte ARGE-Funktion.' };
            ensureTeacherInList(settings, { name: name || email, email: email });
            setClassFunctionOnRecord(arge, fnId, { name: name || email, email: email });
            return { ok: true, role: role };
        }
        settings.schulOrganigramm = upsertSchulOrganigrammAreaPerson(
            settings.schulOrganigramm,
            areaNode.id,
            { email: email, name: name }
        );
        return { ok: true, role: 'argeMember' };
    }

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
export function removePersonFromAreaSettings(settings, areaId, email) {
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
 * Bei Klasse/ARGE: person.role = kv|student|argeHead|argeMember (bzw. Aliase).
 * @param {object} settings
 * @param {object} areaNode
 * @param {{ email: string, name?: string, role?: string }} person
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
    next.classes = Array.isArray(settings.classes)
        ? settings.classes.map(function (c) {
              return Object.assign({}, c);
          })
        : [];
    next.students = Array.isArray(settings.students)
        ? settings.students.map(function (st) {
              return Object.assign({}, st);
          })
        : [];
    next.teachers = Array.isArray(settings.teachers)
        ? settings.teachers.map(function (t) {
              return Object.assign({}, t);
          })
        : [];
    next.arges = Array.isArray(settings.arges)
        ? settings.arges.map(function (a) {
              return Object.assign({}, a, {
                  subjects: Array.isArray(a.subjects) ? a.subjects.slice() : []
              });
          })
        : [];
    next.schulOrganigramm = normalizeSchulOrganigramm(settings.schulOrganigramm);
    const placed = placePersonOnArea(next, areaNode, person);
    if (!placed.ok) return placed;
    let message = 'Person hinzugefügt.';
    if (placed.role === 'kv') message = 'Als Klassenvorstand gesetzt.';
    else if (placed.role === 'student') message = 'Als Schüler:in hinzugefügt.';
    else if (placed.role === 'argeHead') message = 'Als ARGE-Leitung gesetzt.';
    else if (placed.role === 'argeMember') message = 'Als ARGE-Mitglied hinzugefügt.';
    else if (placed.role === 'teacher') message = 'Zur Lehrerliste hinzugefügt.';
    else if (placed.role && String(placed.role).indexOf('fn:') === 0) {
        const fnId = String(placed.role).slice(3);
        const def = listSchulOrgClassFunctions(next).find(function (d) {
            return d.id === fnId;
        });
        message = 'Als ' + ((def && def.label) || 'Klassenfunktion') + ' gesetzt.';
    }
    return { ok: true, settings: next, message: message, role: placed.role || '' };
}

/**
 * @param {object} settings
 * @param {object} areaNode
 * @param {string} email
 */
export function removeSchulOrgPersonFromArea(settings, areaNode, email) {
    if (!settings || !areaNode) return { ok: false, message: 'Bereich fehlt.' };
    const em = normEmail(email);
    if (!em) return { ok: false, message: 'E-Mail fehlt.' };
    const next = Object.assign({}, settings);
    next.adminAudienceMemberships = normalizeAdminAudienceMemberships(
        settings.adminAudienceMemberships
    ).map(function (m) {
        return Object.assign({}, m);
    });
    next.classes = Array.isArray(settings.classes)
        ? settings.classes.map(function (c) {
              return Object.assign({}, c);
          })
        : [];
    next.students = Array.isArray(settings.students)
        ? settings.students.map(function (st) {
              return Object.assign({}, st);
          })
        : [];
    next.teachers = Array.isArray(settings.teachers)
        ? settings.teachers.map(function (t) {
              return Object.assign({}, t);
          })
        : [];
    next.arges = Array.isArray(settings.arges)
        ? settings.arges.map(function (a) {
              return Object.assign({}, a, {
                  subjects: Array.isArray(a.subjects) ? a.subjects.slice() : []
              });
          })
        : [];
    next.schulOrganigramm = normalizeSchulOrganigramm(settings.schulOrganigramm);

    if (sammelgruppeKeyForArea(areaNode) === 'lehrer') {
        next.teachers = next.teachers.filter(function (t) {
            return normEmail(t && t.email) !== em;
        });
        return { ok: true, settings: next, message: 'Lehrkraft aus der Liste entfernt.' };
    }

    if ((isClassAreaNode(areaNode) || isClassNode(areaNode)) && areaNode.refKey) {
        const cls = findClass(next, areaNode.refKey);
        if (cls && normEmail(cls.headEmail) === em) {
            clearHeadOnRecord(cls);
            removePersonFromAreaSettings(next, areaNode.id, em);
            return { ok: true, settings: next, message: 'Klassenvorstand entfernt.' };
        }
        if (cls) {
            const clearedFn = clearClassFunctionSeatByEmail(cls, em);
            if (clearedFn) {
                removePersonFromAreaSettings(next, areaNode.id, em);
                const def = listSchulOrgClassFunctions(next).find(function (d) {
                    return d.id === clearedFn;
                });
                return {
                    ok: true,
                    settings: next,
                    message: (def && def.label ? def.label : 'Klassenfunktion') + ' entfernt.'
                };
            }
            next.students = next.students.map(function (st) {
                if (normEmail(st && st.email) !== em) return st;
                if (!studentMatchesClass(cls, st)) return st;
                return Object.assign({}, st, { klasse: '' });
            });
        }
        removePersonFromAreaSettings(next, areaNode.id, em);
        return { ok: true, settings: next, message: 'Schüler:in entfernt.' };
    }

    if ((isArgeAreaNode(areaNode) || isArgeNode(areaNode)) && areaNode.refKey) {
        const arge = findArge(next, areaNode.refKey);
        if (arge && normEmail(arge.headEmail) === em) {
            clearHeadOnRecord(arge);
            removePersonFromAreaSettings(next, areaNode.id, em);
            return { ok: true, settings: next, message: 'ARGE-Leitung entfernt.' };
        }
        if (arge) {
            const clearedFn = clearClassFunctionSeatByEmail(arge, em);
            if (clearedFn) {
                removePersonFromAreaSettings(next, areaNode.id, em);
                const def = listSchulOrgArgeFunctions(next).find(function (d) {
                    return d.id === clearedFn;
                });
                return {
                    ok: true,
                    settings: next,
                    message: (def && def.label ? def.label : 'ARGE-Funktion') + ' entfernt.'
                };
            }
        }
        removePersonFromAreaSettings(next, areaNode.id, em);
        return { ok: true, settings: next, message: 'Person entfernt.' };
    }

    removePersonFromAreaSettings(next, areaNode.id, em);
    return { ok: true, settings: next, message: 'Person entfernt.' };
}



export function syncSeatMode(role, people) {
    const filled = people.filter(function (p) {
        return !isBlankPerson(p);
    }).length;
    if (filled >= 2) role.seatMode = 'multi';
    else if (!role.seatMode) role.seatMode = 'single';
}

export {
    patchSchulOrganigrammRefLabel,
    removeSchulOrganigrammByRef
} from './schul-organigramm-edit-ops-ref.js';
