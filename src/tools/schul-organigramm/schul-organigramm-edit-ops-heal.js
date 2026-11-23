/**
 * Class-head heal helpers for Schul-Organigramm edit.
 */
import { normStr, normEmail } from './schul-organigramm-edit-utils.js';
import { clearHeadOnRecord, setHeadOnRecord } from './schul-organigramm-edit-ops.js';
import { findClass } from './schul-organigramm-edit-find.js';
import {
    isClassAreaNode,
    normalizeSchulOrganigramm,
    childrenOf,
    NODE,
    removeSchulOrganigrammAreaPerson
} from './schul-organigramm-logic.js';

export function ensureClassHeadsNamedFromTeachers(settings) {
    const s = settings && typeof settings === 'object' ? Object.assign({}, settings) : {};
    const classes = Array.isArray(s.classes) ? s.classes.map(function (c) {
        return Object.assign({}, c);
    }) : [];
    /** @type {Map<string, string>} */
    const teacherNameByEmail = new Map();
    (Array.isArray(s.teachers) ? s.teachers : []).forEach(function (t) {
        const em = normEmail(t && t.email);
        if (em) teacherNameByEmail.set(em, normStr(t && t.name));
    });
    let changed = 0;
    classes.forEach(function (c, idx) {
        const em = normEmail(c && c.headEmail);
        if (!em) return;
        const fromTeacher = teacherNameByEmail.get(em) || '';
        const nextName = normStr(c && c.headName) || fromTeacher;
        if (nextName && nextName !== normStr(c && c.headName)) {
            classes[idx] = Object.assign({}, c, { headName: nextName, headEmail: em });
            changed += 1;
        } else if (em !== normStr(c && c.headEmail).toLowerCase()) {
            classes[idx] = Object.assign({}, c, { headEmail: em });
            changed += 1;
        }
    });
    if (!changed) return { settings: s, changed: 0 };
    s.classes = classes;
    return { settings: s, changed: changed };
}

/**
 * Klassen ohne KV: Lehrkraft aus Area-Personen-Slot → Klassenvorstand.
 * @param {object} settings
 * @returns {{ settings: object, changed: number }}
 */
export function healClassHeadsFromAreaPeople(settings) {
    const s = settings && typeof settings === 'object' ? Object.assign({}, settings) : {};
    s.classes = Array.isArray(s.classes)
        ? s.classes.map(function (c) {
              return Object.assign({}, c);
          })
        : [];
    s.schulOrganigramm = normalizeSchulOrganigramm(s.schulOrganigramm);
    /** @type {Set<string>} */
    const teacherEmails = new Set();
    (Array.isArray(s.teachers) ? s.teachers : []).forEach(function (t) {
        const em = normEmail(t && t.email);
        if (em) teacherEmails.add(em);
    });
    let changed = 0;
    (s.schulOrganigramm.nodes || []).forEach(function (area) {
        if (!isClassAreaNode(area) || !area.refKey) return;
        const cls = findClass(s, area.refKey);
        if (!cls || normEmail(cls.headEmail)) return;
        const people = childrenOf(s.schulOrganigramm, area.id).filter(function (ch) {
            return ch && ch.type === NODE.PERSON && normStr(ch.refKind) === 'adminPerson';
        });
        let pick = null;
        for (let i = 0; i < people.length; i += 1) {
            const em = normEmail(people[i].refKey);
            if (em && teacherEmails.has(em)) {
                pick = people[i];
                break;
            }
        }
        if (!pick && people.length === 1) pick = people[0];
        if (!pick) return;
        const em = normEmail(pick.refKey);
        if (!em) return;
        setHeadOnRecord(cls, { name: normStr(pick.label) || em, email: em });
        s.schulOrganigramm = removeSchulOrganigrammAreaPerson(s.schulOrganigramm, em, area.id);
        changed += 1;
    });
    return { settings: s, changed: changed };
}
