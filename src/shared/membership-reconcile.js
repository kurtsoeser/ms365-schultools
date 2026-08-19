/**
 * Abgleich lokaler E-Mail-Listen mit Microsoft-365-Gruppenmitgliedern.
 * @file
 */
import { reconcileSammelgruppe, reconcileClassMembers } from './student-class-lifecycle.js';
import {
    applyStudentImportSelection,
    applyTeacherImportSelection,
    buildStudentImportPreview,
    buildTeacherImportPreview,
    suggestKlasseFromUser,
    summarizeUserLicenses,
    teacherEmailOfUser
} from './graph-licenses.js';
import { normStr } from './utils/strings.js';

export function normMemberEmail(v) {
    return String(v ?? '')
        .trim()
        .toLowerCase();
}

export function normEmailList(arr) {
    const out = [];
    const seen = new Set();
    (Array.isArray(arr) ? arr : []).forEach(function (em) {
        const n = normMemberEmail(em);
        if (!n || n.indexOf('@') === -1 || seen.has(n)) return;
        seen.add(n);
        out.push(n);
    });
    return out;
}

/**
 * @param {string[]} localEmails
 * @param {string[]} graphEmails
 * @returns {{ onlyLocal: string[], onlyGraph: string[], both: string[] }}
 */
export function diffMemberships(localEmails, graphEmails) {
    const local = normEmailList(localEmails);
    const graph = normEmailList(graphEmails);
    const rec = reconcileSammelgruppe(local, graph);
    const graphSet = new Set(graph);
    const both = local.filter(function (em) {
        return graphSet.has(em);
    });
    const onlyLocal = rec.join.slice().sort();
    const onlyGraph = rec.leave.slice().sort();
    both.sort();
    return { onlyLocal: onlyLocal, onlyGraph: onlyGraph, both: both };
}

export function memberEmailFromGraph(person) {
    if (!person || typeof person !== 'object') return '';
    const em = normMemberEmail(person.mail || person.userPrincipalName);
    return em.indexOf('@') !== -1 ? em : '';
}

export function memberDisplayNameFromGraph(person, email) {
    const dn = String((person && person.displayName) || '').trim();
    if (dn) return dn;
    const em = email || memberEmailFromGraph(person);
    const local = em ? em.split('@')[0] : '';
    return local || em || '';
}

/**
 * @param {string} email
 * @param {string[]} existingCodes
 * @returns {string}
 */
export function suggestTeacherCode(email, existingCodes) {
    const local = String(email || '').split('@')[0] || '';
    let code = local.replace(/[^a-zA-Z0-9]/g, '').toUpperCase().slice(0, 12);
    if (!code) code = 'T';
    const seen = new Set(
        (Array.isArray(existingCodes) ? existingCodes : []).map(function (c) {
            return String(c || '')
                .trim()
                .toUpperCase();
        })
    );
    const base = code;
    let n = 2;
    while (seen.has(code)) {
        const suffix = String(n);
        code = (base.slice(0, Math.max(1, 12 - suffix.length)) + suffix).slice(0, 12);
        n += 1;
        if (n > 9999) break;
    }
    return code;
}

/**
 * @param {object} person Graph-Mitglied
 * @param {Array<{ code?: string }>} existingTeachers
 */
export function buildTeacherImportRow(person, existingTeachers) {
    const email = memberEmailFromGraph(person);
    const codes = (Array.isArray(existingTeachers) ? existingTeachers : []).map(function (t) {
        return t && t.code;
    });
    return {
        code: suggestTeacherCode(email, codes),
        name: memberDisplayNameFromGraph(person, email),
        email: email
    };
}

/**
 * @param {object} person Graph-Mitglied
 * @param {string} [defaultClass]
 */
export function buildStudentImportRow(person, defaultClass) {
    const email = memberEmailFromGraph(person);
    return {
        klasse: String(defaultClass || '').trim(),
        name: memberDisplayNameFromGraph(person, email),
        email: email
    };
}

/**
 * @param {object[]} items Graph-Mitglieder
 * @returns {Map<string, object>}
 */
export function indexGraphMembersByEmail(items) {
    const map = new Map();
    (Array.isArray(items) ? items : []).forEach(function (p) {
        const em = memberEmailFromGraph(p);
        if (em && !map.has(em)) map.set(em, p);
    });
    return map;
}

function compareDeName(a, b) {
    try {
        return String(a || '').localeCompare(String(b || ''), 'de', { sensitivity: 'base' });
    } catch {
        return String(a || '').localeCompare(String(b || ''));
    }
}

/**
 * Import-Vorschau für Gruppenmitglieder: Education-Lizenz + Fallback ohne passende Lizenz.
 * @param {'lehrer'|'schueler'} kind
 * @param {object[]} users Graph-User mit assignedLicenses
 * @param {object[]} existingRows bestehende Stammdaten
 * @param {Map|null} skuLookup
 * @param {{ activeOnly?: boolean, families?: string[] }} [opts]
 */
export function buildMembershipImportPreview(kind, users, existingRows, skuLookup, opts) {
    const opt = opts && typeof opts === 'object' ? opts : {};
    const previewOpts = {
        activeOnly: opt.activeOnly !== false,
        guests: false,
        families: Array.isArray(opt.families) ? opt.families : ['a1', 'a3', 'a5']
    };
    const isTeacher = kind === 'lehrer';
    const licensed = isTeacher
        ? buildTeacherImportPreview(users, existingRows, skuLookup, previewOpts)
        : buildStudentImportPreview(users, existingRows, skuLookup, previewOpts);
    const byEmail = new Map();
    licensed.forEach(function (row) {
        if (row && row.email) byEmail.set(normMemberEmail(row.email), row);
    });

    (Array.isArray(users) ? users : []).forEach(function (u) {
        const email = teacherEmailOfUser(u);
        if (!email || byEmail.has(email)) return;
        const sum = summarizeUserLicenses(u, skuLookup);
        const expected = isTeacher ? sum.hasFacultyUserPlan : sum.hasStudentUserPlan;
        if (expected) return;
        const existing = (Array.isArray(existingRows) ? existingRows : []).find(function (r) {
            return normMemberEmail(r && r.email) === email;
        });
        const row = isTeacher
            ? buildTeacherImportRow(u, existingRows)
            : buildStudentImportRow(u, suggestKlasseFromUser(u));
        byEmail.set(email, {
            graphUserId: String(u.id || ''),
            displayName: normStr(u.displayName),
            givenName: normStr(u.givenName),
            surname: normStr(u.surname),
            userPrincipalName: normStr(u.userPrincipalName),
            accountEnabled: u.accountEnabled !== false,
            userType: String(u.userType || 'Member'),
            email: email,
            code: row.code,
            klasse: row.klasse,
            name: row.name,
            licenseLabel: sum.primaryLabel || 'Keine passende Lizenz',
            licenseWarning: true,
            warningText: isTeacher
                ? 'Kein A1/A3/A5 für Lehrpersonal – nur manuell übernehmen'
                : 'Kein A1/A3/A5 für Schüler:innen – nur manuell übernehmen',
            alreadyInList: !!existing,
            selected: false
        });
    });

    return Array.from(byEmail.values()).sort(function (a, b) {
        return compareDeName(a.name, b.name);
    });
}

/**
 * @param {'lehrer'|'schueler'} kind
 * @param {object[]} existingRows
 * @param {object[]} previewRows
 */
export function applyMembershipImportSelection(kind, existingRows, previewRows) {
    if (kind === 'lehrer') {
        return applyTeacherImportSelection(existingRows, previewRows);
    }
    return applyStudentImportSelection(existingRows, previewRows);
}

/**
 * Klassengruppen-Abgleich: Schüler der Klasse vs. Gruppe; Lehrkräfte/andere bleiben erhalten.
 */
export function diffClassMemberships(classEmails, allStudentEmails, graphEmails) {
    const rec = reconcileClassMembers(classEmails, allStudentEmails, graphEmails);
    const local = normEmailList(classEmails);
    const graph = normEmailList(graphEmails);
    const localSet = new Set(local);
    const leaveSet = new Set(rec.leave);
    const both = local.filter(function (em) {
        return new Set(graph).has(em);
    });
    const preserved = graph.filter(function (em) {
        return !localSet.has(em) && !leaveSet.has(em);
    });
    return {
        onlyLocal: rec.join.slice().sort(),
        onlyGraph: rec.leave.slice().sort(),
        both: both.sort(),
        preserved: preserved.sort()
    };
}

/**
 * Verwaltungskontakte aus Import-Vorschau in admin-Liste übernehmen.
 */
export function applyAdminImportSelection(existingRows, previewRows) {
    const out = (Array.isArray(existingRows) ? existingRows : []).map(function (r) {
        return {
            role: normStr(r.role),
            name: normStr(r.name),
            email: normMemberEmail(r.email)
        };
    });
    const emailIndex = new Map();
    out.forEach(function (r, i) {
        if (r.email) emailIndex.set(r.email, i);
    });
    const added = [];
    const updated = [];
    const skipped = [];
    const directoryMatches = {};
    const iso = new Date().toISOString();

    (Array.isArray(previewRows) ? previewRows : []).forEach(function (row) {
        if (!row || !row.selected) return;
        const email = normMemberEmail(row.email);
        const name = normStr(row.name || row.displayName);
        const role = normStr(row.role);
        if (!email || email.indexOf('@') === -1) {
            skipped.push(row);
            return;
        }
        if (row.graphUserId) {
            directoryMatches[email] = {
                graphUserId: String(row.graphUserId),
                displayName: name,
                userPrincipalName: normStr(row.userPrincipalName),
                notFound: false,
                checkedAt: iso
            };
        }
        if (emailIndex.has(email)) {
            const i = emailIndex.get(email);
            const prev = out[i];
            const next = {
                role: role || prev.role,
                name: name || prev.name,
                email: email
            };
            if (next.role !== prev.role || next.name !== prev.name) {
                out[i] = next;
                updated.push(next);
            } else {
                skipped.push(row);
            }
            return;
        }
        const next = { role: role, name: name, email: email };
        emailIndex.set(email, out.length);
        out.push(next);
        added.push(next);
    });

    return { admin: out, added: added, updated: updated, skipped: skipped, directoryMatches: directoryMatches };
}

const api = {
    normMemberEmail: normMemberEmail,
    normEmailList: normEmailList,
    diffMemberships: diffMemberships,
    memberEmailFromGraph: memberEmailFromGraph,
    memberDisplayNameFromGraph: memberDisplayNameFromGraph,
    suggestTeacherCode: suggestTeacherCode,
    buildTeacherImportRow: buildTeacherImportRow,
    buildStudentImportRow: buildStudentImportRow,
    indexGraphMembersByEmail: indexGraphMembersByEmail,
    buildMembershipImportPreview: buildMembershipImportPreview,
    applyMembershipImportSelection: applyMembershipImportSelection,
    diffClassMemberships: diffClassMemberships,
    applyAdminImportSelection: applyAdminImportSelection
};

if (typeof window !== 'undefined') {
    window.ms365MembershipReconcile = api;
}

export default api;
