/**
 * Schüler-Klassenwechsel: Diff der Stammliste und Vorschau der Graph-Mitgliedschaften.
 * Kein DOM, kein Graph.
 */

function normStr(v) {
    return String(v == null ? '' : v).trim();
}

function normEmail(v) {
    return normStr(v).toLowerCase();
}

function normClass(v) {
    return normStr(v)
        .toLowerCase()
        .replace(/\s+/g, '');
}

export function studentKey(row) {
    const em = normEmail(row && row.email);
    if (em && em.indexOf('@') !== -1) return 'e:' + em;
    const name = normStr(row && row.name).toLowerCase();
    const klasse = normClass(row && row.klasse);
    return 'n:' + name + '|' + klasse;
}

function cloneStudent(row) {
    return {
        id: row && row.id != null ? row.id : '',
        klasse: normStr(row && row.klasse),
        name: normStr(row && row.name),
        email: normEmail(row && row.email)
    };
}

/**
 * @param {unknown} prev
 * @param {unknown} next
 * @returns {{ added: object[], removed: object[], classChanged: { student: object, fromClass: string, toClass: string }[] }}
 */
export function diffStudents(prev, next) {
    const prevArr = Array.isArray(prev) ? prev : [];
    const nextArr = Array.isArray(next) ? next : [];
    const prevMap = new Map();
    prevArr.forEach(function (s) {
        prevMap.set(studentKey(s), s);
    });
    const nextMap = new Map();
    nextArr.forEach(function (s) {
        nextMap.set(studentKey(s), s);
    });

    const added = [];
    const removed = [];
    const classChanged = [];

    nextMap.forEach(function (s, k) {
        if (!prevMap.has(k)) {
            added.push(cloneStudent(s));
            return;
        }
        const old = prevMap.get(k);
        if (normClass(old && old.klasse) !== normClass(s && s.klasse)) {
            classChanged.push({
                student: cloneStudent(s),
                fromClass: normStr(old && old.klasse),
                toClass: normStr(s && s.klasse)
            });
        }
    });
    prevMap.forEach(function (s, k) {
        if (!nextMap.has(k)) removed.push(cloneStudent(s));
    });

    return { added: added, removed: removed, classChanged: classChanged };
}

function findClassTeam(classTeams, klasse) {
    const want = normClass(klasse);
    if (!want) return null;
    const list = Array.isArray(classTeams) ? classTeams : [];
    for (let i = 0; i < list.length; i++) {
        const t = list[i];
        const code = normClass(t && (t.classCode || t.code));
        if (code && code === want) return t;
    }
    return null;
}

function groupIdOf(team) {
    return team && team.graphGroupId ? String(team.graphGroupId).trim() : '';
}

function ensureGroup(map, groupId, label) {
    const id = String(groupId || '').trim();
    if (!id) return null;
    if (!map.has(id)) {
        map.set(id, { groupId: id, label: label || id, join: [], leave: [] });
    }
    return map.get(id);
}

function pushUnique(arr, email) {
    const em = normEmail(email);
    if (!em || em.indexOf('@') === -1) return;
    if (arr.indexOf(em) === -1) arr.push(em);
}

/**
 * @param {{ added?: object[], removed?: object[], classChanged?: { student: object, fromClass: string, toClass: string }[] }} diff
 * @param {object[]} classTeams
 * @param {string} [schuelerGroupId]
 */
export function previewMemberships(diff, classTeams, schuelerGroupId) {
    const d = diff && typeof diff === 'object' ? diff : {};
    const map = new Map();
    const sammelId = String(schuelerGroupId || '').trim();

    (Array.isArray(d.added) ? d.added : []).forEach(function (s) {
        const em = s && s.email;
        const team = findClassTeam(classTeams, s && s.klasse);
        const gid = groupIdOf(team);
        const g = ensureGroup(map, gid, (s && s.klasse) || 'Klasse');
        if (g) pushUnique(g.join, em);
        const sg = ensureGroup(map, sammelId, 'Alle Schülerinnen');
        if (sg) pushUnique(sg.join, em);
    });

    (Array.isArray(d.removed) ? d.removed : []).forEach(function (s) {
        const em = s && s.email;
        const team = findClassTeam(classTeams, s && s.klasse);
        const gid = groupIdOf(team);
        const g = ensureGroup(map, gid, (s && s.klasse) || 'Klasse');
        if (g) pushUnique(g.leave, em);
        const sg = ensureGroup(map, sammelId, 'Alle Schülerinnen');
        if (sg) pushUnique(sg.leave, em);
    });

    (Array.isArray(d.classChanged) ? d.classChanged : []).forEach(function (row) {
        const s = row && row.student;
        const em = s && s.email;
        const fromTeam = findClassTeam(classTeams, row && row.fromClass);
        const toTeam = findClassTeam(classTeams, row && row.toClass);
        const fromG = ensureGroup(map, groupIdOf(fromTeam), row && row.fromClass);
        const toG = ensureGroup(map, groupIdOf(toTeam), row && row.toClass);
        if (fromG) pushUnique(fromG.leave, em);
        if (toG) pushUnique(toG.join, em);
    });

    const groups = [];
    map.forEach(function (g) {
        if (!g.join.length && !g.leave.length) return;
        groups.push(g);
    });
    return { groups: groups };
}

export function summarizePreview(preview) {
    const groups = preview && Array.isArray(preview.groups) ? preview.groups : [];
    let join = 0;
    let leave = 0;
    groups.forEach(function (g) {
        join += Array.isArray(g.join) ? g.join.length : 0;
        leave += Array.isArray(g.leave) ? g.leave.length : 0;
    });
    return { join: join, leave: leave, groupCount: groups.length };
}

export function hasMembershipWork(preview) {
    const s = summarizePreview(preview);
    return s.join + s.leave > 0;
}

/**
 * Klassengruppe an Stammliste anpassen: Schüler anderer Klassen entfernen, fehlende aufnehmen.
 * Lehrer und sonstige Mitglieder (nicht in allStudentEmails) bleiben.
 */
export function reconcileClassMembers(classEmails, allStudentEmails, currentMemberEmails) {
    const classSet = new Set((Array.isArray(classEmails) ? classEmails : []).map(normEmail).filter(Boolean));
    const studentSet = new Set((Array.isArray(allStudentEmails) ? allStudentEmails : []).map(normEmail).filter(Boolean));
    const members = (Array.isArray(currentMemberEmails) ? currentMemberEmails : []).map(normEmail).filter(Boolean);
    const memberSet = new Set(members);
    const join = [];
    classSet.forEach(function (em) {
        if (em && !memberSet.has(em)) join.push(em);
    });
    const leave = [];
    members.forEach(function (em) {
        if (studentSet.has(em) && !classSet.has(em) && leave.indexOf(em) === -1) leave.push(em);
    });
    return { join: join, leave: leave };
}

export function reconcileSammelgruppe(allStudentEmails, currentMemberEmails) {
    const studentSet = new Set((Array.isArray(allStudentEmails) ? allStudentEmails : []).map(normEmail).filter(Boolean));
    const members = (Array.isArray(currentMemberEmails) ? currentMemberEmails : []).map(normEmail).filter(Boolean);
    const memberSet = new Set(members);
    const join = [];
    studentSet.forEach(function (em) {
        if (em && !memberSet.has(em)) join.push(em);
    });
    const leave = [];
    members.forEach(function (em) {
        if (em && !studentSet.has(em) && leave.indexOf(em) === -1) leave.push(em);
    });
    return { join: join, leave: leave };
}

const api = {
    studentKey: studentKey,
    diffStudents: diffStudents,
    previewMemberships: previewMemberships,
    summarizePreview: summarizePreview,
    hasMembershipWork: hasMembershipWork,
    reconcileClassMembers: reconcileClassMembers,
    reconcileSammelgruppe: reconcileSammelgruppe
};

if (typeof window !== 'undefined') {
    window.ms365StudentClassLifecycle = api;
}

export default api;
