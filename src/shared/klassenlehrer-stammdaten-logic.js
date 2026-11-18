/**
 * Lehrer je Klasse aus der Unterrichtsbelegung (nicht Klassenvorstand).
 * Quelle: years.*.unterrichtsbelegung.rows – Klasse × Lehrkraft × Fach.
 */

function normStr(v) {
    return String(v ?? '').trim();
}

function normEmail(v) {
    return normStr(v).toLowerCase();
}

function normCode(v) {
    return normStr(v).toUpperCase();
}

function classKey(rowOrCode) {
    if (rowOrCode && typeof rowOrCode === 'object') {
        return normCode(rowOrCode.code) || normStr(rowOrCode.name);
    }
    return normCode(rowOrCode) || normStr(rowOrCode);
}

function teacherKey(row) {
    const em = normEmail(row && row.lehrerEmail);
    if (em && em.indexOf('@') !== -1) return 'e:' + em;
    const code = normCode(row && row.lehrerCode);
    if (code) return 'c:' + code;
    return '';
}

/**
 * Lehrerliste → Optionen für Auswahl (Name + E-Mail).
 * @param {Array<{ name?: string, email?: string, code?: string }>|null|undefined} teachers
 * @returns {Array<{ label: string, name: string, email: string, code: string }>}
 */
export function teacherPickOptions(teachers) {
    const seen = new Set();
    /** @type {Array<{ label: string, name: string, email: string, code: string }>} */
    const out = [];
    (Array.isArray(teachers) ? teachers : []).forEach(function (t) {
        const email = normEmail(t && t.email);
        if (!email || email.indexOf('@') === -1 || seen.has(email)) return;
        seen.add(email);
        const name = normStr(t && t.name) || email;
        const code = normCode(t && t.code);
        out.push({
            label: name + ' · ' + email,
            name: name,
            email: email,
            code: code
        });
    });
    out.sort(function (a, b) {
        return a.name.localeCompare(b.name, 'de');
    });
    return out;
}

/**
 * Belegungszeilen → Lehrkräfte pro Klasse (mit Fächern).
 * @param {Array<object>|null|undefined} belegungRows
 * @returns {Map<string, Array<{ email: string, code: string, name: string, faecher: string[] }>>}
 */
export function teachersWithSubjectsByKlasse(belegungRows) {
    /** @type {Map<string, Map<string, { email: string, code: string, name: string, faecher: string[] }>>} */
    const byClass = new Map();
    (Array.isArray(belegungRows) ? belegungRows : []).forEach(function (r) {
        const klasse = classKey(r && r.klasse) || normStr(r && r.klasse);
        if (!klasse) return;
        const key = teacherKey(r);
        if (!key) return;
        if (!byClass.has(klasse)) byClass.set(klasse, new Map());
        const map = byClass.get(klasse);
        let entry = map.get(key);
        if (!entry) {
            entry = {
                email: normEmail(r.lehrerEmail),
                code: normCode(r.lehrerCode),
                name: normStr(r.lehrerName),
                faecher: []
            };
            map.set(key, entry);
        } else {
            if (!entry.name && normStr(r.lehrerName)) entry.name = normStr(r.lehrerName);
            if (!entry.email && normEmail(r.lehrerEmail)) entry.email = normEmail(r.lehrerEmail);
            if (!entry.code && normCode(r.lehrerCode)) entry.code = normCode(r.lehrerCode);
        }
        const fach = normStr(r.fach);
        if (fach && entry.faecher.indexOf(fach) === -1) entry.faecher.push(fach);
    });
    /** @type {Map<string, Array<{ email: string, code: string, name: string, faecher: string[] }>>} */
    const out = new Map();
    byClass.forEach(function (map, klasse) {
        const list = Array.from(map.values());
        list.forEach(function (t) {
            t.faecher.sort(function (a, b) {
                return a.localeCompare(b, 'de');
            });
        });
        list.sort(function (a, b) {
            return (a.name || a.code || a.email).localeCompare(b.name || b.code || b.email, 'de');
        });
        out.set(klasse, list);
    });
    return out;
}

/**
 * @param {Array<object>|null|undefined} classes
 * @param {Array<object>|null|undefined} belegungRows
 * @param {'all'|'with'|'without'} [filter]
 * @param {string} [query]
 * @returns {Array<{
 *   klasse: string,
 *   classRow: object|null,
 *   teachers: Array<{ email: string, code: string, name: string, faecher: string[] }>,
 *   kvName: string,
 *   kvEmail: string
 * }>}
 */
export function listClassTeachingBoards(classes, belegungRows, filter, query) {
    const byTeachers = teachersWithSubjectsByKlasse(belegungRows);
    const f = String(filter || 'all').toLowerCase();
    const q = normStr(query).toLowerCase();
    /** @type {Map<string, object>} */
    const classByKey = new Map();
    (Array.isArray(classes) ? classes : []).forEach(function (c) {
        const key = classKey(c);
        if (!key) return;
        if (!classByKey.has(key)) classByKey.set(key, c);
    });
    const keys = new Set(classByKey.keys());
    byTeachers.forEach(function (_list, k) {
        keys.add(k);
    });
    /** @type {Array<{ klasse: string, classRow: object|null, teachers: Array<object>, kvName: string, kvEmail: string }>} */
    const out = [];
    Array.from(keys).forEach(function (klasse) {
        const classRow = classByKey.get(klasse) || null;
        const teachers = byTeachers.get(klasse) || [];
        const kvName = normStr(classRow && classRow.headName);
        const kvEmail = normEmail(classRow && classRow.headEmail);
        if (f === 'with' && !teachers.length) return;
        if (f === 'without' && teachers.length) return;
        if (q) {
            const hay = [
                klasse,
                classRow && classRow.name,
                classRow && classRow.year,
                kvName,
                kvEmail,
                teachers
                    .map(function (t) {
                        return [t.name, t.email, t.code, (t.faecher || []).join(' ')].join(' ');
                    })
                    .join(' ')
            ]
                .map(function (x) {
                    return normStr(x).toLowerCase();
                })
                .join(' ');
            if (hay.indexOf(q) === -1) return;
        }
        out.push({
            klasse: klasse,
            classRow: classRow,
            teachers: teachers,
            kvName: kvName,
            kvEmail: kvEmail
        });
    });
    out.sort(function (a, b) {
        return a.klasse.localeCompare(b.klasse, 'de', { numeric: true });
    });
    return out;
}

/**
 * @param {Array<object>|null|undefined} classes
 * @param {Array<object>|null|undefined} belegungRows
 * @returns {{ classTotal: number, withTeachers: number, withoutTeachers: number, teacherAssignments: number, rowCount: number }}
 */
export function summarizeClassTeaching(classes, belegungRows) {
    const boards = listClassTeachingBoards(classes, belegungRows, 'all', '');
    let withTeachers = 0;
    let teacherAssignments = 0;
    boards.forEach(function (b) {
        if (b.teachers.length) withTeachers += 1;
        teacherAssignments += b.teachers.length;
    });
    return {
        classTotal: boards.length,
        withTeachers: withTeachers,
        withoutTeachers: Math.max(0, boards.length - withTeachers),
        teacherAssignments: teacherAssignments,
        rowCount: Array.isArray(belegungRows) ? belegungRows.length : 0
    };
}

/**
 * E-Mails der unterrichtenden Lehrkräfte je Klassenkürzel (für SharePoint-Personenfeld).
 * Fehlende E-Mails in der Belegung werden über die Lehrer-Stammdaten (Kürzel) ergänzt.
 *
 * @param {Array<object>|null|undefined} belegungRows
 * @param {Array<{ code?: string, email?: string }>|null|undefined} teachers
 * @returns {Map<string, string[]>}
 */
export function teachingTeacherEmailsByKlasse(belegungRows, teachers) {
    /** @type {Map<string, string>} */
    const emailByCode = new Map();
    (Array.isArray(teachers) ? teachers : []).forEach(function (t) {
        const code = normCode(t && t.code);
        const email = normEmail(t && t.email);
        if (!code || !email || email.indexOf('@') === -1) return;
        if (!emailByCode.has(code)) emailByCode.set(code, email);
    });
    const byTeachers = teachersWithSubjectsByKlasse(belegungRows);
    /** @type {Map<string, string[]>} */
    const out = new Map();
    byTeachers.forEach(function (list, klasse) {
        const seen = new Set();
        /** @type {string[]} */
        const emails = [];
        (Array.isArray(list) ? list : []).forEach(function (t) {
            let email = normEmail(t && t.email);
            if ((!email || email.indexOf('@') === -1) && t && t.code) {
                email = emailByCode.get(normCode(t.code)) || '';
            }
            if (!email || email.indexOf('@') === -1 || seen.has(email)) return;
            seen.add(email);
            emails.push(email);
        });
        out.set(klasse, emails);
    });
    return out;
}

/**
 * @param {Map<string, string[]>|null|undefined} emailsByKlasse
 * @param {{ code?: string, name?: string }|string|null|undefined} classRowOrCode
 * @returns {string[]}
 */
export function teachingEmailsForClassRow(emailsByKlasse, classRowOrCode) {
    if (!emailsByKlasse || typeof emailsByKlasse.get !== 'function') return [];
    const key = classKey(classRowOrCode);
    if (key && emailsByKlasse.has(key)) return emailsByKlasse.get(key) || [];
    if (classRowOrCode && typeof classRowOrCode === 'object') {
        const nameKey = normCode(classRowOrCode.name) || normStr(classRowOrCode.name).toUpperCase();
        if (nameKey && emailsByKlasse.has(nameKey)) return emailsByKlasse.get(nameKey) || [];
    }
    return [];
}

/**
 * Belegungszeilen → kanonische Zeilen für SharePoint-Liste „Lehrer je Klasse“.
 * Eine Zeile je Klasse × Lehrkraft × Fach (Lookup-fähig über Codes/E-Mails).
 *
 * @param {Array<object>|null|undefined} belegungRows
 * @param {Array<{ code?: string, email?: string, name?: string }>|null|undefined} teachers
 * @returns {Array<{
 *   syncKey: string,
 *   title: string,
 *   klasse: string,
 *   lehrerCode: string,
 *   lehrerEmail: string,
 *   lehrerName: string,
 *   fach: string
 * }>}
 */
export function belegungRowsForSpoSync(belegungRows, teachers) {
    /** @type {Map<string, string>} */
    const emailByCode = new Map();
    /** @type {Map<string, string>} */
    const nameByCode = new Map();
    /** @type {Map<string, string>} */
    const nameByEmail = new Map();
    (Array.isArray(teachers) ? teachers : []).forEach(function (t) {
        const code = normCode(t && t.code);
        const email = normEmail(t && t.email);
        const name = normStr(t && t.name);
        if (code && email && email.indexOf('@') !== -1 && !emailByCode.has(code)) {
            emailByCode.set(code, email);
        }
        if (code && name && !nameByCode.has(code)) nameByCode.set(code, name);
        if (email && name && !nameByEmail.has(email)) nameByEmail.set(email, name);
    });

    /** @type {Map<string, {
     *   syncKey: string,
     *   title: string,
     *   klasse: string,
     *   lehrerCode: string,
     *   lehrerEmail: string,
     *   lehrerName: string,
     *   fach: string
     * }>} */
    const byKey = new Map();
    (Array.isArray(belegungRows) ? belegungRows : []).forEach(function (r) {
        const klasse = classKey(r && r.klasse) || normStr(r && r.klasse).toUpperCase();
        if (!klasse) return;
        let email = normEmail(r && r.lehrerEmail);
        let code = normCode(r && r.lehrerCode);
        if ((!email || email.indexOf('@') === -1) && code) {
            email = emailByCode.get(code) || '';
        }
        if (!code && email) {
            const at = email.indexOf('@');
            if (at > 0) code = normCode(email.slice(0, at));
        }
        if (!email && !code) return;
        const fach = normStr(r && r.fach) || '—';
        const teacherPart = (email && email.indexOf('@') !== -1 ? email : code).toLowerCase();
        const syncKey = klasse.toLowerCase() + '|' + teacherPart + '|' + fach.toLowerCase();
        if (byKey.has(syncKey)) return;
        let name = normStr(r && r.lehrerName);
        if (!name && code) name = nameByCode.get(code) || '';
        if (!name && email) name = nameByEmail.get(email) || '';
        const titleBits = [klasse, name || code || email, fach === '—' ? '' : fach].filter(Boolean);
        byKey.set(syncKey, {
            syncKey: syncKey,
            title: titleBits.join(' · '),
            klasse: klasse,
            lehrerCode: code,
            lehrerEmail: email && email.indexOf('@') !== -1 ? email : '',
            lehrerName: name,
            fach: fach === '—' ? '' : fach
        });
    });
    const out = Array.from(byKey.values());
    out.sort(function (a, b) {
        return a.syncKey.localeCompare(b.syncKey, 'de');
    });
    return out;
}

/**
 * Lehrkraft einer Klasse hinzufügen (eine Belegungszeile, Fach optional).
 * @param {Array<object>} belegungRows
 * @param {string} klasse
 * @param {{ email?: string, code?: string, name?: string, fach?: string }} teacher
 * @returns {Array<object>}
 */
export function addTeacherToClassBelegung(belegungRows, klasse, teacher) {
    const k = classKey(klasse) || normStr(klasse);
    const t = teacher || {};
    const email = normEmail(t.email);
    const code =
        normCode(t.code) ||
        (email && email.indexOf('@') !== -1 ? normCode(email.split('@')[0]) : '');
    const name = normStr(t.name);
    const fach = normStr(t.fach);
    if (!k || (!email && !code)) {
        return Array.isArray(belegungRows) ? belegungRows.slice() : [];
    }
    const rows = (Array.isArray(belegungRows) ? belegungRows : []).map(function (r) {
        return Object.assign({}, r);
    });
    const already = rows.some(function (r) {
        if (classKey(r.klasse) !== k && normStr(r.klasse).toUpperCase() !== k) return false;
        if (email && normEmail(r.lehrerEmail) === email) return true;
        if (!email && code && normCode(r.lehrerCode) === code) return true;
        return false;
    });
    if (already && !fach) return rows;
    if (fach) {
        const sameFach = rows.some(function (r) {
            if (classKey(r.klasse) !== k && normStr(r.klasse).toUpperCase() !== k) return false;
            if (email && normEmail(r.lehrerEmail) !== email) return false;
            if (!email && code && normCode(r.lehrerCode) !== code) return false;
            return normStr(r.fach) === fach;
        });
        if (sameFach) return rows;
    } else if (already) {
        return rows;
    }
    const row = {
        klasse: k,
        lehrerCode: code,
        lehrerEmail: email,
        fach: fach,
        gruppe: '',
        teamName: '',
        gruppenmail: ''
    };
    if (name) row.lehrerName = name;
    rows.push(row);
    rows.sort(function (a, b) {
        const c = normStr(a.klasse).localeCompare(normStr(b.klasse), 'de');
        if (c !== 0) return c;
        const fcmp = normStr(a.fach).localeCompare(normStr(b.fach), 'de');
        if (fcmp !== 0) return fcmp;
        return normCode(a.lehrerCode).localeCompare(normCode(b.lehrerCode), 'de');
    });
    return rows;
}

/**
 * Alle Belegungszeilen einer Lehrkraft in einer Klasse entfernen.
 * @param {Array<object>} belegungRows
 * @param {string} klasse
 * @param {{ email?: string, code?: string }} teacher
 * @returns {Array<object>}
 */
export function removeTeacherFromClassBelegung(belegungRows, klasse, teacher) {
    const k = classKey(klasse) || normStr(klasse);
    const email = normEmail(teacher && teacher.email);
    const code = normCode(teacher && teacher.code);
    return (Array.isArray(belegungRows) ? belegungRows : []).filter(function (r) {
        const rk = classKey(r.klasse) || normStr(r.klasse).toUpperCase();
        if (rk !== k) return true;
        if (email && normEmail(r.lehrerEmail) === email) return false;
        if (!email && code && normCode(r.lehrerCode) === code) return false;
        if (email && !normEmail(r.lehrerEmail) && code && normCode(r.lehrerCode) === code) {
            return false;
        }
        return true;
    });
}

/**
 * Belegungszeilen einer Klasse (für Tabellenansicht), inkl. Index in der Gesamtliste.
 * @param {Array<object>|null|undefined} belegungRows
 * @param {string} klasse
 * @returns {Array<{ index: number, row: object }>}
 */
export function listBelegungRowsForKlasse(belegungRows, klasse) {
    const k = classKey(klasse) || normStr(klasse);
    /** @type {Array<{ index: number, row: object }>} */
    const out = [];
    (Array.isArray(belegungRows) ? belegungRows : []).forEach(function (r, index) {
        const rk = classKey(r && r.klasse) || normStr(r && r.klasse).toUpperCase();
        if (rk !== k) return;
        out.push({ index: index, row: r });
    });
    out.sort(function (a, b) {
        const na = normStr(a.row.lehrerName) || normCode(a.row.lehrerCode) || normEmail(a.row.lehrerEmail);
        const nb = normStr(b.row.lehrerName) || normCode(b.row.lehrerCode) || normEmail(b.row.lehrerEmail);
        const c = na.localeCompare(nb, 'de');
        if (c !== 0) return c;
        return normStr(a.row.fach).localeCompare(normStr(b.row.fach), 'de');
    });
    return out;
}

/**
 * @param {Array<object>} belegungRows
 * @param {number} index
 * @param {{ lehrerName?: string, lehrerEmail?: string, lehrerCode?: string, fach?: string }} patch
 * @returns {Array<object>}
 */
export function patchBelegungRowAt(belegungRows, index, patch) {
    const rows = (Array.isArray(belegungRows) ? belegungRows : []).map(function (r) {
        return Object.assign({}, r);
    });
    if (!rows[index]) return rows;
    const p = patch || {};
    if (Object.prototype.hasOwnProperty.call(p, 'lehrerName')) {
        const name = normStr(p.lehrerName);
        if (name) rows[index].lehrerName = name;
        else delete rows[index].lehrerName;
    }
    if (Object.prototype.hasOwnProperty.call(p, 'lehrerEmail')) {
        rows[index].lehrerEmail = normEmail(p.lehrerEmail);
    }
    if (Object.prototype.hasOwnProperty.call(p, 'lehrerCode')) {
        rows[index].lehrerCode = normCode(p.lehrerCode);
    }
    if (Object.prototype.hasOwnProperty.call(p, 'fach')) {
        rows[index].fach = normStr(p.fach);
    }
    return rows;
}

/**
 * @param {Array<object>} belegungRows
 * @param {number} index
 * @returns {Array<object>}
 */
export function removeBelegungRowAt(belegungRows, index) {
    return (Array.isArray(belegungRows) ? belegungRows : []).filter(function (_r, i) {
        return i !== index;
    });
}

/**
 * Snapshot aus Zeilen bauen (für setUnterrichtsbelegung).
 * @param {Array<object>} rows
 * @param {object|null|undefined} prevSnapshot
 * @returns {object}
 */
export function belegungSnapshotFromRows(rows, prevSnapshot) {
    const prev = prevSnapshot && typeof prevSnapshot === 'object' ? prevSnapshot : {};
    const list = Array.isArray(rows) ? rows : [];
    const classSet = new Set();
    const teacherSet = new Set();
    list.forEach(function (r) {
        if (normStr(r && r.klasse)) classSet.add(normStr(r.klasse));
        if (normEmail(r && r.lehrerEmail)) teacherSet.add(normEmail(r.lehrerEmail));
        else if (normCode(r && r.lehrerCode)) teacherSet.add(normCode(r.lehrerCode));
    });
    return {
        updatedAt: new Date().toISOString(),
        yearPrefix: normStr(prev.yearPrefix),
        source: normStr(prev.source) || 'stammdaten-klassenlehrer',
        teamsCount: list.length,
        classCount: classSet.size,
        teacherCount: teacherSet.size,
        rows: list
    };
}

export default {
    teacherPickOptions,
    teachersWithSubjectsByKlasse,
    listClassTeachingBoards,
    summarizeClassTeaching,
    teachingTeacherEmailsByKlasse,
    teachingEmailsForClassRow,
    belegungRowsForSpoSync,
    addTeacherToClassBelegung,
    removeTeacherFromClassBelegung,
    listBelegungRowsForKlasse,
    patchBelegungRowAt,
    removeBelegungRowAt,
    belegungSnapshotFromRows
};
