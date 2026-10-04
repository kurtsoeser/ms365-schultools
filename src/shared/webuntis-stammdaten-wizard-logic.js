/**
 * Logik für den WebUntis-Stammdaten-Import-Wizard (Vorschau, Filter, Übernahme).
 */

import { groupSubjectsByBase, subjectHasVariantSuffix } from './subject-code-family.js';

export const WIZARD_TILES = [
    { id: 'subjects', label: 'Fächer', icon: 'bi-journal-text' },
    { id: 'classes', label: 'Klassen', icon: 'bi-mortarboard' },
    { id: 'teachers', label: 'Lehrer', icon: 'bi-person-workspace' },
    { id: 'students', label: 'Schüler', icon: 'bi-people' },
    { id: 'guardians', label: 'Eltern', icon: 'bi-person-hearts' }
];

export function normStr(v) {
    return String(v ?? '').trim();
}

export function rowMatchesFilter(row, query, keys) {
    const tokens = String(query || '')
        .trim()
        .toLowerCase()
        .split(/\s+/)
        .filter(Boolean);
    if (!tokens.length) return true;
    const hay = (keys || [])
        .map(function (k) {
            return String(row && row[k] != null ? row[k] : '').toLowerCase();
        })
        .join('\n');
    return tokens.every(function (t) {
        return hay.includes(t);
    });
}

/**
 * @param {string} fileName
 * @param {string} headerKind aus detectExportKindFromAoa
 */
export function formatGuardianStudentRef(g) {
    if (!g) return '';
    const stuName = [normStr(g.studentFirstName), normStr(g.studentLastName)].filter(Boolean).join(' ');
    const id = normStr(g.studentInternalId) || normStr(g.studentExternalId);
    if (stuName && id) return stuName + ' (ID ' + id + ')';
    if (stuName) return stuName;
    if (id) return 'ID ' + id;
    return '';
}

export function resolveSheetKind(fileName, headerKind) {
    const fn = String(fileName || '').toLowerCase();
    if (/legalguardian|legal_guardian|guardian|elternstamm/i.test(fn)) return 'guardian';
    if (/^student|schueler|schüler/i.test(fn) || /student_\d/i.test(fn)) return 'student';
    if (/^teacher|lehrer/i.test(fn) || /teacher_\d/.test(fn)) return 'teacher';
    if (/^subject|fach/i.test(fn) || /subject_\d/.test(fn)) return 'subject';
    if (/lesson|stundenplan|exportlesson/i.test(fn)) return 'lessons';
    if (/^class|klassen/i.test(fn)) return 'classes';
    return headerKind || '';
}

function mergeTeachersByCode(list) {
    const by = new Map();
    (list || []).forEach(function (t) {
        const code = normStr(t.code).toUpperCase();
        if (!code) return;
        if (!by.has(code)) by.set(code, Object.assign({}, t, { code: code }));
        else {
            const cur = by.get(code);
            if (!cur.name && t.name) cur.name = t.name;
            if (!cur.email && t.email) cur.email = t.email;
        }
    });
    return Array.from(by.values());
}

function mergeSubjectsByCode(list) {
    const by = new Map();
    (list || []).forEach(function (s) {
        const code = normStr(s.code).toUpperCase();
        if (!code) return;
        if (!by.has(code)) {
            by.set(code, {
                code: code,
                name: normStr(s.name) || code,
                curriculumNo: normStr(s.curriculumNo),
                admin: !!s.admin
            });
        } else {
            const cur = by.get(code);
            if (!cur.name && s.name) cur.name = normStr(s.name);
            if (!cur.curriculumNo && s.curriculumNo) cur.curriculumNo = normStr(s.curriculumNo);
            if (s.admin) cur.admin = true;
        }
    });
    return Array.from(by.values());
}

function mergeClassesByCode(list) {
    const by = new Map();
    (list || []).forEach(function (c) {
        const code = normStr(c.code).toUpperCase();
        if (!code) return;
        if (!by.has(code)) by.set(code, Object.assign({}, c, { code: code }));
        else {
            const cur = by.get(code);
            if (!cur.year && c.year) cur.year = c.year;
            if (!cur.headName && c.headName) cur.headName = c.headName;
            if (!cur.headEmail && c.headEmail) cur.headEmail = c.headEmail;
            if (!cur.headCode && c.headCode) cur.headCode = c.headCode;
        }
    });
    return Array.from(by.values());
}

function entityMergeStats() {
    return { added: 0, updated: 0, unchanged: 0 };
}

function bumpEntityStats(stats, changed) {
    if (changed === 'added') stats.added++;
    else if (changed === 'updated') stats.updated++;
    else stats.unchanged++;
}

/**
 * Bestehende Stammdaten mit WebUntis-Import zusammenführen (keine Doppelzeilen).
 * Schüler: Abgleich per E-Mail, WebUntis-Kennzahl (#id:) und Klasse+Name (school-sis-import).
 *
 * @param {{ teachersLines?: string, studentsLines?: string, subjectsLines?: string, classesLines?: string }} existingLines
 * @param {{ teachersLines?: string, studentsLines?: string, subjectsLines?: string, classesLines?: string, counts?: object }} payload
 * @param {{ webuntis?: object, schoolSis?: object, parseTeachersLines: (t:string)=>array, parseStudentsLines: (t:string)=>array, parseSubjectsLines: (t:string)=>array, parseClassesLines: (t:string)=>array }} deps
 */
export function mergeWebuntisImportWithExisting(existingLines, payload, deps) {
    const wu = deps && deps.webuntis;
    const sis = deps && deps.schoolSis;
    const parseTeachers = deps && deps.parseTeachersLines;
    const parseStudents = deps && deps.parseStudentsLines;
    const parseSubjects = deps && deps.parseSubjectsLines;
    const parseClasses = deps && deps.parseClassesLines;
    const ex = existingLines || {};
    const inc = payload || {};

    const out = {
        teachersLines: normStr(ex.teachersLines),
        studentsLines: normStr(ex.studentsLines),
        subjectsLines: normStr(ex.subjectsLines),
        classesLines: normStr(ex.classesLines)
    };
    const stats = {
        teachers: entityMergeStats(),
        subjects: entityMergeStats(),
        classes: entityMergeStats(),
        students: null
    };
    let studentDiff = null;

    if (typeof parseTeachers === 'function' && normStr(inc.teachersLines)) {
        const by = new Map();
        parseTeachers(ex.teachersLines || '').forEach(function (t) {
            const code = normStr(t.code).toUpperCase();
            if (!code) return;
            by.set(code, {
                code: code,
                name: normStr(t.name),
                email: normStr(t.email).toLowerCase()
            });
        });
        parseTeachers(inc.teachersLines || '').forEach(function (t) {
            const code = normStr(t.code).toUpperCase();
            if (!code) return;
            if (!by.has(code)) {
                by.set(code, {
                    code: code,
                    name: normStr(t.name),
                    email: normStr(t.email).toLowerCase()
                });
                bumpEntityStats(stats.teachers, 'added');
                return;
            }
            const cur = by.get(code);
            let changed = false;
            const inName = normStr(t.name);
            const inEmail = normStr(t.email).toLowerCase();
            if (inName && inName !== cur.name) {
                cur.name = inName;
                changed = true;
            }
            if (inEmail && inEmail !== cur.email) {
                cur.email = inEmail;
                changed = true;
            }
            bumpEntityStats(stats.teachers, changed ? 'updated' : 'unchanged');
        });
        const rows = Array.from(by.values()).sort(function (a, b) {
            return a.code.localeCompare(b.code);
        });
        if (wu && typeof wu.teachersToSemicolonLines === 'function') {
            out.teachersLines = wu.teachersToSemicolonLines(rows);
        } else {
            out.teachersLines = rows
                .map(function (t) {
                    return [t.code, t.name, t.email].filter(Boolean).join(';');
                })
                .join('\n');
        }
    }

    if (typeof parseSubjects === 'function' && normStr(inc.subjectsLines)) {
        const by = new Map();
        parseSubjects(ex.subjectsLines || '').forEach(function (s) {
            const code = normStr(s.code).toUpperCase();
            if (!code) return;
            by.set(code, { code: code, name: normStr(s.name) || code });
        });
        parseSubjects(inc.subjectsLines || '').forEach(function (s) {
            const code = normStr(s.code).toUpperCase();
            if (!code) return;
            if (!by.has(code)) {
                by.set(code, { code: code, name: normStr(s.name) || code });
                bumpEntityStats(stats.subjects, 'added');
                return;
            }
            const cur = by.get(code);
            const inName = normStr(s.name);
            if (inName && inName !== cur.name) {
                cur.name = inName;
                bumpEntityStats(stats.subjects, 'updated');
            } else bumpEntityStats(stats.subjects, 'unchanged');
        });
        const rows = Array.from(by.values()).sort(function (a, b) {
            return a.code.localeCompare(b.code);
        });
        out.subjectsLines = rows
            .map(function (s) {
                return [s.code, s.name || ''].join(';');
            })
            .join('\n');
    }

    if (typeof parseClasses === 'function' && normStr(inc.classesLines)) {
        const by = new Map();
        parseClasses(ex.classesLines || '').forEach(function (c) {
            const code = normStr(c.code).toUpperCase();
            if (!code) return;
            by.set(code, Object.assign({}, c, { code: code }));
        });
        parseClasses(inc.classesLines || '').forEach(function (c) {
            const code = normStr(c.code).toUpperCase();
            if (!code) return;
            if (!by.has(code)) {
                by.set(code, Object.assign({}, c, { code: code }));
                bumpEntityStats(stats.classes, 'added');
                return;
            }
            const cur = by.get(code);
            let changed = false;
            ['year', 'name', 'headName', 'headCode'].forEach(function (k) {
                const v = normStr(c[k]);
                if (!v) return;
                const prev = k === 'headCode' ? normStr(cur[k]).toUpperCase() : normStr(cur[k]);
                const next = k === 'headCode' ? v.toUpperCase() : v;
                if (next !== prev) {
                    cur[k] = next;
                    changed = true;
                }
            });
            const inHeadEmail = normStr(c.headEmail).toLowerCase();
            if (inHeadEmail && inHeadEmail !== normStr(cur.headEmail).toLowerCase()) {
                cur.headEmail = inHeadEmail;
                changed = true;
            }
            bumpEntityStats(stats.classes, changed ? 'updated' : 'unchanged');
        });
        const rows = Array.from(by.values()).sort(function (a, b) {
            return a.code.localeCompare(b.code);
        });
        out.classesLines = rows
            .map(function (c) {
                let line = c.code + ';' + (c.year || '') + ';' + (c.name || '');
                if (c.headName || c.headEmail) line += ';' + (c.headName || '') + ';' + (c.headEmail || '');
                return line;
            })
            .join('\n');
    }

    if (typeof parseStudents === 'function' && normStr(inc.studentsLines)) {
        const existingStudents = parseStudents(ex.studentsLines || '');
        const incomingStudents = parseStudents(inc.studentsLines || '');
        if (sis && typeof sis.preferExistingStudentEmails === 'function' && incomingStudents.length) {
            const incoming = sis.preferExistingStudentEmails(incomingStudents, existingStudents);
            if (typeof sis.diffSisImport === 'function') {
                studentDiff = sis.diffSisImport(existingStudents, incoming);
                stats.students = studentDiff.counts;
            }
            const merged =
                typeof sis.applySisImport === 'function'
                    ? sis.applySisImport(existingStudents, incoming, { mode: 'merge' })
                    : mergeStudentRecordsFallback(existingStudents, incomingStudents);
            if (typeof sis.recordsToSemicolonLines === 'function') {
                out.studentsLines = sis.recordsToSemicolonLines(merged);
            }
        } else if (incomingStudents.length) {
            const merged = mergeStudentRecordsFallback(existingStudents, incomingStudents);
            out.studentsLines =
                typeof sis.recordsToSemicolonLines === 'function'
                    ? sis.recordsToSemicolonLines(merged)
                    : normStr(ex.studentsLines);
        }
    }

    return { lines: out, stats: stats, studentDiff: studentDiff };
}

function mergeStudentRecordsFallback(existing, incoming) {
    const out = (existing || []).slice();
    const keys = new Set();
    out.forEach(function (s) {
        keys.add(studentRowKey(s));
    });
    (incoming || []).forEach(function (s) {
        const k = studentRowKey(s);
        if (keys.has(k)) return;
        keys.add(k);
        out.push(s);
    });
    return out;
}

function studentRowKey(s) {
    const em = normStr(s && s.email).toLowerCase();
    if (em && em.indexOf('@') !== -1) return 'e:' + em;
    const ext = normStr(s && s.externalId).toLowerCase();
    if (ext) return 'x:' + ext;
    return 'n:' + normStr(s && s.klasse).toLowerCase() + '|' + normStr(s && s.name).toLowerCase();
}

function formatEntityMergeStats(s) {
    if (!s) return '';
    return (
        String(s.added || 0) +
        ' neu, ' +
        String(s.updated || 0) +
        ' aktualisiert, ' +
        String(s.unchanged || 0) +
        ' unverändert'
    );
}

/**
 * @param {object} stats mergeWebuntisImportWithExisting.stats
 * @param {object|null} studentDiff
 * @param {{ summarizeSisDiff?: (d:object)=>string }} [sis]
 */
export function summarizeWebuntisMergeResult(stats, studentDiff, sis) {
    const parts = [];
    if (stats && stats.teachers && (stats.teachers.added || stats.teachers.updated || stats.teachers.unchanged)) {
        parts.push('Lehrer: ' + formatEntityMergeStats(stats.teachers));
    }
    if (stats && stats.subjects && (stats.subjects.added || stats.subjects.updated || stats.subjects.unchanged)) {
        parts.push('Fächer: ' + formatEntityMergeStats(stats.subjects));
    }
    if (stats && stats.classes && (stats.classes.added || stats.classes.updated || stats.classes.unchanged)) {
        parts.push('Klassen: ' + formatEntityMergeStats(stats.classes));
    }
    if (studentDiff && sis && typeof sis.summarizeSisDiff === 'function') {
        parts.push('Schüler/Eltern: ' + sis.summarizeSisDiff(studentDiff));
    } else if (stats && stats.students) {
        parts.push(
            'Schüler: ' +
                String(stats.students.added || 0) +
                ' neu, ' +
                String(stats.students.updated || 0) +
                ' geändert'
        );
    }
    return parts.length ? parts.join(' · ') : 'Keine neuen Daten (alles bereits vorhanden).';
}

function withSelection(rows, idFn) {
    return (rows || []).map(function (r, i) {
        const id = idFn ? idFn(r, i) : String(i);
        return Object.assign({ id: id, selected: true }, r);
    });
}

/**
 * @param {{ sheets: { name?: string, aoa: any[][] }[], classImports?: { classes: object[], source?: string }[], classPdfImports?: { text?: string, words?: object[], source?: string }[], schoolYearEnd?: number, kvTeachers?: object[], emailOpts?: object }} input
 * @param {{ webuntis: object, schoolSis?: object }} deps
 */
export function buildWizardPreview(input, deps) {
    const wu = deps && deps.webuntis;
    if (!wu) throw new Error('WebUntis-Modul fehlt.');
    const sheets = Array.isArray(input.sheets) ? input.sheets : [];
    const emailOpts = input.emailOpts || {};
    const bucketSheets = [];
    sheets.forEach(function (sh) {
        const aoa = sh && sh.aoa;
        if (!aoa || !aoa.length) return;
        const headerKind = typeof wu.detectExportKindFromAoa === 'function' ? wu.detectExportKindFromAoa(aoa) : '';
        const kind = resolveSheetKind(sh.name, headerKind);
        bucketSheets.push({ name: sh.name, aoa: aoa, kind: kind });
    });

    const classified =
        typeof wu.classifySheets === 'function'
            ? wu.classifySheets(
                  bucketSheets.map(function (s) {
                      return { name: s.name, aoa: s.aoa };
                  })
              )
            : {};

    const teacherAoas = bucketSheets.filter(function (s) {
        return s.kind === 'teacher';
    }).map(function (s) {
        return s.aoa;
    });
    if (classified.teacherAoa && teacherAoas.indexOf(classified.teacherAoa) < 0) teacherAoas.push(classified.teacherAoa);

    function mergeTeacherAoas(aoas) {
        const list = (aoas || []).filter(function (a) {
            return a && a.length > 1;
        });
        if (!list.length) return [];
        const header = list[0][0];
        const out = [header];
        list.forEach(function (a) {
            for (let r = 1; r < a.length; r++) out.push(a[r]);
        });
        return out;
    }

    let teachers = [];
    const mergedTeacherAoa = mergeTeacherAoas(teacherAoas);
    if (mergedTeacherAoa.length && typeof wu.importTeachersFromWebuntis === 'function') {
        const em = wu.importTeachersFromWebuntis({
            teacherAoa: mergedTeacherAoa,
            domain: emailOpts.domain,
            pattern: emailOpts.pattern,
            firstNameMode: emailOpts.firstNameMode,
            applyEmails: !!emailOpts.domain
        });
        teachers = em && em.teachers ? em.teachers : [];
    } else {
        teacherAoas.forEach(function (aoa) {
            if (typeof wu.parseTeachersAoa === 'function') {
                teachers = teachers.concat(wu.parseTeachersAoa(aoa));
            }
        });
        teachers = mergeTeachersByCode(teachers);
    }

    function mergeGuardianAoas(aoas) {
        const list = (aoas || []).filter(function (a) {
            return a && a.length > 1;
        });
        if (!list.length) return null;
        const header = list[0][0];
        const out = [header];
        list.forEach(function (a) {
            for (let r = 1; r < a.length; r++) out.push(a[r]);
        });
        return out;
    }

    const guardianAoas = bucketSheets
        .filter(function (s) {
            return s.kind === 'guardian';
        })
        .map(function (s) {
            return s.aoa;
        });
    if (classified.guardianAoa && guardianAoas.indexOf(classified.guardianAoa) < 0) {
        guardianAoas.push(classified.guardianAoa);
    }
    const mergedGuardianAoa = mergeGuardianAoas(guardianAoas);

    let studentResult = null;
    if (classified.studentAoa && typeof wu.importStudentsFromWebuntis === 'function') {
        studentResult = wu.importStudentsFromWebuntis({
            studentAoa: classified.studentAoa,
            guardianAoa: mergedGuardianAoa || classified.guardianAoa || null,
            domain: emailOpts.domain,
            pattern: emailOpts.pattern,
            firstNameMode: emailOpts.firstNameMode,
            applyEmails: !!emailOpts.domain
        });
    } else if (mergedGuardianAoa && typeof wu.importFromSheets === 'function') {
        studentResult = wu.importFromSheets({
            sheets: [{ aoa: mergedGuardianAoa }],
            applyEmails: false
        }).students;
    }

    const students = studentResult && studentResult.records ? studentResult.records : [];
    const unmatchedGuardians =
        studentResult && studentResult.unmatchedGuardians ? studentResult.unmatchedGuardians : [];
    let guardianParseCount = 0;
    if (mergedGuardianAoa && typeof wu.parseGuardiansAoa === 'function') {
        guardianParseCount = wu.parseGuardiansAoa(mergedGuardianAoa).length;
    }

    let subjects = [];
    bucketSheets.forEach(function (s) {
        if (s.kind === 'subject' && typeof wu.parseSubjectsAoa === 'function') {
            subjects = subjects.concat(wu.parseSubjectsAoa(s.aoa));
        }
        if (s.kind === 'lessons' && typeof wu.uniqueSubjectsFromLessonsAoa === 'function') {
            subjects = subjects.concat(wu.uniqueSubjectsFromLessonsAoa(s.aoa));
        }
    });
    if (classified.subjectAoa && typeof wu.parseSubjectsAoa === 'function') {
        subjects = subjects.concat(wu.parseSubjectsAoa(classified.subjectAoa));
    }
    if (classified.lessonsAoa && typeof wu.uniqueSubjectsFromLessonsAoa === 'function') {
        subjects = subjects.concat(wu.uniqueSubjectsFromLessonsAoa(classified.lessonsAoa));
    }
    (input.subjectPdfImports || []).forEach(function (block) {
        (block.subjects || []).forEach(function (s) {
            subjects.push(s);
        });
    });
    subjects = mergeSubjectsByCode(subjects);

    const classes = [];
    (input.classImports || []).forEach(function (block) {
        (block.classes || []).forEach(function (c) {
            classes.push(Object.assign({}, c));
        });
    });

    const kvTeachers = Array.isArray(input.kvTeachers) ? input.kvTeachers : [];
    const teachersForKv = mergeTeachersByCode(teachers.concat(kvTeachers));

    (input.classPdfImports || []).forEach(function (block) {
        if (!block || typeof wu.importClassesFromWebuntisPdf !== 'function') return;
        const imported = wu.importClassesFromWebuntisPdf({
            text: block.text || '',
            words: block.words || null,
            teachers: teachersForKv,
            skipWithoutTeacher: true,
            inferYear: true,
            schoolYearEnd: input.schoolYearEnd
        });
        (imported.classes || []).forEach(function (c) {
            classes.push(Object.assign({}, c));
        });
    });

    if (classes.length && typeof wu.enrichClassesWithTeachers === 'function' && teachersForKv.length) {
        const again = wu.enrichClassesWithTeachers(classes, teachersForKv);
        classes.length = 0;
        (again.classes || []).forEach(function (c) {
            classes.push(c);
        });
    }

    const warnings = [];
    if (classified.guardianAoa && !classified.studentAoa) {
        warnings.push('Eltern-Export ohne Schüler-Export – Zuordnung eingeschränkt.');
    }
    if (studentResult && studentResult.meta && studentResult.meta.error) {
        warnings.push(String(studentResult.meta.error));
    }
    if ((input.classPdfImports || []).length && !classes.length) {
        warnings.push(
            'Klassen-PDF: keine Klassen mit Klassenlehrkraft erkannt – Teacher-CSV mitladen oder PDF prüfen.'
        );
    }
    bucketSheets
        .filter(function (s) {
            return !s.kind;
        })
        .forEach(function (s) {
            warnings.push('Nicht erkannt: ' + (s.name || 'Datei'));
        });

    const result = {
        warnings: warnings,
        fileSummary: bucketSheets.map(function (s) {
            return { name: s.name, kind: s.kind || 'unknown' };
        }),
        subjects: withSelection(
            subjects.map(function (s) {
                return {
                    code: normStr(s.code).toUpperCase(),
                    name: normStr(s.name) || normStr(s.code),
                    curriculumNo: normStr(s.curriculumNo),
                    admin: !!s.admin
                };
            }),
            function (r) {
                return 'sub:' + r.code;
            }
        ),
        classes: withSelection(
            mergeClassesByCode(classes).map(function (c) {
                return {
                    code: normStr(c.code).toUpperCase(),
                    year: normStr(c.year),
                    name: normStr(c.name || c.code),
                    headName: normStr(c.headName || ''),
                    headEmail: normStr(c.headEmail || '').toLowerCase(),
                    headCode: normStr(c.headCode || '').toUpperCase()
                };
            }),
            function (r) {
                return 'cls:' + r.code;
            }
        ),
        teachers: withSelection(
            teachers.map(function (t) {
                return {
                    code: normStr(t.code).toUpperCase(),
                    name: normStr(t.name),
                    email: normStr(t.email).toLowerCase()
                };
            }),
            function (r) {
                return 'tch:' + r.code;
            }
        ),
        students: withSelection(
            students.map(function (s) {
                const pairs = Array.isArray(s.parentPairs) ? s.parentPairs : [];
                return {
                    klasse: normStr(s.klasse),
                    name: normStr(s.name),
                    email: normStr(s.email).toLowerCase(),
                    externalId: normStr(s.externalId),
                    parentCount: pairs.length,
                    parentPairs: pairs,
                    parentSummary: pairs
                        .map(function (p) {
                            return normStr(p.name) || normStr(p.email);
                        })
                        .filter(Boolean)
                        .join(', ')
                };
            }),
            function (r, i) {
                return 'stu:' + (r.email || r.externalId || r.name || i);
            }
        ),
        guardians: withSelection(
            unmatchedGuardians.map(function (g, i) {
                const ref = formatGuardianStudentRef(g);
                return {
                    name: normStr(g.name),
                    email: normStr(g.email).toLowerCase(),
                    phone: normStr(g.phone),
                    studentRef: ref,
                    note: ref
                        ? 'Schüler:in nicht in Student-Export – Zuordnung prüfen'
                        : 'Kein Schülerbezug in der Elternzeile'
                };
            }),
            function (r, i) {
                return 'gua:' + (r.email || i);
            }
        ),
        meta: {
            subjectCount: subjects.length,
            classCount: classes.length,
            teacherCount: teachers.length,
            studentCount: students.length,
            unmatchedGuardians: unmatchedGuardians.length,
            guardianRows: guardianParseCount,
            guardiansLinked: Math.max(0, guardianParseCount - unmatchedGuardians.length),
            privateStudentEmailsIgnored:
                studentResult && studentResult.meta && studentResult.meta.privateStudentEmailsIgnored
                    ? studentResult.meta.privateStudentEmailsIgnored
                    : 0
        }
    };
    return applySubjectImportDefaults(result);
}

const SUBJECT_ADMIN_CODES = new Set(['ADM', 'DIR', 'KUST', 'BFK', 'BIB', 'AUFSICHT', 'ORD', 'KV']);

/**
 * @param {{ code?: string, admin?: boolean, selected?: boolean }[]} rows
 * @returns {{ total: number, selected: number, excluded: number }}
 */
export function applyRecommendedSubjectImport(rows) {
    const list = rows || [];
    const codes = list
        .map(function (r) {
            return normStr(r.code).toUpperCase();
        })
        .filter(Boolean);
    const exclude = new Set(SUBJECT_ADMIN_CODES);
    codes.forEach(function (c) {
        if (subjectHasVariantSuffix(c)) exclude.add(c);
    });
    groupSubjectsByBase(codes).forEach(function (g) {
        if (!g.isFamily) return;
        g.variants.forEach(function (v) {
            if (v !== g.base) exclude.add(v);
        });
    });
    let selected = 0;
    list.forEach(function (r) {
        const code = normStr(r.code).toUpperCase();
        const pick = code && !exclude.has(code) && !r.admin;
        r.selected = !!pick;
        if (pick) selected++;
    });
    return { total: list.length, selected, excluded: list.length - selected };
}

export function subjectImportStats(rows) {
    const list = rows || [];
    const selected = list.filter(function (r) {
        return r.selected;
    }).length;
    return { total: list.length, selected, excluded: list.length - selected };
}

/** Bei vielen Fächern empfohlene Auswahl (Basis + ohne Verwaltung), sonst alles an. */
export function applySubjectImportDefaults(preview, options) {
    const p = preview || {};
    const rows = p.subjects || [];
    const threshold =
        options && options.deselectThreshold != null ? Number(options.deselectThreshold) : 25;
    if (rows.length > threshold) {
        const stats = applyRecommendedSubjectImport(rows);
        p.subjectImportHint =
            'Empfohlene Auswahl: ' +
            stats.selected +
            ' von ' +
            stats.total +
            ' Fächern werden übernommen (ohne Verwaltung/KUST; pro Familie nur das Haupt-Kürzel – ohne Endziffer, Ü oder +). Liste unten prüfen – bei Bedarf unter „Einzelne Fächer anpassen“.';
    } else if (rows.length) {
        rows.forEach(function (r) {
            r.selected = true;
        });
    }
    return p;
}

export function subjectRowPassesQuickFilter(row, mode) {
    const m = String(mode || 'all');
    if (m === 'all') return true;
    if (m === 'numbered') {
        const no = Number(row && row.curriculumNo);
        return no >= 1 && no <= 20;
    }
    if (m === 'teaching') {
        return !(row && row.admin);
    }
    return true;
}

/**
 * @param {object} preview buildWizardPreview-Ergebnis
 * @param {{ webuntis?: object, schoolSis?: object }} deps
 */
export function compileWizardApplyPayload(preview, deps) {
    const wu = deps && deps.webuntis;
    const sis = deps && deps.schoolSis;
    const p = preview || {};

    const teachers = (p.teachers || []).filter(function (r) {
        return r.selected;
    });
    const students = (p.students || []).filter(function (r) {
        return r.selected;
    });
    const subjects = (p.subjects || []).filter(function (r) {
        return r.selected;
    });
    const classes = (p.classes || []).filter(function (r) {
        return r.selected;
    });

    let teachersLines = '';
    if (wu && typeof wu.teachersToSemicolonLines === 'function') {
        teachersLines = wu.teachersToSemicolonLines(teachers);
    } else {
        teachersLines = teachers
            .map(function (t) {
                return [t.code, t.name, t.email].filter(Boolean).join(';');
            })
            .join('\n');
    }

    let studentsLines = '';
    if (sis && typeof sis.recordsToSemicolonLines === 'function') {
        studentsLines = sis.recordsToSemicolonLines(
            students.map(function (s) {
                return {
                    klasse: s.klasse,
                    name: s.name,
                    email: s.email,
                    externalId: s.externalId,
                    parentPairs: s.parentPairs || []
                };
            })
        );
    } else {
        studentsLines = students
            .map(function (s) {
                const parts = [s.klasse, s.name, s.email];
                (s.parentPairs || []).forEach(function (pair) {
                    parts.push(normStr(pair.name));
                    parts.push(normStr(pair.email));
                });
                return parts.join(';');
            })
            .join('\n');
    }

    const subjectsLines = subjects
        .map(function (s) {
            return [s.code, s.name || ''].join(';');
        })
        .join('\n');

    const classesLines = classes
        .map(function (c) {
            let line = c.code + ';' + (c.year || '') + ';' + (c.name || '');
            if (c.headName || c.headEmail) line += ';' + (c.headName || '') + ';' + (c.headEmail || '');
            return line;
        })
        .join('\n');

    return {
        teachersLines: teachersLines,
        studentsLines: studentsLines,
        subjectsLines: subjectsLines,
        classesLines: classesLines,
        counts: {
            teachers: teachers.length,
            students: students.length,
            subjects: subjects.length,
            classes: classes.length
        }
    };
}
