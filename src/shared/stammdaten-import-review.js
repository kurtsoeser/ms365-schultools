/**
 * Zusammenfassung für die Tenant-Review nach WebUntis (Phase 3).
 */
import { summarizeWebuntisMergeResult, diffTeachersImport, summarizeTeachersDiff } from './webuntis-stammdaten-wizard-logic.js';

function normStr(v) {
    return String(v ?? '').trim();
}

/**
 * @param {object} merged mergeWebuntisImportWithExisting-Ergebnis
 * @param {{ counts?: object, teachersLines?: string }} payload
 * @param {object} [sis] ms365SchoolSisImport
 * @param {{ existingTeachersLines?: string, parseTeachersLines?: (t:string)=>array }} [opts]
 * @returns {{ title: string, bullets: string[], hasConflicts: boolean, studentRows: Array<{ kind: string, text: string }>, teacherRows: Array<{ kind: string, text: string }> }}
 */
export function buildWebuntisTenantReview(merged, payload, sis, opts) {
    const stats = (merged && merged.stats) || {};
    const studentDiff = merged && merged.studentDiff;
    const summary = summarizeWebuntisMergeResult(stats, studentDiff, sis || {});
    const counts = (payload && payload.counts) || {};
    const title =
        'WebUntis-Übergabe: ' +
        (counts.teachers || 0) +
        ' Lehrer, ' +
        (counts.students || 0) +
        ' Schüler, ' +
        (counts.subjects || 0) +
        ' Fächer, ' +
        (counts.classes || 0) +
        ' Klassen (markiert für Übernahme)';

    const bullets = [];
    if (summary) bullets.push(summary);
    bullets.push(
        'Es werden nur die markierten Bereiche mit den bestehenden Listen zusammengeführt – nichts wird doppelt angelegt.'
    );

    const studentRows = [];
    if (studentDiff) {
        (studentDiff.added || []).slice(0, 80).forEach(function (r) {
            studentRows.push({
                kind: 'add',
                text: '+ ' + normStr(r.klasse) + ' · ' + normStr(r.name) + (r.email ? ' · ' + r.email : '')
            });
        });
        (studentDiff.updated || []).slice(0, 80).forEach(function (u) {
            const bits = [];
            if (u.klasseChanged) bits.push('Klasse');
            if (u.nameChanged) bits.push('Name');
            if (u.emailChanged) bits.push('E-Mail');
            if (u.parentsChanged) bits.push('Eltern');
            const inc = u.incoming || {};
            studentRows.push({
                kind: 'update',
                text: '~ ' + normStr(inc.klasse) + ' · ' + normStr(inc.name) + ' (' + bits.join(', ') + ')'
            });
        });
        (studentDiff.conflicts || []).forEach(function (c) {
            studentRows.push({ kind: 'conflict', text: '! ' + (c.summary || c.type || 'Konflikt') });
        });
        if ((studentDiff.removed || []).length) {
            bullets.push(
                (studentDiff.removed || []).length +
                    ' Schüler:in(nen) nur lokal, nicht im Import – werden beim Zusammenführen nicht gelöscht.'
            );
        }
    }

    const teacherRows = [];
    const parseTeachers = opts && typeof opts.parseTeachersLines === 'function' ? opts.parseTeachersLines : null;
    let teacherDiff = null;
    if (parseTeachers && payload && normStr(payload.teachersLines)) {
        teacherDiff = diffTeachersImport(
            parseTeachers(opts.existingTeachersLines || ''),
            parseTeachers(payload.teachersLines)
        );
        const tSum = summarizeTeachersDiff(teacherDiff);
        if (tSum && tSum !== 'Keine Änderungen') bullets.push('Lehrer: ' + tSum);
        (teacherDiff.conflicts || []).forEach(function (c) {
            teacherRows.push({ kind: 'conflict', text: '! ' + (c.summary || c.code) });
        });
        (teacherDiff.added || []).slice(0, 60).forEach(function (r) {
            teacherRows.push({
                kind: 'add',
                text: '+ ' + normStr(r.code) + ' · ' + normStr(r.name) + (r.email ? ' · ' + r.email : '')
            });
        });
        (teacherDiff.updated || []).slice(0, 60).forEach(function (u) {
            const inc = u.incoming || {};
            const bits = [];
            if (u.nameChanged) bits.push('Name');
            if (u.emailChanged) bits.push('E-Mail');
            teacherRows.push({
                kind: 'update',
                text: '~ ' + normStr(inc.code) + ' · ' + normStr(inc.name) + ' (' + bits.join(', ') + ')'
            });
        });
    }

    const hasConflicts =
        !!(studentDiff && studentDiff.conflicts && studentDiff.conflicts.length) ||
        !!(teacherDiff && teacherDiff.conflicts && teacherDiff.conflicts.length);
    return { title, bullets, hasConflicts, studentRows, teacherRows };
}
