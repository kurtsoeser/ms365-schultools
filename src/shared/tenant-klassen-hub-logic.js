/**
 * Klassen-zentrierte Gruppierung für das Schulregister (Phase 4).
 */

export function normClassKey(code, normCodeFn) {
    const fn = normCodeFn || ((v) => String(v ?? '').trim());
    return fn(code).toLowerCase();
}

/**
 * @param {Array<{ code: string, name?: string, headName?: string, headEmail?: string }>} classes
 * @param {Array<{ klasse: string, name: string, email?: string }>} students
 * @param {(v: unknown) => string} normCodeFn
 */
export function groupStudentsByClass(classes, students, normCodeFn) {
    const norm = normCodeFn || ((v) => String(v ?? '').trim());
    const byKey = new Map();
    (classes || []).forEach(function (c) {
        const key = normClassKey(c.code, norm);
        if (!key) return;
        byKey.set(key, {
            classRow: c,
            classKey: key,
            students: []
        });
    });
    const unassigned = [];
    (students || []).forEach(function (s, index) {
        const key = normClassKey(s.klasse, norm);
        const entry = key ? byKey.get(key) : null;
        const item = { index, row: s };
        if (entry) entry.students.push(item);
        else unassigned.push(item);
    });
    const byClass = Array.from(byKey.values()).sort(function (a, b) {
        return String(a.classRow.code || '').localeCompare(String(b.classRow.code || ''), 'de', {
            numeric: true
        });
    });
    byClass.forEach(function (b) {
        b.students.sort(function (x, y) {
            return String(x.row.name || '').localeCompare(String(y.row.name || ''), 'de', {
                sensitivity: 'base'
            });
        });
    });
    return { byClass, unassigned };
}

/**
 * @param {object|null} match Eintrag aus classGroupMatchByKey
 */
export function describeClassM365Match(match) {
    if (!match || typeof match !== 'object') {
        return { kind: 'muted', label: 'M365: nicht geprüft', title: 'Klassengruppe noch nicht mit Entra verknüpft.' };
    }
    if (match.groupId) {
        return {
            kind: 'ok',
            label: 'M365: ' + (match.displayName || match.mailNickname || 'verknüpft'),
            title: 'Microsoft-365-Gruppe ist verknüpft.'
        };
    }
    if (match.notFound) {
        return { kind: 'error', label: 'M365: Gruppe fehlt', title: 'Erwartete Klassengruppe wurde nicht gefunden.' };
    }
    return { kind: 'warn', label: 'M365: offen', title: 'Abgleich ausstehend.' };
}
