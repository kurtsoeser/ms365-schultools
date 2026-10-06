/**
 * WebUntis-Dateien im Schulregister: Ziel-Liste aus Dateiname / Tabellenkopf.
 */

/** @typedef {'subjects'|'teachers'|'students'|'classes'|'arges'|'generic'} RegisterImportTarget */

/**
 * @param {string} filename
 * @returns {'subjects'|'classes'|''}
 */
export function pdfImportTargetFromFilename(filename) {
    const low = String(filename || '').toLowerCase();
    if (/subject|unterrichtsgegenstand|faecher|fächer|fachgruppe/.test(low)) return 'subjects';
    if (/(^|[_\-.])class|klasse/.test(low)) return 'classes';
    return '';
}

/**
 * @param {unknown[][]} [aoa]
 * @param {(aoa: unknown[][]) => string} [detectKind]
 * @returns {RegisterImportTarget}
 */
export function spreadsheetImportTargetFromAoa(aoa, detectKind) {
    const detect = typeof detectKind === 'function' ? detectKind : function () {
        return '';
    };
    const kind = detect(aoa || []);
    if (kind === 'teacher') return 'teachers';
    if (kind === 'student' || kind === 'guardian') return 'students';
    if (kind === 'subject' || kind === 'lessons') return 'subjects';
    return 'generic';
}

/**
 * @param {RegisterImportTarget} originTab
 * @param {RegisterImportTarget} resolved
 * @returns {RegisterImportTarget}
 */
export function resolveRegisterImportTarget(originTab, resolved) {
    if (resolved && resolved !== 'generic') return resolved;
    const o = originTab || 'generic';
    if (o === 'arges') return 'arges';
    if (o === 'subjects' || o === 'teachers' || o === 'students' || o === 'classes') return o;
    return 'generic';
}

export const REGISTER_IMPORT_TAB_BTN = {
    subjects: 'tabMainSubjects',
    teachers: 'tabMainLehrer',
    students: 'tabMainSchueler',
    classes: 'tabMainKlassen',
    arges: 'tabMainArges'
};

export const REGISTER_IMPORT_TARGET_LABEL = {
    subjects: 'Fächerliste',
    teachers: 'Lehrerliste',
    students: 'Schülerliste',
    classes: 'Klassenliste',
    arges: 'ARGE-Liste'
};
