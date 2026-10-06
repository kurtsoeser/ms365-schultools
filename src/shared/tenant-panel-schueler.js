/**
 * Tenant-Panel Schüler (Phase 7b) – reine Helfer + Events.
 * Tabellen-Render bleibt vorerst in tenant-settings-ui (Closure/DOM).
 */
export { studentsToLines, sortStudentRows } from './tenant-settings-ui-students.js';
export { groupStudentsByClass } from './tenant-klassen-hub-logic.js';

export function notifyTenantStudentsChanged(reason) {
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-tenant-settings-changed', {
                detail: { reason: reason || 'students' }
            })
        );
    } catch {
        /* ignore */
    }
}

/**
 * @param {{ getStudents: () => object[], getClasses: () => object[], normClassCode?: (v:unknown)=>string }} api
 */
export function summarizeStudentsByClass(api) {
    if (!api || typeof api.getStudents !== 'function' || typeof api.getClasses !== 'function') {
        return { classCount: 0, studentCount: 0, unassigned: 0 };
    }
    const classes = api.getClasses() || [];
    const students = api.getStudents() || [];
    const norm =
        typeof api.normClassCode === 'function'
            ? api.normClassCode
            : function (v) {
                  return String(v ?? '').trim();
              };
    const keys = new Set(
        classes.map(function (c) {
            return norm(c && c.code).toLowerCase();
        }).filter(Boolean)
    );
    let unassigned = 0;
    students.forEach(function (s) {
        const k = norm(s && s.klasse).toLowerCase();
        if (!k || !keys.has(k)) unassigned += 1;
    });
    return {
        classCount: classes.length,
        studentCount: students.length,
        unassigned: unassigned
    };
}
