/**
 * Tenant-Panel Klassen (Phase 7b) – Facade + reine Helfer.
 */
export { initTenantKlassenHub } from './tenant-klassen-hub-ui.js';
export { groupStudentsByClass, describeClassM365Match, normClassKey } from './tenant-klassen-hub-logic.js';
export { classesToLines } from './tenant-settings-ui-classes.js';

/**
 * @param {Array<{ code?: string, year?: string }>} classes
 * @param {(code: string) => object|null|undefined} getClassM365
 */
export function summarizeKlassenM365(classes, getClassM365) {
    const rows = Array.isArray(classes) ? classes : [];
    let linked = 0;
    let missing = 0;
    let unchecked = 0;
    rows.forEach(function (c) {
        const code = String((c && c.code) || '').trim();
        if (!code) return;
        const m = typeof getClassM365 === 'function' ? getClassM365(code) : null;
        if (!m) {
            unchecked += 1;
            return;
        }
        if (m.groupId) linked += 1;
        else if (m.notFound) missing += 1;
        else unchecked += 1;
    });
    return { total: rows.length, linked, missing, unchecked };
}
