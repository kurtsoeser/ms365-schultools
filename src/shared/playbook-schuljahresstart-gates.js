/**
 * Voraussetzungen für das Playbook Schuljahresstart (Stammdaten-Gates).
 */
export const SCHULJAHRSTART_PLAYBOOK_STORAGE_KEY = 'ms365-playbook-schuljahresstart-v1';

export const SCHULJAHRSTART_STEP_IDS = [
    'stammdaten',
    'klassen',
    'slg',
    'kursteams',
    'monitor',
    'lifecycle',
    'cleanup',
    'org'
];

/**
 * @param {Record<string, unknown>|null|undefined} settings
 */
export function tenantCountsFromSettings(settings) {
    const s = settings && typeof settings === 'object' ? settings : {};
    let year = '';
    try {
        if (typeof globalThis !== 'undefined' && globalThis.ms365AppDataV2?.getContainer) {
            const c = globalThis.ms365AppDataV2.getContainer();
            year = String((c && c.years && c.years.current) || '').trim();
        }
    } catch {
        /* ignore */
    }
    if (!year) {
        year = String(s.schoolYear || s.currentSchoolYear || '').trim();
    }
    const classes = Array.isArray(s.classes) ? s.classes.length : 0;
    const students = Array.isArray(s.students) ? s.students.length : 0;
    const teachers = Array.isArray(s.teachers) ? s.teachers.length : 0;
    const subjects = Array.isArray(s.subjects) ? s.subjects.length : 0;
    const domain = String(s.domain || '').trim();
    const schoolName = String(s.schoolName || '').trim();
    return { year, classes, students, teachers, subjects, domain, schoolName };
}

/**
 * @param {ReturnType<typeof tenantCountsFromSettings>} counts
 * @returns {Record<string, { ok: boolean, message: string }>}
 */
export function evaluateSchuljahresstartPrerequisites(counts) {
    const c = counts || {};
    const year = String(c.year || '').trim();
    const classes = Number(c.classes) || 0;
    const students = Number(c.students) || 0;
    const teachers = Number(c.teachers) || 0;
    const subjects = Number(c.subjects) || 0;
    const hasSchool = !!(String(c.domain || '').trim() || String(c.schoolName || '').trim());

    return {
        stammdaten: {
            ok: hasSchool && !!year && classes >= 1,
            message: 'Schuljahr und mindestens eine Klasse im Schulregister pflegen.'
        },
        klassen: {
            ok: classes >= 1,
            message: 'Mindestens eine Klasse im Schulregister (tenant.html → Klassen).'
        },
        slg: {
            ok: students >= 1 || teachers >= 1,
            message: 'Mindestens ein Schüler oder eine Lehrkraft in den Stammdaten.'
        },
        kursteams: {
            ok: subjects >= 1 && classes >= 1,
            message: 'Fächerliste und Klassen im Schulregister – dann Kursteams sinnvoll.'
        },
        monitor: {
            ok: classes >= 1,
            message: 'Klassenliste sollte stehen, bevor der Sync-Monitor Sinn ergibt.'
        },
        lifecycle: {
            ok: students >= 1,
            message: 'Schülerliste im Schulregister für Zu-/Abgänge.'
        },
        cleanup: { ok: true, message: '' },
        org: {
            ok: classes >= 1 && !!year,
            message: 'Schuljahr und Klassen für den Schuljahr-Assistenten.'
        }
    };
}

/**
 * @param {string} stepId
 * @param {number} stepIndex
 * @param {Record<string, boolean>} state
 * @param {Record<string, { ok: boolean, message: string }>} prereq
 * @param {string[]} stepIds
 */
export function evaluateStepUnlock(stepId, stepIndex, state, prereq, stepIds) {
    const ids = Array.isArray(stepIds) ? stepIds : SCHULJAHRSTART_STEP_IDS;
    const st = state && typeof state === 'object' ? state : {};
    const gates = prereq && typeof prereq === 'object' ? prereq : {};

    if (stepIndex <= 0) {
        const g = gates[stepId];
        if (g && !g.ok) return { unlocked: false, message: g.message };
        return { unlocked: true, message: '' };
    }

    const prevId = ids[stepIndex - 1];
    if (st[prevId]) return { unlocked: true, message: '' };

    const gate = gates[stepId];
    if (gate && gate.ok) return { unlocked: true, message: '' };

    if (gate && gate.message) return { unlocked: false, message: gate.message };
    return {
        unlocked: false,
        message: 'Vorherigen Schritt abhaken oder die Stammdaten-Voraussetzung erfüllen.'
    };
}

/**
 * @param {Record<string, boolean>} state
 * @param {string[]} stepIds
 */
export function schuljahresstartPlaybookProgress(state, stepIds) {
    const ids = Array.isArray(stepIds) ? stepIds : SCHULJAHRSTART_STEP_IDS;
    const st = state && typeof state === 'object' ? state : {};
    let done = 0;
    ids.forEach(function (id) {
        if (st[id]) done += 1;
    });
    return { done, total: ids.length };
}

/**
 * @param {Record<string, boolean>} state
 * @param {ReturnType<typeof tenantCountsFromSettings>} counts
 * @param {string[]} stepIds
 */
export function nextSchuljahresstartGateHint(state, counts, stepIds) {
    const ids = Array.isArray(stepIds) ? stepIds : SCHULJAHRSTART_STEP_IDS;
    const st = state && typeof state === 'object' ? state : {};
    const prereq = evaluateSchuljahresstartPrerequisites(counts);
    for (let i = 0; i < ids.length; i++) {
        const id = ids[i];
        if (st[id]) continue;
        const unlock = evaluateStepUnlock(id, i, st, prereq, ids);
        if (!unlock.unlocked && unlock.message) return unlock.message;
        if (!st[id]) return null;
    }
    return null;
}

if (typeof window !== 'undefined') {
    window.ms365SchuljahresstartPlaybook = {
        SCHULJAHRSTART_PLAYBOOK_STORAGE_KEY,
        SCHULJAHRSTART_STEP_IDS,
        tenantCountsFromSettings,
        evaluateSchuljahresstartPrerequisites,
        evaluateStepUnlock,
        schuljahresstartPlaybookProgress,
        nextSchuljahresstartGateHint
    };
}
