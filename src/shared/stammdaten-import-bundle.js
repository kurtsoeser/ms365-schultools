/**
 * Normalisiertes Ergebnis eines Stammdaten-Imports (Adapter-Ausgabe).
 */

export function emptyStammdatenImportBundle(adapterId) {
    return {
        adapterId: String(adapterId || ''),
        provenance: {},
        counts: { teachers: 0, students: 0, subjects: 0, classes: 0 },
        lines: {
            teachersLines: '',
            studentsLines: '',
            subjectsLines: '',
            classesLines: ''
        },
        meta: {}
    };
}

/**
 * @param {object} payload WebUntis-Handoff oder ähnlich
 */
export function bundleFromWebuntisHandoff(payload) {
    const p = payload && typeof payload === 'object' ? payload : {};
    const counts = p.counts && typeof p.counts === 'object' ? p.counts : {};
    return {
        adapterId: 'webuntis-handoff',
        provenance: { source: 'webuntis', handoff: true, skipTenantReview: !!p.skipTenantReview },
        counts: {
            teachers: Number(counts.teachers) || 0,
            students: Number(counts.students) || 0,
            subjects: Number(counts.subjects) || 0,
            classes: Number(counts.classes) || 0
        },
        lines: {
            teachersLines: String(p.teachersLines || ''),
            studentsLines: String(p.studentsLines || ''),
            subjectsLines: String(p.subjectsLines || ''),
            classesLines: String(p.classesLines || '')
        },
        meta: p.meta && typeof p.meta === 'object' ? Object.assign({}, p.meta) : {},
        rawPayload: p
    };
}

export function bundleCounts(bundle) {
    const b = bundle || {};
    const c = b.counts || {};
    return {
        teachers: Number(c.teachers) || 0,
        students: Number(c.students) || 0,
        subjects: Number(c.subjects) || 0,
        classes: Number(c.classes) || 0
    };
}
