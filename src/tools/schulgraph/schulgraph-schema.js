/** Knoten- und Kantentypen für die Schuldaten-Visualisierung. */

/** @typedef {'school'|'class'|'subject'|'arge'|'teacher'|'student'|'guardian'|'m365group'|'entraUser'} SchulgraphNodeKind */

/** @typedef {'in_class'|'guardian_of'|'kv_of'|'subject_in_arge'|'teaches'|'class_subject'|'m365_link'|'belongs_to'|'member_of'|'group_member'} SchulgraphEdgeKind */

/** @type {Record<SchulgraphNodeKind, { label: string, icon: string, colorVar: string }>} */
export const NODE_META = {
    school: { label: 'Schule', icon: 'bi-building', colorVar: '--brand1' },
    class: { label: 'Klasse', icon: 'bi-people', colorVar: '#2563eb' },
    subject: { label: 'Fach', icon: 'bi-book', colorVar: '#7c3aed' },
    arge: { label: 'ARGE', icon: 'bi-diagram-3', colorVar: '#c026d3' },
    teacher: { label: 'Lehrkraft', icon: 'bi-person-badge', colorVar: '#059669' },
    student: { label: 'Schüler:in', icon: 'bi-person', colorVar: '#0891b2' },
    guardian: { label: 'Erziehungsberechtigte:r', icon: 'bi-person-hearts', colorVar: '#db2777' },
    m365group: { label: 'M365-Gruppe', icon: 'bi-microsoft-teams', colorVar: '#6264a7' },
    entraUser: { label: 'M365-Benutzer', icon: 'bi-person-circle', colorVar: '#475569' }
};

/** @type {Record<SchulgraphEdgeKind, string>} */
export const EDGE_LABELS = {
    in_class: 'ist in Klasse',
    guardian_of: 'Betreuung',
    kv_of: 'Klassenvorstand',
    subject_in_arge: 'Fach in ARGE',
    teaches: 'unterrichtet',
    class_subject: 'Unterricht in Klasse',
    m365_link: 'verknüpft mit M365',
    belongs_to: 'Teil der Schulstruktur',
    member_of: 'Mitglied in',
    group_member: 'Mitglied in'
};

export const DEFAULT_GRAPH_OPTIONS = {
    /** @deprecated – Migration über aspectsFromLegacyLayers */
    layers: {
        org: true,
        teaching: true,
        people: false,
        m365: true
    },
    preset: 'overview',
    aspects: null,
    klasseFilter: '',
    /** Ab dieser Schülerzahl werden Personen standardmäßig aggregiert (nur Klassen-Knoten). */
    peopleAutoOffThreshold: 80,
    maxStudents: 400,
    maxGuardians: 400,
    includeSchoolHub: true
};
