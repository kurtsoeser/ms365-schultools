export const DEFAULT_INTRANET_LIST_TITLES = {
    schueler: 'Schülerinnen',
    faecher: 'Fächer',
    fachgruppen: 'Fachgruppen',
    arges: 'ARGEs',
    klassen: 'Klassen',
    lehrer: 'Lehrerinnen'
};

/** UI: Zeilen „Listen-Typ + Name“ im Schulregister (Stammdaten → SharePoint) */
export const INTRANET_LIST_KIND_OPTIONS = [
    { kind: 'schueler', label: 'Schüler:innen' },
    { kind: 'faecher', label: 'Fächer' },
    { kind: 'fachgruppen', label: 'Fachgruppen' },
    { kind: 'arges', label: 'ARGEs' },
    { kind: 'klassen', label: 'Klassen' },
    { kind: 'lehrer', label: 'Lehrer:innen' }
];

export function resolveIntranetListTitle(kind, storedTitle) {
    const v = String(storedTitle != null ? storedTitle : '').trim();
    if (v) return v;
    const k = String(kind || '').trim();
    return DEFAULT_INTRANET_LIST_TITLES[k] || '';
}
