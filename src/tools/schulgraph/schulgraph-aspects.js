/**
 * Aspekte: welche Beziehungen / Entitäten der Graph abbildet.
 * Presets setzen mehrere Aspekte; „custom“ = manuelle Checkboxen.
 */

/** @typedef {'overview'|'unterricht'|'fach_arge'|'klassen'|'familie'|'m365'|'custom'} GraphPresetId */

/** @typedef {{
 *   schulorganisation: boolean,
 *   klassenvorstand: boolean,
 *   fachgruppen: boolean,
 *   unterricht: boolean,
 *   schueler: boolean,
 *   eltern: boolean,
 *   microsoft365: boolean
 * }} GraphAspects
 */

export const DEFAULT_ASPECTS = {
    schulorganisation: true,
    klassenvorstand: true,
    fachgruppen: true,
    unterricht: true,
    schueler: false,
    eltern: false,
    microsoft365: true
};

/** @type {Array<{ id: keyof GraphAspects, label: string, hint: string }>} */
export const ASPECT_CATALOG = [
    {
        id: 'schulorganisation',
        label: 'Schulorganisation',
        hint: 'Klassen, Fächer, ARGE am Schulzentrum'
    },
    {
        id: 'klassenvorstand',
        label: 'Klassenvorstand',
        hint: 'Lehrkraft ↔ Klasse (KV aus Stammdaten)'
    },
    {
        id: 'fachgruppen',
        label: 'Fächer & ARGE',
        hint: 'Zuordnung Fach → Fachgruppe / ARGE'
    },
    {
        id: 'unterricht',
        label: 'Unterrichtsbelegung',
        hint: 'Lehrkraft ↔ Klasse ↔ Fach (Kursteams)'
    },
    {
        id: 'schueler',
        label: 'Schüler:innen',
        hint: 'Schüler:in → Klasse'
    },
    {
        id: 'eltern',
        label: 'Erziehungsberechtigte',
        hint: 'Eltern → Schüler:in (benötigt Schüler:innen)'
    },
    {
        id: 'microsoft365',
        label: 'Microsoft 365',
        hint: 'Stammdaten ↔ Entra-Gruppen (catalogLinks)'
    }
];

/** @type {Record<GraphPresetId, { label: string, description: string, aspects: GraphAspects }>} */
export const GRAPH_PRESETS = {
    overview: {
        label: 'Gesamtüberblick',
        description: 'Organisation, KV, Fächer/ARGE, Unterricht und M365 – ohne Einzelschüler.',
        aspects: {
            schulorganisation: true,
            klassenvorstand: true,
            fachgruppen: true,
            unterricht: true,
            schueler: false,
            eltern: false,
            microsoft365: true
        }
    },
    unterricht: {
        label: 'Unterricht',
        description: 'Belegung und Fachstruktur – weniger IT-Gruppen.',
        aspects: {
            schulorganisation: true,
            klassenvorstand: true,
            fachgruppen: true,
            unterricht: true,
            schueler: false,
            eltern: false,
            microsoft365: false
        }
    },
    fach_arge: {
        label: 'Fächer & ARGE',
        description: 'Fachkatalog und ARGE-Zuordnungen.',
        aspects: {
            schulorganisation: true,
            klassenvorstand: false,
            fachgruppen: true,
            unterricht: false,
            schueler: false,
            eltern: false,
            microsoft365: false
        }
    },
    klassen: {
        label: 'Klassen & KV',
        description: 'Klassenstruktur und Klassenvorstände.',
        aspects: {
            schulorganisation: true,
            klassenvorstand: true,
            fachgruppen: false,
            unterricht: false,
            schueler: false,
            eltern: false,
            microsoft365: false
        }
    },
    familie: {
        label: 'Familie & Klasse',
        description: 'Schüler:innen, Eltern und Klassen (Datenschutz: Klasse filtern).',
        aspects: {
            schulorganisation: true,
            klassenvorstand: false,
            fachgruppen: false,
            unterricht: false,
            schueler: true,
            eltern: true,
            microsoft365: false
        }
    },
    m365: {
        label: 'Microsoft 365',
        description: 'Gruppen-Verknüpfungen aus der Einrichtung.',
        aspects: {
            schulorganisation: true,
            klassenvorstand: false,
            fachgruppen: false,
            unterricht: false,
            schueler: false,
            eltern: false,
            microsoft365: true
        }
    },
    custom: {
        label: 'Frei kombinieren',
        description: 'Aspekte einzeln aktivieren.',
        aspects: { ...DEFAULT_ASPECTS }
    }
};

/**
 * @param {object} raw
 * @returns {GraphAspects}
 */
export function normalizeAspects(raw) {
    const base = { ...DEFAULT_ASPECTS };
    const a = raw && typeof raw === 'object' ? raw : {};
    ASPECT_CATALOG.forEach(({ id }) => {
        if (typeof a[id] === 'boolean') base[id] = a[id];
    });
    if (base.eltern && !base.schueler) base.eltern = false;
    if (!Object.values(base).some(Boolean)) {
        return { ...DEFAULT_ASPECTS };
    }
    return base;
}

/**
 * Alte Layer-Option (v1) → Aspekte.
 * @param {object} layers
 * @returns {GraphAspects}
 */
export function aspectsFromLegacyLayers(layers) {
    const l = layers && typeof layers === 'object' ? layers : {};
    return normalizeAspects({
        schulorganisation: l.org !== false,
        klassenvorstand: l.org !== false,
        fachgruppen: l.org !== false,
        unterricht: l.teaching !== false,
        schueler: l.people === true,
        eltern: l.people === true,
        microsoft365: l.m365 !== false
    });
}

/**
 * @param {GraphAspects} aspects
 * @returns {GraphPresetId}
 */
export function detectPreset(aspects) {
    const norm = normalizeAspects(aspects);
    for (const [id, preset] of Object.entries(GRAPH_PRESETS)) {
        if (id === 'custom') continue;
        const p = preset.aspects;
        const match = ASPECT_CATALOG.every(({ id: key }) => norm[key] === p[key]);
        if (match) return /** @type {GraphPresetId} */ (id);
    }
    return 'custom';
}

/**
 * @param {GraphAspects} aspects
 */
export function aspectsSummary(aspects) {
    const a = normalizeAspects(aspects);
    return ASPECT_CATALOG.filter(({ id }) => a[id])
        .map(({ label }) => label)
        .join(', ');
}

/**
 * @param {GraphPresetId|string} presetId
 * @param {GraphAspects} [current]
 */
export function aspectsForPreset(presetId, current) {
    const id = String(presetId || 'overview');
    if (id === 'custom' && current) return normalizeAspects(current);
    const preset = GRAPH_PRESETS[id] || GRAPH_PRESETS.overview;
    return normalizeAspects(preset.aspects);
}
