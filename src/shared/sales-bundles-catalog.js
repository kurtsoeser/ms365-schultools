/**
 * Kunden-Verkaufsbündel (Admin-Referenz + Landing).
 * Technisch: Package-IDs aus license-feature-catalog.js.
 * Bundle-Preise günstiger als Summe der Modul-Listenpreise (priceEuro im Feature-Catalog).
 */

/**
 * @typedef {{
 *   id: string,
 *   label: string,
 *   priceEuro: number,
 *   priceLabel: string,
 *   tagline: string,
 *   audience: string,
 *   packageIds: string[],
 *   highlights: string[],
 *   featured?: boolean
 * }} SalesBundle
 *
 * @typedef {{
 *   id: string,
 *   label: string,
 *   priceEuro: number | null,
 *   priceLabel: string,
 *   tagline: string,
 *   packageIds?: string[],
 *   once?: boolean
 * }} SalesAddon
 */

/** @type {SalesBundle[]} */
export const SALES_BUNDLES = [
    {
        id: 'starter-it',
        label: 'Starter IT',
        priceEuro: 390,
        priceLabel: '390 € / Schuljahr',
        tagline: 'Schuljahr, Personen, Kursteams und Hygiene – der IT-Einstieg.',
        audience: 'IT & Administration',
        packageIds: ['core', 'personen', 'schuljahr', 'unterricht'],
        highlights: [
            'Stammdaten, Personen, Gäste, Lizenzen',
            'Schuljahreswechsel & Import (WebUntis/Sokrates)',
            'Kursteams, Klassengruppen, Diplom, ARGEs',
            'Hygiene, Organigramm, Teams-Antrag'
        ]
    },
    {
        id: 'schulalltag',
        label: 'Schulalltag',
        priceEuro: 890,
        priceLabel: '890 € / Schuljahr',
        tagline: 'Starter plus Lehrer-Portal: Planer, Freistellungen, Intranet.',
        audience: 'IT + Lehrkräfte / Direktion',
        packageIds: [
            'core',
            'personen',
            'schuljahr',
            'unterricht',
            'intranet',
            'schulapps',
            'freistellung'
        ],
        highlights: [
            'Alles aus Starter IT',
            'Schularbeiten, Schulaktivitäten, Projektwochen',
            'Schüler- & Lehrer-Freistellungen',
            'Intranet-Listen + Lehrer-Portal'
        ],
        featured: true
    },
    {
        id: 'schule-plus',
        label: 'Schule Plus',
        priceEuro: 1690,
        priceLabel: '1.690 € / Schuljahr',
        tagline: 'Schulalltag plus Elternkanal, Ausgaben, Betrieb und Automationen.',
        audience: 'Gesamtschule inkl. Support-Story',
        packageIds: [
            'core',
            'personen',
            'schuljahr',
            'unterricht',
            'intranet',
            'schulapps',
            'freistellung',
            'ablauf',
            'ausgaben',
            'betrieb',
            'automationen'
        ],
        highlights: [
            'Alles aus Schulalltag',
            'Elternbrief, Einsammler, Krankmeldung, Kollegium',
            'Ausgaben (Schlüssel/Karten)',
            'Automationen & Betrieb (Termine, Technik, Raum)'
        ]
    }
];

/** @type {SalesAddon[]} */
export const SALES_ADDONS = [
    {
        id: 'addon-ablauf',
        label: 'Eltern & Ablauf',
        priceEuro: 290,
        priceLabel: '+290 € / Jahr',
        tagline: 'Einsammler, Elternbrief, Krankmeldung, Kollegiumsumfragen.',
        packageIds: ['ablauf']
    },
    {
        id: 'addon-ausgaben',
        label: 'Ausgaben & Zugang',
        priceEuro: 150,
        priceLabel: '+150 € / Jahr',
        tagline: 'Schlüssel, Raum-/Schrankenkarten, Parkplatz.',
        packageIds: ['ausgaben']
    },
    {
        id: 'addon-stammdaten-portal',
        label: 'Stammdaten-Portal',
        priceEuro: 390,
        priceLabel: '+390 € / Jahr',
        tagline: 'Öffentliches Einschreibe-/Stammdatenformular für Eltern.',
        packageIds: ['stammdaten-portal']
    },
    {
        id: 'addon-setup',
        label: 'Setup / Onboarding',
        priceEuro: null,
        priceLabel: '590–790 € einmalig',
        tagline: 'Gemeinsam einrichten: Intranet, Listen, Pakete, Audience.',
        once: true
    }
];

/**
 * @param {string} id
 * @returns {SalesBundle|null}
 */
export function salesBundleById(id) {
    const want = String(id || '').trim().toLowerCase();
    return SALES_BUNDLES.find((b) => b.id === want) || null;
}

/**
 * Beste Übereinstimmung: Bundle, dessen packageIds eine Teilmenge der aktiven sind
 * und die größte Schnittmenge hat (bevorzugt exakt).
 * @param {string[]|null|undefined} activePackageIds
 * @returns {{ bundle: SalesBundle|null, exact: boolean }}
 */
export function matchSalesBundle(activePackageIds) {
    const active = new Set(
        (Array.isArray(activePackageIds) ? activePackageIds : [])
            .map((x) => String(x || '').trim().toLowerCase())
            .filter(Boolean)
    );
    if (!active.size) return { bundle: null, exact: false };

    let best = null;
    let bestScore = -1;
    let exact = false;

    for (const bundle of SALES_BUNDLES) {
        const need = bundle.packageIds || [];
        if (!need.length) continue;
        let hit = 0;
        let missing = 0;
        for (const pid of need) {
            if (active.has(pid)) hit += 1;
            else missing += 1;
        }
        if (missing > 0) continue;
        const score = hit;
        const isExact = hit === active.size && missing === 0;
        if (score > bestScore || (score === bestScore && isExact)) {
            best = bundle;
            bestScore = score;
            exact = isExact || hit === need.length;
        }
    }

    return { bundle: best, exact };
}

/**
 * Entitlements-Draft für ein Verkaufsbündel (Pakete setzen, Overrides leeren).
 * @param {string} bundleId
 * @returns {{ packages: string[], toolsInclude: string[], toolsExclude: string[] }|null}
 */
export function entitlementsForSalesBundle(bundleId) {
    const bundle = salesBundleById(bundleId);
    if (!bundle) return null;
    return {
        packages: (bundle.packageIds || []).slice(),
        toolsInclude: [],
        toolsExclude: []
    };
}
