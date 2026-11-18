/**
 * Lizenz-Module (Browser): Pakete → Tool-IDs.
 * Spiegel von backend/license-api/src/lib/feature-catalog.js – Tests prüfen Gleichheit.
 */

/**
 * @typedef {{
 *   id: string,
 *   label: string,
 *   description: string,
 *   requires?: string[],
 *   toolIds: string[],
 *   priceEuro?: number,
 *   priceHint?: string
 * }} LicensePackage
 *
 * priceEuro / priceHint: Listenpreis Einzelmodul (€/Schuljahr), Admin + Verkauf.
 * Bundles (sales-bundles-catalog) sind günstiger als die Summe der Module.
 */

/** @type {LicensePackage[]} */
export const LICENSE_PACKAGES = [
    {
        id: 'core',
        label: 'Basis (Core)',
        description: 'Gruppen, Hygiene, Organigramm und schulweite Grundeinstellungen.',
        priceEuro: 250,
        priceHint: '250 € / Jahr',
        toolIds: [
            'slg-schueler',
            'slg-lehrer',
            'verwaltung',
            'klassenvorstaende',
            'arge-fachgruppen',
            'klassenchats',
            'weitere-teams-gruppen',
            'teams-antrag',
            'sharepoint-liste-teams-antrag',
            'datenhygiene',
            'schulstruktur-sync',
            'datenlandkarte',
            'schul-organigramm',
            'schulgraph',
            'stammdaten-uebergabe',
            'stammdaten-backup-abgleich',
            'cleanup-playbook',
            'datei-migration',
            'nextcloud-teams-migrate',
            'leere-gruppen-report',
            'teams-archiv',
            'teams-kanaele-bulk',
            'gruppenerstellung',
            'sharepoint-mandant-website',
            'sharepoint-mandant-teilen',
            'schul-baseline'
        ]
    },
    {
        id: 'personen',
        label: 'Personen',
        description: 'Konten, Gäste, Lifecycle und M365-Lizenzübersicht.',
        priceEuro: 100,
        priceHint: '100 € / Jahr',
        toolIds: [
            'lizenzverwaltung',
            'personen-verwaltung',
            'namenskonvention-audit',
            'gaeste-verwalten',
            'schueler-lifecycle',
            'pa-gast-erinnerung'
        ]
    },
    {
        id: 'schuljahr',
        label: 'Schuljahr',
        description: 'Import, Schuljahreswechsel und Cleanup.',
        priceEuro: 130,
        priceHint: '130 € / Jahr',
        toolIds: [
            'playbook-schuljahresstart',
            'playbook-daten-import-verknuepfen',
            'bildungsportal-stammdaten',
            'organisations-assistent',
            'klassen-umbenennen',
            'webuntis-sync-monitor',
            'webuntis-stammdaten-import',
            'cleanup-playbook-schuljahr'
        ]
    },
    {
        id: 'unterricht',
        label: 'Unterricht',
        description: 'Unterrichtsteams, OneNote, Diplom und Spielwiesen.',
        priceEuro: 200,
        priceHint: '200 € / Jahr',
        toolIds: [
            'jahrgang',
            'klassen-merge',
            'kursteams',
            'playbook-kursteams',
            'unterrichtsteams-katalog',
            'kursteam-einzeln',
            'kursteam-templates',
            'onenote-verteilung',
            'diplomarbeiten',
            'pa-diplom-ordner',
            'spielwiesen'
        ]
    },
    {
        id: 'kommunikation',
        label: 'Kommunikation',
        description: 'Postfächer, Verteiler, Elternsprechtag und Räume.',
        priceEuro: 150,
        priceHint: '150 € / Jahr',
        toolIds: [
            'postfaecher',
            'kalender-postfaecher',
            'verteilerlisten',
            'eltern-verteiler',
            'elternsprechtag-bookings',
            'playbook-eltern',
            'playbook-elternsprechtag',
            'raeume-ressourcen'
        ]
    },
    {
        id: 'intranet',
        label: 'Intranet & Listen',
        description: 'Schul-Intranet und SharePoint-Listen.',
        priceEuro: 130,
        priceHint: '130 € / Jahr',
        toolIds: [
            'sharepoint-intranet-hub',
            'playbook-intranet',
            'sharepoint-liste-lehrer',
            'sharepoint-liste-stammdaten',
            'sharepoint-liste-schultermine',
            'sharepoint-liste-srdp',
            'srdp-anmeldung',
            'sharepoint-liste-vertretung'
        ]
    },
    {
        id: 'schulapps',
        label: 'Schul-Apps',
        description: 'Schularbeiten, Aktivitäten, Projektwochen und Wahl.',
        requires: ['intranet'],
        priceEuro: 300,
        priceHint: '300 € / Jahr',
        toolIds: [
            'playbook-schularbeiten-planer',
            'sharepoint-liste-schularbeiten',
            'sharepoint-liste-schulaktivitaeten',
            'sharepoint-liste-projektwochen',
            'projektwochen',
            'schularbeiten-planer',
            'schulaktivitaeten-planer',
            'schulaktivitaeten-setup',
            'termine-hub',
            'aufgaben',
            'technik-meldungen',
            'raumplan',
            'schulsprecherwahl',
            'wahlfach-zuteilung',
            'wahlfach-anmeldung',
            'sharepoint-liste-wahlfach'
        ]
    },
    {
        id: 'freistellung',
        label: 'Freistellungen',
        description: 'Schüler- und Lehrer-Freistellungen inkl. Setup.',
        requires: ['schulapps'],
        priceEuro: 150,
        priceHint: '150 € / Jahr',
        toolIds: [
            'freistellung-planer',
            'lehrer-freistellung-planer',
            'freistellung-setup',
            'lehrer-freistellung-setup',
            'playbook-freistellungen'
        ]
    },
    {
        id: 'automationen',
        label: 'Automationen',
        description: 'Power-Automate-Rezepte und Flows.',
        priceEuro: 130,
        priceHint: '130 € / Jahr',
        toolIds: [
            'pa-erst-setup',
            'power-automate-rezepte',
            'pa-termine-sync',
            'pa-schularbeiten-mail',
            'pa-projektwochen-mail',
            'pa-antraege',
            'pa-seminar',
            'pa-schilf',
            'srdp-mail-setup'
        ]
    },
    {
        id: 'stammdaten-portal',
        label: 'Stammdaten-Portal',
        description: 'Öffentliches Schüler-Stammdaten-/Einschreibeportal.',
        requires: ['intranet'],
        priceEuro: 390,
        priceHint: '390 € / Jahr',
        toolIds: ['schueler-stammdaten-portal']
    },
    {
        id: 'ablauf',
        label: 'Eltern & Ablauf',
        description:
            'Einsammler, Elternbrief, Einwilligungen, Lesebestätigungen, Krankmeldung, Kollegiumsumfragen.',
        requires: ['intranet'],
        priceEuro: 290,
        priceHint: '290 € / Jahr',
        toolIds: [
            'einsammler',
            'sharepoint-liste-einsammler',
            'elternbrief',
            'sharepoint-liste-elternbrief',
            'krankmeldung',
            'sharepoint-liste-krankmeldung',
            'kollegium-umfragen',
            'sharepoint-liste-kollegium-umfragen',
            'teams-antrag',
            'sharepoint-liste-teams-antrag'
        ]
    },
    {
        id: 'ausgaben',
        label: 'Ausgaben & Zugang',
        description: 'Schlüssel, Raum-/Schrankenkarten, Parkplatz und eigene Inventar-Typen.',
        requires: ['intranet'],
        priceEuro: 150,
        priceHint: '150 € / Jahr',
        toolIds: ['ausgaben', 'sharepoint-liste-ausgaben']
    },
    {
        id: 'betrieb',
        label: 'Betrieb & Portal',
        description: 'Kalender und Termine, Aufgaben, Technik-Meldungen und Raumplan für den Schulalltag.',
        requires: ['intranet'],
        priceEuro: 130,
        priceHint: '130 € / Jahr',
        toolIds: [
            'termine-hub',
            'aufgaben',
            'sharepoint-liste-aufgaben',
            'technik-meldungen',
            'sharepoint-liste-technik-meldungen',
            'raumplan',
            'sharepoint-liste-raumplan'
        ]
    }
];

export const DEFAULT_NEW_SCHOOL_PACKAGES = ['core', 'personen', 'schuljahr'];

export function allPackageIds() {
    return LICENSE_PACKAGES.map((p) => p.id);
}

export function allKnownToolIds() {
    const set = new Set();
    for (const p of LICENSE_PACKAGES) {
        for (const id of p.toolIds || []) set.add(id);
    }
    return set;
}

export function packageById(packageId) {
    const want = String(packageId || '')
        .trim()
        .toLowerCase();
    return LICENSE_PACKAGES.find((p) => p.id === want) || null;
}

/**
 * Alle Pakete, die dieses Werkzeug enthalten (Tools können in mehreren Paketen stehen).
 * @param {string} toolId
 * @returns {LicensePackage[]}
 */
export function packagesContainingTool(toolId) {
    const id = String(toolId || '');
    if (!id) return [];
    return LICENSE_PACKAGES.filter((p) => (p.toolIds || []).indexOf(id) >= 0);
}

/**
 * Ob das Werkzeug bereits über aktive Pakete freigeschaltet wäre (ohne Exclude).
 * @param {{ packages?: string[] } | null | undefined} doc
 * @param {string} toolId
 */
export function toolCoveredByPackages(doc, toolId) {
    const id = String(toolId || '');
    const pkgs = Array.isArray(doc && doc.packages) ? doc.packages : [];
    if (!id || !pkgs.length) return false;
    const set = {};
    for (let i = 0; i < pkgs.length; i++) set[pkgs[i]] = true;
    for (const p of packagesContainingTool(id)) {
        if (set[p.id]) return true;
    }
    return false;
}

/**
 * Redundante Include/Exclude-Einträge entfernen (nach Paket- oder Tool-Toggles).
 * @param {{ packages?: string[], toolsInclude?: string[], toolsExclude?: string[] }} doc
 */
export function pruneToolOverrides(doc) {
    if (!doc || typeof doc !== 'object') return doc;
    const include = Array.isArray(doc.toolsInclude) ? doc.toolsInclude : [];
    const exclude = Array.isArray(doc.toolsExclude) ? doc.toolsExclude : [];
    doc.toolsInclude = include.filter((id) => !toolCoveredByPackages(doc, id));
    doc.toolsExclude = exclude.filter((id) => toolCoveredByPackages(doc, id));
    return doc;
}

/**
 * Einzelnes Werkzeug freischalten oder sperren (unabhängig vom Paket-Toggle).
 * Schaltet bei Bedarf Paket-Abhängigkeiten (`requires`) mit.
 * @param {{ packages: string[], toolsInclude: string[], toolsExclude: string[] }} doc
 * @param {string} toolId
 * @param {boolean} on
 * @returns {{ packagesEnabled: string[] }}
 */
export function applyToolEntitlementToggle(doc, toolId, on) {
    const id = String(toolId || '');
    /** @type {string[]} */
    const packagesEnabled = [];
    if (!doc || !id) return { packagesEnabled };

    if (!Array.isArray(doc.packages)) doc.packages = [];
    if (!Array.isArray(doc.toolsInclude)) doc.toolsInclude = [];
    if (!Array.isArray(doc.toolsExclude)) doc.toolsExclude = [];

    function addPkg(pid) {
        if (doc.packages.indexOf(pid) >= 0) return;
        doc.packages.push(pid);
        packagesEnabled.push(pid);
    }

    function ensureRequiresForTool() {
        const owners = packagesContainingTool(id);
        for (const pkg of owners) {
            if (!Array.isArray(pkg.requires)) continue;
            for (const req of pkg.requires) {
                const need = packageById(req);
                if (!need) continue;
                addPkg(need.id);
                if (Array.isArray(need.requires)) {
                    need.requires.forEach((r) => {
                        const nested = packageById(r);
                        if (nested) addPkg(nested.id);
                    });
                }
            }
        }
    }

    if (on) {
        doc.toolsExclude = doc.toolsExclude.filter((x) => x !== id);
        ensureRequiresForTool();
        if (!toolCoveredByPackages(doc, id)) {
            if (doc.toolsInclude.indexOf(id) < 0) doc.toolsInclude.push(id);
        }
    } else {
        doc.toolsInclude = doc.toolsInclude.filter((x) => x !== id);
        if (toolCoveredByPackages(doc, id)) {
            if (doc.toolsExclude.indexOf(id) < 0) doc.toolsExclude.push(id);
        }
    }
    pruneToolOverrides(doc);
    return { packagesEnabled };
}

/**
 * @param {unknown} raw
 */
export function parseEntitlementsRaw(raw) {
    if (raw == null || raw === '') return null;
    let obj = raw;
    if (typeof raw === 'string') {
        const s = raw.trim();
        if (!s) return null;
        try {
            obj = JSON.parse(s);
        } catch {
            return null;
        }
    }
    if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return null;
    const packages = Array.isArray(obj.packages)
        ? obj.packages.map((x) => String(x || '').trim().toLowerCase()).filter(Boolean)
        : [];
    const toolsInclude = Array.isArray(obj.toolsInclude)
        ? obj.toolsInclude.map((x) => String(x || '').trim()).filter(Boolean)
        : [];
    const toolsExclude = Array.isArray(obj.toolsExclude)
        ? obj.toolsExclude.map((x) => String(x || '').trim()).filter(Boolean)
        : [];
    return {
        v: Number(obj.v) > 0 ? Number(obj.v) : 1,
        packages,
        toolsInclude,
        toolsExclude,
        notes: String(obj.notes || '').trim()
    };
}

/**
 * @param {{ packages?: string[], toolsInclude?: string[], toolsExclude?: string[], notes?: string } | null | undefined} doc
 */
export function serializeEntitlements(doc) {
    const base = {
        v: 1,
        packages: Array.isArray(doc && doc.packages) ? doc.packages.slice() : [],
        toolsInclude: Array.isArray(doc && doc.toolsInclude) ? doc.toolsInclude.slice() : [],
        toolsExclude: Array.isArray(doc && doc.toolsExclude) ? doc.toolsExclude.slice() : []
    };
    if (doc && doc.notes) base.notes = String(doc.notes);
    return JSON.stringify(base);
}

/**
 * @param {{ packages?: string[], toolsInclude?: string[], toolsExclude?: string[] } | null} doc
 * @param {{ legacyFullAccess?: boolean }} [opts]
 */
export function resolveFeatureEntitlements(doc, opts) {
    const legacyFullAccess = !opts || opts.legacyFullAccess !== false;
    const known = allKnownToolIds();
    const allPkgs = allPackageIds();

    let packages = doc && Array.isArray(doc.packages) ? doc.packages.slice() : null;
    const toolsInclude = (doc && Array.isArray(doc.toolsInclude) ? doc.toolsInclude : []).slice();
    const toolsExclude = (doc && Array.isArray(doc.toolsExclude) ? doc.toolsExclude : []).slice();

    if (!packages || !packages.length) {
        packages = legacyFullAccess ? allPkgs.slice() : [];
    }

    const wildcard = packages.some((p) => p === '*' || p === 'all');
    /** @type {Set<string>} */
    const featureSet = new Set();

    if (wildcard) {
        known.forEach((id) => featureSet.add(id));
        packages = allPkgs.slice();
    } else {
        const validPkgs = [];
        for (const pid of packages) {
            const pkg = packageById(pid);
            if (!pkg) continue;
            validPkgs.push(pkg.id);
            for (const tid of pkg.toolIds) featureSet.add(tid);
        }
        packages = validPkgs;
    }

    for (const id of toolsInclude) {
        if (known.has(id)) featureSet.add(id);
    }
    for (const id of toolsExclude) featureSet.delete(id);

    const features = Array.from(featureSet).sort();
    const denied = Array.from(known)
        .filter((id) => !featureSet.has(id))
        .sort();

    return {
        packages,
        toolsInclude,
        toolsExclude,
        features,
        denied,
        featureCount: features.length,
        legacyFullAccess: !doc || !doc.packages || !doc.packages.length
    };
}

export function defaultNewSchoolEntitlements() {
    return {
        v: 1,
        packages: DEFAULT_NEW_SCHOOL_PACKAGES.slice(),
        toolsInclude: [],
        toolsExclude: []
    };
}

export function packageLabel(packageId) {
    const p = packageById(packageId);
    return (p && p.label) || packageId;
}

/** Bootstrap-Icons je Paket */
export const LICENSE_PACKAGE_ICONS = {
    core: 'bi-shield-check',
    personen: 'bi-person-badge',
    schuljahr: 'bi-calendar2-range',
    unterricht: 'bi-mortarboard',
    kommunikation: 'bi-envelope',
    intranet: 'bi-house-door',
    schulapps: 'bi-window-stack',
    freistellung: 'bi-calendar2-check',
    automationen: 'bi-lightning-charge',
    'stammdaten-portal': 'bi-journal-richtext',
    ablauf: 'bi-envelope-paper',
    ausgaben: 'bi-key',
    betrieb: 'bi-building-gear'
};

/** Listenpreis-Text (€/Jahr). */
export function packagePriceHint(packageId) {
    const p = packageById(packageId);
    return (p && p.priceHint) || '';
}

/** Listenpreis Zahl (€/Schuljahr), 0 wenn unbekannt. */
export function packagePriceEuro(packageId) {
    const p = packageById(packageId);
    const n = p && typeof p.priceEuro === 'number' ? p.priceEuro : NaN;
    return Number.isFinite(n) && n > 0 ? n : 0;
}

/**
 * Summe der Listenpreise mehrerer Module.
 * @param {string[]|null|undefined} packageIds
 */
export function sumPackageListPrices(packageIds) {
    let sum = 0;
    (Array.isArray(packageIds) ? packageIds : []).forEach((id) => {
        sum += packagePriceEuro(id);
    });
    return sum;
}

export function packageIcon(packageId) {
    return LICENSE_PACKAGE_ICONS[packageId] || 'bi-box';
}
