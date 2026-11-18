'use strict';

/**
 * Lizenz-Module: Pakete → Tool-IDs (Dashboard data-tool-id).
 * Single source of truth für evaluate-license + Admin-API.
 */

/**
 * priceEuro / priceHint: Listenpreis Einzelmodul (€/Schuljahr). Bundles günstiger.
 * @type {Array<{ id: string, label: string, description: string, requires?: string[], toolIds: string[], priceEuro?: number, priceHint?: string }>}
 */
const LICENSE_PACKAGES = [
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

/** Neue Schulen: Core + Starter */
const DEFAULT_NEW_SCHOOL_PACKAGES = ['core', 'personen', 'schuljahr'];

const ENTITLEMENTS_FIELD = 'LicenseEntitlements';

/**
 * @returns {string[]}
 */
function allPackageIds() {
    return LICENSE_PACKAGES.map((p) => p.id);
}

/**
 * @returns {Set<string>}
 */
function allKnownToolIds() {
    const set = new Set();
    for (let i = 0; i < LICENSE_PACKAGES.length; i++) {
        const tools = LICENSE_PACKAGES[i].toolIds || [];
        for (let j = 0; j < tools.length; j++) set.add(tools[j]);
    }
    return set;
}

/**
 * @param {string} packageId
 */
function packageById(packageId) {
    const want = String(packageId || '')
        .trim()
        .toLowerCase();
    return LICENSE_PACKAGES.find((p) => p.id === want) || null;
}

/**
 * @param {unknown} raw
 * @returns {{ v: number, packages: string[], toolsInclude: string[], toolsExclude: string[], notes: string } | null}
 */
function parseEntitlementsRaw(raw) {
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
 * @param {{ packages?: string[], toolsInclude?: string[], toolsExclude?: string[], notes?: string, v?: number } | null | undefined} doc
 */
function serializeEntitlements(doc) {
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
function resolveFeatureEntitlements(doc, opts) {
    const legacyFullAccess = !opts || opts.legacyFullAccess !== false;
    const known = allKnownToolIds();
    const allPkgs = allPackageIds();

    let packages = doc && Array.isArray(doc.packages) ? doc.packages.slice() : null;
    const toolsInclude = (doc && Array.isArray(doc.toolsInclude) ? doc.toolsInclude : []).slice();
    const toolsExclude = (doc && Array.isArray(doc.toolsExclude) ? doc.toolsExclude : []).slice();

    if (!packages || !packages.length) {
        if (legacyFullAccess) {
            packages = allPkgs.slice();
        } else {
            packages = [];
        }
    }

    const wildcard = packages.some((p) => p === '*' || p === 'all');
    /** @type {Set<string>} */
    const featureSet = new Set();

    if (wildcard) {
        known.forEach((id) => featureSet.add(id));
        packages = allPkgs.slice();
    } else {
        const validPkgs = [];
        for (let i = 0; i < packages.length; i++) {
            const pkg = packageById(packages[i]);
            if (!pkg) continue;
            validPkgs.push(pkg.id);
            for (let j = 0; j < pkg.toolIds.length; j++) featureSet.add(pkg.toolIds[j]);
        }
        packages = validPkgs;
    }

    for (let i = 0; i < toolsInclude.length; i++) {
        const id = toolsInclude[i];
        if (known.has(id)) featureSet.add(id);
    }
    for (let i = 0; i < toolsExclude.length; i++) {
        featureSet.delete(toolsExclude[i]);
    }

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

/**
 * Entitlement-Dokument für neue Schulen.
 */
function defaultNewSchoolEntitlements() {
    return {
        v: 1,
        packages: DEFAULT_NEW_SCHOOL_PACKAGES.slice(),
        toolsInclude: [],
        toolsExclude: []
    };
}

/**
 * Öffentliche Katalog-Payload für Admin-UI.
 */
function getFeatureCatalogPayload() {
    return {
        packages: LICENSE_PACKAGES.map((p) => ({
            id: p.id,
            label: p.label,
            description: p.description,
            requires: Array.isArray(p.requires) ? p.requires.slice() : [],
            toolIds: p.toolIds.slice(),
            toolCount: p.toolIds.length,
            priceEuro: typeof p.priceEuro === 'number' ? p.priceEuro : null,
            priceHint: p.priceHint || ''
        })),
        defaultNewSchoolPackages: DEFAULT_NEW_SCHOOL_PACKAGES.slice(),
        entitlementsField: ENTITLEMENTS_FIELD
    };
}

module.exports = {
    LICENSE_PACKAGES,
    DEFAULT_NEW_SCHOOL_PACKAGES,
    ENTITLEMENTS_FIELD,
    allPackageIds,
    allKnownToolIds,
    packageById,
    parseEntitlementsRaw,
    serializeEntitlements,
    resolveFeatureEntitlements,
    defaultNewSchoolEntitlements,
    getFeatureCatalogPayload
};
