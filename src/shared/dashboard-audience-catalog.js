/**
 * Sichtbarkeit von Dashboard-Werkzeugen nach Persona (IT / Lehrkraft / Schüler).
 * tool-id entspricht data-tool-id im Werkzeugkatalog (index.html).
 *
 * audience:
 * - it: nur Schul-IT / Administration
 * - lehrer: Lehrkräfte (und IT)
 * - schueler: Schülerinnen und Schüler (und IT)
 * - all: alle angemeldeten Nutzer mit erkannter Rolle
 *
 * planner: optionale Zusatzbedingung aus Schularbeiten- oder Freistellungs-Planer
 */

/** @typedef {'it'|'lehrer'|'schueler'|'all'} DashboardAudience */

/**
 * @typedef {{
 *   audience: DashboardAudience|DashboardAudience[],
 *   planner?: { app: 'schularbeiten'|'freistellung', roles: string[] }
 * }} DashboardToolRule
 */

/** @type {Record<string, DashboardToolRule>} */
export const DASHBOARD_TOOL_RULES = {
    // —— Unterricht (IT-heavy) ——
    jahrgang: { audience: 'it' },
    'klassen-merge': { audience: 'it' },
    kursteams: { audience: 'lehrer' },
    'playbook-kursteams': { audience: 'it' },
    'unterrichtsteams-katalog': { audience: 'lehrer' },
    'kursteam-einzeln': { audience: 'lehrer' },
    'kursteam-templates': { audience: 'it' },
    'onenote-verteilung': { audience: 'lehrer' },
    'arge-fachgruppen': { audience: 'it' },
    diplomarbeiten: { audience: 'lehrer' },
    'pa-diplom-ordner': { audience: 'it' },
    spielwiesen: { audience: 'it' },
    klassenchats: { audience: 'lehrer' },

    // —— Personen ——
    lizenzverwaltung: { audience: 'it' },
    'personen-verwaltung': { audience: 'it' },
    'namenskonvention-audit': { audience: 'it' },
    'gaeste-verwalten': { audience: 'it' },
    'schueler-lifecycle': { audience: 'it' },
    'pa-gast-erinnerung': { audience: 'it' },
    'slg-schueler': { audience: 'it' },
    'slg-lehrer': { audience: 'it' },
    verwaltung: { audience: 'it' },
    klassenvorstaende: { audience: 'it' },
    'weitere-teams-gruppen': { audience: 'it' },
    'teams-antrag': { audience: 'all' },
    'sharepoint-liste-teams-antrag': { audience: 'it' },

    // —— Schuljahr ——
    'playbook-schuljahresstart': { audience: 'it' },
    'playbook-daten-import-verknuepfen': { audience: 'it' },
    'bildungsportal-stammdaten': { audience: 'it' },
    'organisations-assistent': { audience: 'it' },
    'klassen-umbenennen': { audience: 'it' },
    'webuntis-sync-monitor': { audience: 'it' },
    'webuntis-stammdaten-import': { audience: 'it' },
    'cleanup-playbook-schuljahr': { audience: 'it' },

    // —— Kommunikation ——
    postfaecher: { audience: 'it' },
    'kalender-postfaecher': { audience: 'it' },
    verteilerlisten: { audience: 'it' },
    'eltern-verteiler': { audience: 'it' },
    'elternsprechtag-bookings': { audience: 'lehrer' },
    'playbook-eltern': { audience: 'it' },
    'playbook-elternsprechtag': { audience: 'it' },
    'raeume-ressourcen': { audience: 'it' },

    // —— Intranet & Apps ——
    'sharepoint-intranet-hub': { audience: 'all' },
    'playbook-intranet': { audience: 'it' },
    'playbook-schularbeiten-planer': { audience: 'it' },
    'sharepoint-liste-lehrer': { audience: 'it' },
    'sharepoint-liste-stammdaten': { audience: 'it' },
    'sharepoint-liste-schultermine': { audience: 'it' },
    'sharepoint-liste-schularbeiten': { audience: 'it' },
    'sharepoint-liste-schulaktivitaeten': { audience: 'it' },
    'sharepoint-liste-projektwochen': { audience: 'it' },
    projektwochen: { audience: 'lehrer' },
    'sharepoint-liste-srdp': { audience: 'it' },
    'srdp-anmeldung': {
        audience: ['lehrer', 'schueler'],
        planner: { app: 'srdp', roles: ['admin', 'lehrer', 'schueler'] }
    },
    'schularbeiten-planer': {
        audience: ['lehrer', 'schueler'],
        planner: { app: 'schularbeiten', roles: ['admin', 'lehrer', 'schueler'] }
    },
    'schulaktivitaeten-planer': {
        audience: 'lehrer',
        planner: { app: 'schularbeiten', roles: ['admin', 'lehrer', 'kv'] }
    },
    'freistellung-planer': {
        audience: ['lehrer', 'schueler'],
        planner: { app: 'freistellung', roles: ['direktion', 'kv', 'schueler'] }
    },
    'lehrer-freistellung-planer': {
        audience: 'lehrer',
        planner: { app: 'schularbeiten', roles: ['admin', 'lehrer'] }
    },
    schulsprecherwahl: { audience: ['lehrer', 'schueler'] },
    'wahlfach-zuteilung': { audience: 'lehrer' },
    'wahlfach-anmeldung': { audience: 'schueler' },
    'sharepoint-liste-wahlfach': { audience: 'it' },
    'schueler-stammdaten-portal': { audience: 'it' },
    einsammler: { audience: 'lehrer' },
    'sharepoint-liste-einsammler': { audience: 'it' },
    elternbrief: { audience: 'lehrer' },
    'sharepoint-liste-elternbrief': { audience: 'it' },
    krankmeldung: { audience: 'lehrer' },
    'sharepoint-liste-krankmeldung': { audience: 'it' },
    'kollegium-umfragen': { audience: 'lehrer' },
    'sharepoint-liste-kollegium-umfragen': { audience: 'it' },
    ausgaben: { audience: 'lehrer' },
    'sharepoint-liste-ausgaben': { audience: 'it' },
    'sharepoint-liste-aufgaben': { audience: 'it' },
    'sharepoint-liste-technik-meldungen': { audience: 'it' },
    'sharepoint-liste-raumplan': { audience: 'it' },
    'sharepoint-liste-vertretung': { audience: 'it' },
    'termine-hub': { audience: ['lehrer', 'schueler'] },
    aufgaben: { audience: 'lehrer' },
    'technik-meldungen': { audience: ['lehrer', 'schueler'] },
    raumplan: { audience: 'lehrer' },

    // —— Automationen ——
    'pa-erst-setup': { audience: 'it' },
    'power-automate-rezepte': { audience: 'it' },
    'freistellung-setup': { audience: 'it' },
    'schulaktivitaeten-setup': { audience: 'it' },
    'srdp-mail-setup': { audience: 'it' },
    'playbook-freistellungen': { audience: 'it' },
    'pa-termine-sync': { audience: 'it' },
    'pa-schularbeiten-mail': { audience: 'it' },
    'pa-projektwochen-mail': { audience: 'it' },
    'pa-antraege': { audience: 'it' },
    'pa-seminar': { audience: 'it' },
    'pa-schilf': { audience: 'it' },

    // —— Einstellungen ——
    gruppenerstellung: { audience: 'it' },
    'sharepoint-mandant-website': { audience: 'it' },
    'sharepoint-mandant-teilen': { audience: 'it' },
    'schulstruktur-sync': { audience: 'it' },
    datenhygiene: { audience: 'it' },
    datenlandkarte: { audience: 'it' },
    'schul-organigramm': { audience: 'it' },
    schulgraph: { audience: 'it' },
    'stammdaten-uebergabe': { audience: 'it' },
    'it-berechtigungen': { audience: 'it' },
    'stammdaten-backup-abgleich': { audience: 'it' },
    'cleanup-playbook': { audience: 'it' },
    'schul-baseline': { audience: 'it' },
    'datei-migration': { audience: 'it' },
    'nextcloud-teams-migrate': { audience: 'it' },
    'leere-gruppen-report': { audience: 'it' },
    'teams-archiv': { audience: 'it' },
    'teams-kanaele-bulk': { audience: 'it' }
};

/**
 * Aufgabenorientierte Cluster (Sidebar „Alle Werkzeuge“).
 * `panel`: DOM-Tab (dash-panel-*), kann vom Cluster-ID abweichen (Legacy-Panel-IDs).
 */
export const DASHBOARD_CLUSTER_ORDER = [
    'gruppen',
    'unterricht',
    'personen',
    'planung',
    'intranet',
    'schulapps',
    'hygiene',
    'kommunikation',
    'automationen'
];

/** @type {Record<string, { label: string, icon: string, panel: string }>} */
export const DASHBOARD_CLUSTER_META = {
    gruppen: { label: '01 Mitgliedschaften', icon: 'bi-people', panel: 'gruppen' },
    unterricht: { label: '02 Klassen & Unterricht', icon: 'bi-mortarboard', panel: 'unterricht' },
    personen: { label: '03 Personen & Gäste', icon: 'bi-person-badge', panel: 'personen' },
    planung: { label: '04 Schuljahresstart', icon: 'bi-rocket-takeoff', panel: 'schuljahr' },
    intranet: { label: '05 Intranet', icon: 'bi-house-door', panel: 'intranet' },
    schulapps: { label: '06 Schul-Apps', icon: 'bi-window-stack', panel: 'schulapps' },
    hygiene: { label: '07 Aufräumen & Audit', icon: 'bi-broom', panel: 'kommunikation' },
    kommunikation: { label: 'Kommunikation', icon: 'bi-envelope', panel: 'intranet' },
    automationen: { label: 'Automatisieren', icon: 'bi-lightning-charge', panel: 'automationen' }
};

/** @type {Record<string, string>} */
export const DASHBOARD_TOOL_CLUSTER = {
    jahrgang: 'unterricht',
    'klassen-merge': 'unterricht',
    kursteams: 'unterricht',
    'playbook-kursteams': 'unterricht',
    'unterrichtsteams-katalog': 'unterricht',
    'kursteam-einzeln': 'unterricht',
    'kursteam-templates': 'unterricht',
    'onenote-verteilung': 'unterricht',
    diplomarbeiten: 'unterricht',
    'pa-diplom-ordner': 'unterricht',
    spielwiesen: 'unterricht',
    'arge-fachgruppen': 'gruppen',
    klassenchats: 'gruppen',
    datenhygiene: 'hygiene',
    'slg-schueler': 'gruppen',
    'slg-lehrer': 'gruppen',
    verwaltung: 'gruppen',
    klassenvorstaende: 'gruppen',
    'weitere-teams-gruppen': 'gruppen',
    'teams-antrag': 'gruppen',
    'sharepoint-liste-teams-antrag': 'gruppen',
    lizenzverwaltung: 'personen',
    'personen-verwaltung': 'personen',
    'namenskonvention-audit': 'personen',
    'gaeste-verwalten': 'personen',
    'schueler-lifecycle': 'personen',
    'pa-gast-erinnerung': 'personen',
    'playbook-schuljahresstart': 'planung',
    'playbook-daten-import-verknuepfen': 'planung',
    'bildungsportal-stammdaten': 'planung',
    'organisations-assistent': 'planung',
    'klassen-umbenennen': 'planung',
    'webuntis-sync-monitor': 'planung',
    'webuntis-stammdaten-import': 'planung',
    'cleanup-playbook-schuljahr': 'planung',
    postfaecher: 'kommunikation',
    'kalender-postfaecher': 'kommunikation',
    verteilerlisten: 'kommunikation',
    'eltern-verteiler': 'kommunikation',
    'elternsprechtag-bookings': 'kommunikation',
    'playbook-eltern': 'kommunikation',
    'playbook-elternsprechtag': 'kommunikation',
    'raeume-ressourcen': 'kommunikation',
    'sharepoint-intranet-hub': 'intranet',
    'playbook-intranet': 'intranet',
    'playbook-schularbeiten-planer': 'schulapps',
    'sharepoint-liste-lehrer': 'intranet',
    'sharepoint-liste-stammdaten': 'intranet',
    'sharepoint-liste-schultermine': 'intranet',
    'sharepoint-liste-schularbeiten': 'schulapps',
    'sharepoint-liste-schulaktivitaeten': 'schulapps',
    'sharepoint-liste-projektwochen': 'schulapps',
    projektwochen: 'schulapps',
    'sharepoint-liste-srdp': 'intranet',
    'srdp-anmeldung': 'intranet',
    'schularbeiten-planer': 'schulapps',
    'schulaktivitaeten-planer': 'schulapps',
    'freistellung-planer': 'schulapps',
    'lehrer-freistellung-planer': 'schulapps',
    schulsprecherwahl: 'schulapps',
    'wahlfach-zuteilung': 'schulapps',
    'wahlfach-anmeldung': 'schulapps',
    'sharepoint-liste-wahlfach': 'intranet',
    'schueler-stammdaten-portal': 'schulapps',
    einsammler: 'schulapps',
    'sharepoint-liste-einsammler': 'schulapps',
    elternbrief: 'schulapps',
    'sharepoint-liste-elternbrief': 'schulapps',
    krankmeldung: 'schulapps',
    'sharepoint-liste-krankmeldung': 'schulapps',
    'kollegium-umfragen': 'schulapps',
    'sharepoint-liste-kollegium-umfragen': 'schulapps',
    ausgaben: 'schulapps',
    'sharepoint-liste-ausgaben': 'schulapps',
    'sharepoint-liste-aufgaben': 'schulapps',
    'sharepoint-liste-technik-meldungen': 'schulapps',
    'sharepoint-liste-raumplan': 'schulapps',
    'sharepoint-liste-vertretung': 'intranet',
    'termine-hub': 'schulapps',
    aufgaben: 'schulapps',
    'technik-meldungen': 'schulapps',
    raumplan: 'schulapps',
    'pa-erst-setup': 'automationen',
    'power-automate-rezepte': 'automationen',
    'freistellung-setup': 'automationen',
    'schulaktivitaeten-setup': 'automationen',
    'srdp-mail-setup': 'automationen',
    'playbook-freistellungen': 'automationen',
    'pa-termine-sync': 'automationen',
    'pa-schularbeiten-mail': 'automationen',
    'pa-projektwochen-mail': 'automationen',
    'pa-antraege': 'automationen',
    'pa-seminar': 'automationen',
    'pa-schilf': 'automationen',
    gruppenerstellung: 'automationen',
    'sharepoint-mandant-website': 'automationen',
    'sharepoint-mandant-teilen': 'automationen',
    'schul-baseline': 'automationen',
    'schulstruktur-sync': 'hygiene',
    datenlandkarte: 'hygiene',
    'schul-organigramm': 'hygiene',
    schulgraph: 'hygiene',
    'stammdaten-uebergabe': 'hygiene',
    'it-berechtigungen': 'hygiene',
    'stammdaten-backup-abgleich': 'hygiene',
    'cleanup-playbook': 'hygiene',
    'datei-migration': 'hygiene',
    'nextcloud-teams-migrate': 'hygiene',
    'leere-gruppen-report': 'hygiene',
    'teams-archiv': 'hygiene',
    'teams-kanaele-bulk': 'gruppen'
};

/**
 * @returns {Array<{ id: string, label: string, icon: string, toolIds: string[] }>}
 */
export function listDashboardToolsByCluster() {
    const buckets = Object.create(null);
    for (const clusterId of DASHBOARD_CLUSTER_ORDER) {
        buckets[clusterId] = [];
    }
    for (const toolId of Object.keys(DASHBOARD_TOOL_RULES)) {
        const clusterId = DASHBOARD_TOOL_CLUSTER[toolId] || 'hygiene';
        if (!buckets[clusterId]) buckets[clusterId] = [];
        buckets[clusterId].push(toolId);
    }
    return DASHBOARD_CLUSTER_ORDER
        .filter((id) => buckets[id] && buckets[id].length)
        .map((id) => {
            const meta = DASHBOARD_CLUSTER_META[id] || { label: id, icon: 'bi-grid' };
            const toolIds = buckets[id].slice().sort((a, b) =>
                toolLabel(a).localeCompare(toolLabel(b), 'de')
            );
            return { id, label: meta.label, icon: meta.icon, toolIds };
        });
}

/** @typedef {'it'|'lehrer'|'schueler'} DashboardView */

/**
 * @param {DashboardToolRule|undefined} rule
 * @returns {DashboardAudience[]}
 */
export function normalizeAudiences(rule) {
    if (!rule) return ['it'];
    const raw = rule.audience;
    if (Array.isArray(raw)) return raw;
    return [raw];
}

/**
 * @param {string} toolId
 * @param {DashboardView} view
 * @param {{
 *   schularbeitenRoles?: string[],
 *   freistellungRoles?: string[]
 * }} ctx
 */
/**
 * @param {string} toolId
 * @param {{ schularbeitenRoles?: string[], freistellungRoles?: string[] }} ctx
 * @param {DashboardView} view
 * @param {DashboardToolRule} [ruleOverride]
 */
/**
 * Portal-fähige Alltags-Apps (nicht IT-Setup/Hygiene).
 * Whitelist = darf überhaupt ins Lehrer-/Schüler-Portal;
 * effektive Audience (Default + IT-Matrix) entscheidet die Persona-Sichtbarkeit.
 * Siehe docs/lehrer-portal-roadmap.md Phase P0 / Iteration Portal-Katalog.
 */
export const DASHBOARD_LEHRER_SCHUELER_APPS = [
    'schularbeiten-planer',
    'freistellung-planer',
    'lehrer-freistellung-planer',
    'schulaktivitaeten-planer',
    'projektwochen',
    'termine-hub',
    'aufgaben',
    'technik-meldungen',
    'raumplan',
    'elternbrief',
    'einsammler',
    'krankmeldung',
    'kollegium-umfragen',
    'teams-antrag',
    'diplomarbeiten',
    'ausgaben',
    'srdp-anmeldung',
    'kursteams',
    'unterrichtsteams-katalog',
    'kursteam-einzeln',
    'onenote-verteilung',
    'elternsprechtag-bookings',
    'sharepoint-intranet-hub',
    'klassenchats',
    'schulsprecherwahl',
    'wahlfach-zuteilung',
    'wahlfach-anmeldung'
];

/**
 * Gruppierung im Lehrer-/Schüler-Portal (Apps-only-Ansicht).
 * @type {Record<'lehrer'|'schueler', Array<{ id: string, label: string, icon: string, toolIds: string[] }>>}
 */
export const DASHBOARD_PORTAL_SECTIONS_BY_VIEW = {
    lehrer: [
        {
            id: 'unterricht',
            label: 'Unterricht',
            icon: 'bi-mortarboard',
            toolIds: [
                'schularbeiten-planer',
                'freistellung-planer',
                'lehrer-freistellung-planer',
                'diplomarbeiten',
                'kursteams',
                'unterrichtsteams-katalog',
                'kursteam-einzeln',
                'onenote-verteilung',
                'srdp-anmeldung',
                'wahlfach-zuteilung'
            ]
        },
        {
            id: 'schulalltag',
            label: 'Schulalltag',
            icon: 'bi-building',
            toolIds: [
                'termine-hub',
                'schulaktivitaeten-planer',
                'projektwochen',
                'aufgaben',
                'schulsprecherwahl'
            ]
        },
        {
            id: 'betrieb',
            label: 'Betrieb & Technik',
            icon: 'bi-wrench',
            toolIds: ['technik-meldungen', 'raumplan', 'ausgaben']
        },
        {
            id: 'eltern',
            label: 'Eltern & Ablauf',
            icon: 'bi-mailbox',
            toolIds: ['elternbrief', 'einsammler', 'krankmeldung', 'elternsprechtag-bookings']
        },
        {
            id: 'kollegium',
            label: 'Kollegium',
            icon: 'bi-people',
            toolIds: ['kollegium-umfragen', 'teams-antrag', 'klassenchats']
        },
        {
            id: 'intranet',
            label: 'Intranet',
            icon: 'bi-house-door',
            toolIds: ['sharepoint-intranet-hub']
        }
    ],
    schueler: [
        {
            id: 'meine',
            label: 'Meine Apps',
            icon: 'bi-grid',
            toolIds: [
                'termine-hub',
                'schularbeiten-planer',
                'freistellung-planer',
                'srdp-anmeldung',
                'schulsprecherwahl',
                'wahlfach-anmeldung',
                'technik-meldungen',
                'teams-antrag',
                'sharepoint-intranet-hub'
            ]
        }
    ]
};

export function isToolVisibleForView(toolId, ctx, view, ruleOverride) {
    if (view === 'it') return true;

    const rule = ruleOverride || DASHBOARD_TOOL_RULES[toolId];
    const audiences = normalizeAudiences(rule);

    if (view === 'lehrer' || view === 'schueler') {
        // Whitelist = portal-fähig; Audience (Default + IT-Matrix-Override) entscheidet.
        if (!DASHBOARD_LEHRER_SCHUELER_APPS.includes(toolId)) return false;
        if (audiences.includes('all')) return true;
        return audiences.includes(view);
    }

    if (audiences.includes('all')) {
        return passesPlannerGate(rule, ctx);
    }

    const personaOk = audiences.includes(view);
    if (!personaOk) return false;
    return passesPlannerGate(rule, ctx);
}

/**
 * @param {DashboardToolRule|undefined} rule
 * @param {{ schularbeitenRoles?: string[], freistellungRoles?: string[] }} ctx
 */
function passesPlannerGate(rule, ctx) {
    if (!rule || !rule.planner) return true;
    const need = rule.planner.roles || [];
    if (!need.length) return true;
    const have =
        rule.planner.app === 'schularbeiten'
            ? ctx.schularbeitenRoles || []
            : ctx.freistellungRoles || [];
    const set = new Set(have.map((r) => String(r).toLowerCase()));
    return need.some((r) => set.has(String(r).toLowerCase()));
}

/**
 * @param {{ isIt?: boolean, isLehrer?: boolean, isSchueler?: boolean }} personas
 * @returns {DashboardView[]}
 */
export function listAvailableViews(personas) {
    if (personas.isIt) return ['it'];
    if (personas.isLehrer) return ['lehrer'];
    return ['schueler'];
}

/**
 * @param {DashboardView[]} views
 * @param {DashboardView|null|undefined} preferred
 * @returns {DashboardView}
 */
/**
 * @param {DashboardView[]} views
 * @param {DashboardView|null|undefined} preferred
 * @param {{ isIt?: boolean }} [personas]
 */
export function resolveActiveView(views, preferred, personas) {
    const list = views && views.length ? views : ['it'];
    if (preferred && list.includes(preferred)) return preferred;
    if (personas && personas.isIt && list.includes('it')) return 'it';
    if (list.length === 1) return list[0];
    if (list.includes('schueler') && !list.includes('lehrer') && !list.includes('it')) return 'schueler';
    if (list.includes('lehrer') && !list.includes('it')) return 'lehrer';
    if (list.includes('it')) return 'it';
    if (list.includes('lehrer')) return 'lehrer';
    if (list.includes('schueler')) return 'schueler';
    return list[0];
}

export function viewLabel(view) {
    if (view === 'it') return 'Schul-IT';
    if (view === 'lehrer') return 'Lehrkraft';
    if (view === 'schueler') return 'Schüler/in';
    return String(view || '');
}

/** Anzeigenamen für Admin-Matrix und Schnellstart */
export const DASHBOARD_TOOL_LABELS = {
    jahrgang: 'Klassengruppen',
    'klassen-merge': 'Klassen zusammenlegen',
    kursteams: 'Unterrichtsteams',
    'playbook-kursteams': 'Playbook Unterrichtsteams',
    'unterrichtsteams-katalog': 'Unterrichtsteams-Katalog',
    'kursteam-einzeln': 'Einzelne Unterrichtsteams',
    'kursteam-templates': 'Kursteam-Vorlagen',
    'onenote-verteilung': 'OneNote-Inhalte verteilen',
    'arge-fachgruppen': 'Fächer und ARGEs',
    'teams-kanaele-bulk': 'Kanäle bulk anlegen',
    diplomarbeiten: 'Diplomarbeiten / VWA',
    'pa-diplom-ordner': 'Diplom-Ordner Flow',
    'teams-antrag': 'Teams-Antrag',
    'sharepoint-liste-teams-antrag': 'Teams-Antrag-Liste',
    spielwiesen: 'Spielwiesen-Teams',
    klassenchats: 'KlassenChats',
    lizenzverwaltung: 'Lizenzverwaltung',
    'personen-verwaltung': 'Personen suchen',
    'namenskonvention-audit': 'Namenskonvention prüfen',
    'gaeste-verwalten': 'Gäste',
    'schueler-lifecycle': 'Schüler-Lifecycle',
    'pa-gast-erinnerung': 'Gast-Erinnerung Flow',
    'slg-schueler': 'Schüler:innen-Sammelgruppe',
    'slg-lehrer': 'Lehrer:innen-Sammelgruppe',
    verwaltung: 'Schulleitung & Verwaltung',
    klassenvorstaende: 'Klassenvorstände',
    'weitere-teams-gruppen': 'Weitere Teams & Gruppen',
    'playbook-schuljahresstart': 'Playbook Schuljahresstart',
    'playbook-daten-import-verknuepfen': 'Daten importieren & Verknüpfen',
    'bildungsportal-stammdaten': 'Bildungsportal (Vorschau)',
    'organisations-assistent': 'Schuljahr wechseln',
    'klassen-umbenennen': 'Klassen umbenennen',
    'webuntis-sync-monitor': 'WebUntis-Sync-Monitor',
    'webuntis-stammdaten-import': 'Daten importieren',
    'cleanup-playbook-schuljahr': 'Cleanup zum Schuljahr',
    postfaecher: 'Gemeinsame Postfächer',
    'kalender-postfaecher': 'Kalender-Postfächer',
    verteilerlisten: 'E-Mail-Verteiler',
    'eltern-verteiler': 'Eltern-Verteiler',
    'elternsprechtag-bookings': 'Elternsprechtag Bookings',
    'playbook-eltern': 'Playbook Elternkommunikation',
    'playbook-elternsprechtag': 'Playbook Elternsprechtag',
    'raeume-ressourcen': 'Räume und Ressourcen',
    'sharepoint-intranet-hub': 'Schul-Intranet',
    'playbook-intranet': 'Playbook Intranet',
    'playbook-schularbeiten-planer': 'Playbook Schularbeiten-Planer',
    'sharepoint-liste-lehrer': 'Lehrerliste',
    'sharepoint-liste-stammdaten': 'Stammdaten-Listen',
    'sharepoint-liste-schultermine': 'Schultermine-Liste',
    'sharepoint-liste-schularbeiten': 'Schularbeiten-Listen (IT)',
    'sharepoint-liste-schulaktivitaeten': 'Schulaktivitäten-Listen',
    'sharepoint-liste-projektwochen': 'Projektwochen-Listen',
    projektwochen: 'Projektwochen',
    'sharepoint-liste-srdp': 'sRDP-Liste einrichten',
    'srdp-anmeldung': 'Maturaanmeldung (sRDP)',
    'schularbeiten-planer': 'Schularbeiten-Planer',
    'schulaktivitaeten-planer': 'Schulaktivitäten',
    'freistellung-planer': 'Freistellungen (Schüler)',
    'lehrer-freistellung-planer': 'Freistellungen Lehrkräfte',
    schulsprecherwahl: 'Schulwahlen',
    'wahlfach-zuteilung': 'Wahlfach-Zuteilung',
    'wahlfach-anmeldung': 'Wahlfach anmelden',
    'sharepoint-liste-wahlfach': 'Wahlfach-Listen',
    'schueler-stammdaten-portal': 'Schüler-Stammdaten-Portal',
    einsammler: 'Einsammler',
    'sharepoint-liste-einsammler': 'Einsammler-Listen',
    elternbrief: 'Elternbrief',
    'sharepoint-liste-elternbrief': 'Elternbrief-Listen',
    krankmeldung: 'Krankmeldung',
    'sharepoint-liste-krankmeldung': 'Krankmeldungen-Liste',
    'kollegium-umfragen': 'Kollegiumsumfragen',
    'sharepoint-liste-kollegium-umfragen': 'Kollegiumsumfragen-Listen',
    ausgaben: 'Ausgabenverwaltung',
    'sharepoint-liste-ausgaben': 'Ausgaben-Listen',
    'sharepoint-liste-aufgaben': 'Aufgaben-Liste',
    'sharepoint-liste-technik-meldungen': 'Technik-Meldungen-Liste',
    'sharepoint-liste-raumplan': 'Raumplan-Liste',
    'sharepoint-liste-vertretung': 'Vertretungsplan-Liste',
    'termine-hub': 'Kalender und Termine',
    aufgaben: 'Aufgaben',
    'technik-meldungen': 'Technik-Meldungen',
    raumplan: 'Raumplan',
    'pa-erst-setup': 'Power Platform Erst-Setup',
    'power-automate-rezepte': 'Automationen-Übersicht',
    'freistellung-setup': 'Freistellungen Setup',
    'schulaktivitaeten-setup': 'Schulaktivitäten Setup',
    'srdp-mail-setup': 'sRDP Bestätigungsmail',
    'playbook-freistellungen': 'Playbook Freistellungen',
    'pa-termine-sync': 'Termine → Kalender',
    'pa-schularbeiten-mail': 'Schularbeiten-Mail Flow',
    'pa-projektwochen-mail': 'Projektwochen-Mail Flow',
    'pa-antraege': 'Forms-Anträge Flow',
    'pa-seminar': 'Seminar Flow',
    'pa-schilf': 'Schilf-Checkliste',
    gruppenerstellung: 'Wer darf Teams anlegen?',
    'sharepoint-mandant-website': 'Neue Websites',
    'sharepoint-mandant-teilen': 'Dateien teilen',
    'schulstruktur-sync': 'Alle Gruppen und Teams',
    datenhygiene: 'Datenhygiene',
    datenlandkarte: 'Datenlandkarte',
    'schul-organigramm': 'Schul-Organigramm',
    schulgraph: 'Schuldaten-Karte',
    'stammdaten-uebergabe': 'Stammdaten-Übergabe',
    'it-berechtigungen': 'Berechtigungen',
    'stammdaten-backup-abgleich': 'Backup-Abgleich',
    'cleanup-playbook': 'Cleanup-Playbook',
    'schul-baseline': 'Schul-Baseline',
    'datei-migration': 'Datei-Migration',
    'nextcloud-teams-migrate': 'Nextcloud → Teams',
    'leere-gruppen-report': 'Leere Gruppen finden',
    'teams-archiv': 'Teams archivieren'
};

/** Standard-Schnellstart (sichtbare Tools in Reihenfolge) */
export const DASHBOARD_QUICKSTART_BY_VIEW = {
    schueler: [
        'termine-hub',
        'schularbeiten-planer',
        'freistellung-planer',
        'technik-meldungen',
        'srdp-anmeldung',
        'schulsprecherwahl',
        'teams-antrag',
        'sharepoint-intranet-hub'
    ],
    lehrer: [
        'termine-hub',
        'schularbeiten-planer',
        'schulaktivitaeten-planer',
        'freistellung-planer',
        'lehrer-freistellung-planer',
        'srdp-anmeldung',
        'kursteams',
        'aufgaben',
        'technik-meldungen',
        'raumplan',
        'elternbrief',
        'einsammler',
        'krankmeldung',
        'kollegium-umfragen',
        'teams-antrag',
        'diplomarbeiten',
        'wahlfach-zuteilung',
        'wahlfach-anmeldung',
        'schulsprecherwahl',
        'ausgaben',
        'onenote-verteilung',
        'elternsprechtag-bookings',
        'klassenchats',
        'sharepoint-intranet-hub'
    ]
};

/** @typedef {import('./dashboard-audience-catalog.js').DashboardView} DashboardView */

export const DASHBOARD_TOOL_LINKS = {
    'schularbeiten-planer': 'tools/schularbeiten-planer.html',
    'freistellung-planer': 'tools/freistellung-planer.html',
    'lehrer-freistellung-planer': 'tools/lehrer-freistellung-planer.html',
    'schulaktivitaeten-planer': 'tools/schulaktivitaeten-planer.html',
    'termine-hub': 'tools/termine-hub.html',
    aufgaben: 'tools/aufgaben.html',
    'technik-meldungen': 'tools/technik-meldungen.html',
    raumplan: 'tools/raumplan.html',
    elternbrief: 'tools/elternbrief.html',
    einsammler: 'tools/einsammler.html',
    krankmeldung: 'tools/krankmeldung.html',
    'kollegium-umfragen': 'tools/kollegium-umfragen.html',
    'teams-antrag': 'tools/teams-antrag.html',
    diplomarbeiten: 'tools/diplomarbeiten.html',
    'srdp-anmeldung': 'tools/srdp-anmeldung.html',
    kursteams: 'tools/kursteams.html',
    'unterrichtsteams-katalog': 'tools/unterrichtsteams-katalog.html',
    'kursteam-einzeln': 'tools/kursteams.html?mode=single',
    'onenote-verteilung': 'tools/onenote-verteilung.html',
    'elternsprechtag-bookings': 'tools/elternsprechtag-bookings.html',
    'sharepoint-intranet-hub': 'tools/sharepoint-intranet-hub.html',
    klassenchats: 'tools/klassenchats.html',
    schulsprecherwahl: 'tools/schulsprecherwahl.html',
    'wahlfach-zuteilung': 'tools/wahlfach-zuteilung.html',
    'wahlfach-anmeldung': 'tools/wahlfach-anmeldung.html',
    ausgaben: 'tools/ausgaben.html',
    projektwochen: 'tools/projektwochen.html'
};

export const DASHBOARD_TOOL_ICONS = {
    'schularbeiten-planer': 'bi-journal-check',
    'freistellung-planer': 'bi-calendar2-check',
    'lehrer-freistellung-planer': 'bi-briefcase',
    'schulaktivitaeten-planer': 'bi-geo-alt',
    'termine-hub': 'bi-calendar-event',
    aufgaben: 'bi-check2-square',
    'technik-meldungen': 'bi-wrench-adjustable',
    raumplan: 'bi-map',
    elternbrief: 'bi-envelope-paper',
    einsammler: 'bi-ui-checks',
    krankmeldung: 'bi-thermometer-half',
    'kollegium-umfragen': 'bi-clipboard2-check',
    'teams-antrag': 'bi-people-fill',
    diplomarbeiten: 'bi-mortarboard',
    'srdp-anmeldung': 'bi-award',
    kursteams: 'bi-collection',
    'unterrichtsteams-katalog': 'bi-grid-3x3-gap',
    'kursteam-einzeln': 'bi-bounding-box',
    'onenote-verteilung': 'bi-journal-richtext',
    'elternsprechtag-bookings': 'bi-calendar2-week',
    'sharepoint-intranet-hub': 'bi-house-door',
    klassenchats: 'bi-chat-dots',
    schulsprecherwahl: 'bi-ui-checks-grid',
    'wahlfach-zuteilung': 'bi-diagram-3',
    'wahlfach-anmeldung': 'bi-pencil-square',
    ausgaben: 'bi-key',
    projektwochen: 'bi-calendar-range'
};

export const DASHBOARD_TOOL_BLURBS = {
    'schularbeiten-planer': 'Termine, Anträge und Kalender für Ihre Klasse.',
    'freistellung-planer': 'Freistellungen beantragen und Status verfolgen (Schüler/KV).',
    'lehrer-freistellung-planer':
        'Freistellungen für Lehrkräfte – Direktion genehmigt, Kalender und iCal.',
    'schulaktivitaeten-planer': 'Exkursionen und Schulaktivitäten beantragen und verfolgen.',
    'termine-hub':
        'Schul-, Lehrer- und Schülerkalender aus Outlook – rollenbasiert, Woche/Monat, anlegen und pflegen.',
    aufgaben: 'Aufgaben an Personen oder Rollen zuweisen und abhaken.',
    'technik-meldungen':
        'Defekte oder problematische Technik melden – Lehrkräfte und Schüler; IT bearbeitet die Inbox.',
    raumplan: 'Räume suchen, Plan ansehen, CSV importieren.',
    elternbrief: 'Briefe an Eltern mit Lesebestätigung.',
    einsammler: 'Einwilligungen und Rückläufe digital einsammeln.',
    krankmeldung: 'Krankmeldungen der Klasse einsehen und bearbeiten.',
    'kollegium-umfragen': 'Umfragen nur für Lehrkräfte – Antworten und Auswertung.',
    'teams-antrag': 'Neues Team beantragen – IT prüft und legt es an.',
    diplomarbeiten: 'Abschlussarbeiten: Jahrgangs-Team und private Kanäle.',
    'srdp-anmeldung': 'Matura-/Diplomprüfungsanmeldung (sRDP) für Abschlussjahrgänge.',
    kursteams: 'Unterrichtsteams anlegen und Mitglieder steuern.',
    'unterrichtsteams-katalog': 'Übersicht aller Unterrichtsteams der Schule.',
    'kursteam-einzeln': 'Einzelnes Unterrichtsteam anlegen oder nachziehen.',
    'onenote-verteilung': 'OneNote-Inhalte an Klassen und Kurse verteilen.',
    'elternsprechtag-bookings': 'Elternsprechtag mit Bookings organisieren.',
    'sharepoint-intranet-hub': 'Schul-Intranet öffnen und Navigation nutzen.',
    klassenchats: 'KlassenChats für den Alltag anlegen und pflegen.',
    schulsprecherwahl: 'Schulwahlen digital durchführen und auswerten.',
    'wahlfach-zuteilung': 'Wahlfächer erfassen, zuteilen und Übersicht behalten.',
    'wahlfach-anmeldung': '1.–3. Wunsch für Wahlfächer eintragen.',
    ausgaben: 'Schlüssel, Karten und Zugänge verwalten.',
    projektwochen: 'Projektwochen planen, Gruppen und Termine steuern.'
};

/**
 * @returns {string[]}
 */
export function listDashboardToolIds() {
    return Object.keys(DASHBOARD_TOOL_RULES).sort((a, b) => {
        const la = DASHBOARD_TOOL_LABELS[a] || a;
        const lb = DASHBOARD_TOOL_LABELS[b] || b;
        return la.localeCompare(lb, 'de');
    });
}

export function toolLabel(toolId) {
    return DASHBOARD_TOOL_LABELS[toolId] || toolId;
}
