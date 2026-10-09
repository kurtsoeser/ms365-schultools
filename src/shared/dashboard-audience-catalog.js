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
    diplomarbeiten: { audience: 'it' },
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
    'schularbeiten-planer': {
        audience: ['lehrer', 'schueler'],
        planner: { app: 'schularbeiten', roles: ['admin', 'lehrer', 'schueler'] }
    },
    'schulaktivitaeten-planer': {
        audience: 'lehrer',
        planner: { app: 'schularbeiten', roles: ['admin', 'lehrer'] }
    },
    'freistellung-planer': {
        audience: ['lehrer', 'schueler'],
        planner: { app: 'freistellung', roles: ['direktion', 'kv', 'schueler'] }
    },
    'lehrer-freistellung-planer': {
        audience: 'lehrer',
        planner: { app: 'schularbeiten', roles: ['admin', 'lehrer'] }
    },
    'sharepoint-liste-vertretung': { audience: 'it' },

    // —— Automationen ——
    'pa-erst-setup': { audience: 'it' },
    'power-automate-rezepte': { audience: 'it' },
    'freistellung-setup': { audience: 'it' },
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
    schulgraph: { audience: 'it' },
    'stammdaten-uebergabe': { audience: 'it' },
    'stammdaten-backup-abgleich': { audience: 'it' },
    'cleanup-playbook': { audience: 'it' },
    'schul-baseline': { audience: 'it' },
    'datei-migration': { audience: 'it' },
    'leere-gruppen-report': { audience: 'it' },
    'teams-archiv': { audience: 'it' }
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
    intranet: { label: '05 Intranet & Schulalltag', icon: 'bi-house-door', panel: 'intranet' },
    hygiene: { label: '06 Aufräumen & Audit', icon: 'bi-broom', panel: 'kommunikation' },
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
    verteilerlisten: 'kommunikation',
    'eltern-verteiler': 'kommunikation',
    'elternsprechtag-bookings': 'kommunikation',
    'playbook-eltern': 'kommunikation',
    'playbook-elternsprechtag': 'kommunikation',
    'raeume-ressourcen': 'kommunikation',
    'sharepoint-intranet-hub': 'intranet',
    'playbook-intranet': 'intranet',
    'playbook-schularbeiten-planer': 'intranet',
    'sharepoint-liste-lehrer': 'intranet',
    'sharepoint-liste-stammdaten': 'intranet',
    'sharepoint-liste-schultermine': 'intranet',
    'sharepoint-liste-schularbeiten': 'intranet',
    'sharepoint-liste-schulaktivitaeten': 'intranet',
    'sharepoint-liste-projektwochen': 'intranet',
    projektwochen: 'intranet',
    'sharepoint-liste-srdp': 'intranet',
    'schularbeiten-planer': 'intranet',
    'schulaktivitaeten-planer': 'intranet',
    'freistellung-planer': 'intranet',
    'lehrer-freistellung-planer': 'intranet',
    'sharepoint-liste-vertretung': 'intranet',
    'pa-erst-setup': 'automationen',
    'power-automate-rezepte': 'automationen',
    'freistellung-setup': 'automationen',
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
    schulgraph: 'hygiene',
    'stammdaten-uebergabe': 'hygiene',
    'stammdaten-backup-abgleich': 'hygiene',
    'cleanup-playbook': 'hygiene',
    'datei-migration': 'hygiene',
    'leere-gruppen-report': 'hygiene',
    'teams-archiv': 'hygiene'
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
/** Nur diese zwei Apps im Dashboard für Lehrkraft und Schüler/in */
export const DASHBOARD_LEHRER_SCHUELER_APPS = ['schularbeiten-planer', 'freistellung-planer'];

export function isToolVisibleForView(toolId, ctx, view, ruleOverride) {
    if (view === 'it') return true;

    if (view === 'lehrer' || view === 'schueler') {
        return DASHBOARD_LEHRER_SCHUELER_APPS.includes(toolId);
    }

    const rule = ruleOverride || DASHBOARD_TOOL_RULES[toolId];
    const audiences = normalizeAudiences(rule);

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
    diplomarbeiten: 'Diplomarbeiten',
    'pa-diplom-ordner': 'Diplom-Ordner Flow',
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
    'sharepoint-liste-srdp': 'sRDP-Anmeldung',
    'schularbeiten-planer': 'Schularbeiten-Planer',
    'schulaktivitaeten-planer': 'Schulaktivitäten',
    'freistellung-planer': 'Freistellungen (Schüler)',
    'lehrer-freistellung-planer': 'Freistellungen Lehrkräfte',
    'sharepoint-liste-vertretung': 'Vertretungsplan-Liste',
    'pa-erst-setup': 'Power Platform Erst-Setup',
    'power-automate-rezepte': 'Automationen-Übersicht',
    'freistellung-setup': 'Freistellungen Setup',
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
    schulgraph: 'Schuldaten-Karte',
    'stammdaten-uebergabe': 'Stammdaten-Übergabe',
    'stammdaten-backup-abgleich': 'Backup-Abgleich',
    'cleanup-playbook': 'Cleanup-Playbook',
    'schul-baseline': 'Schul-Baseline',
    'datei-migration': 'Datei-Migration',
    'leere-gruppen-report': 'Leere Gruppen finden',
    'teams-archiv': 'Teams archivieren'
};

/** Standard-Schnellstart (sichtbare Tools in Reihenfolge) */
export const DASHBOARD_QUICKSTART_BY_VIEW = {
    schueler: ['schularbeiten-planer', 'freistellung-planer'],
    lehrer: ['schularbeiten-planer', 'freistellung-planer', 'lehrer-freistellung-planer']
};

/** @typedef {import('./dashboard-audience-catalog.js').DashboardView} DashboardView */

export const DASHBOARD_TOOL_LINKS = {
    'schularbeiten-planer': 'tools/schularbeiten-planer.html',
    'freistellung-planer': 'tools/freistellung-planer.html',
    'lehrer-freistellung-planer': 'tools/lehrer-freistellung-planer.html'
};

export const DASHBOARD_TOOL_ICONS = {
    'schularbeiten-planer': 'bi-journal-check',
    'freistellung-planer': 'bi-calendar2-check',
    'lehrer-freistellung-planer': 'bi-briefcase'
};

export const DASHBOARD_TOOL_BLURBS = {
    'schularbeiten-planer': 'Termine, Anträge und Kalender für Ihre Klasse.',
    'freistellung-planer': 'Freistellungen beantragen und Status verfolgen (Schüler/KV).',
    'lehrer-freistellung-planer':
        'Freistellungen für Lehrkräfte – Direktion genehmigt, Kalender und iCal.'
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
