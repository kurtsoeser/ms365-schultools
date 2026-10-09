/**
 * Werkzeugkatalog (index.html): Tab-Labels, Sidebar und Platzierung nach Cluster.
 */
import {
    DASHBOARD_CLUSTER_META,
    DASHBOARD_CLUSTER_ORDER,
    DASHBOARD_TOOL_CLUSTER
} from './dashboard-audience-catalog.js';
import { enhanceCatalogToolCards } from './dashboard-tool-card.js';

/** Panel-IDs in der Sidebar (entspricht dash-panel-* im DOM). */
export const DASHBOARD_CATALOG_TAB_IDS = [
    'gruppen',
    'unterricht',
    'personen',
    'schuljahr',
    'intranet',
    'schulapps',
    'kommunikation',
    'automationen'
];

/** @type {Record<string, string>} */
export const TAB_LABEL_BY_PANEL = {
    gruppen: '01 Mitgliedschaften',
    unterricht: '02 Klassen & Unterricht',
    personen: '03 Personen & Gäste',
    schuljahr: '04 Schuljahresstart',
    intranet: '05 Intranet',
    schulapps: '06 Schul-Apps',
    kommunikation: '07 Aufräumen & Audit',
    automationen: 'Automatisieren'
};

/** @type {Record<string, string>} */
export const TAB_ICON_BY_PANEL = {
    gruppen: 'bi-people',
    unterricht: 'bi-mortarboard',
    personen: 'bi-person-badge',
    schuljahr: 'bi-rocket-takeoff',
    intranet: 'bi-house-door',
    schulapps: 'bi-window-stack',
    kommunikation: 'bi-broom',
    automationen: 'bi-lightning-charge'
};

/** Cluster-ID → Panel-ID (Tab-Inhalt). */
function panelIdForCluster(clusterId) {
    const meta = DASHBOARD_CLUSTER_META[clusterId];
    return (meta && meta.panel) || clusterId;
}

/**
 * @param {string} tab
 */
export function normalizeDashboardCatalogTab(tab) {
    const t = String(tab || '').trim();
    if (t === 'regeln' || t === 'uebersicht') return 'kommunikation';
    if (t === 'website') return 'intranet';
    if (DASHBOARD_CATALOG_TAB_IDS.includes(t)) return t;
    return 'gruppen';
}

/**
 * @param {HTMLElement} catalog
 */
/**
 * @param {HTMLElement} catalog
 */
export function ensureIntranetKommunikationSection(catalog) {
    const panel = catalog.querySelector('#dash-panel-intranet');
    if (!panel || panel.querySelector('[data-cluster-grid="kommunikation"]')) return;
    const cluster = panel.querySelector('.dashboard-cluster');
    if (!cluster) return;
    const h4 = document.createElement('h4');
    h4.className = 'dashboard-cluster-subhead';
    h4.innerHTML = '<i class="bi bi-envelope" aria-hidden="true"></i>Kommunikation';
    const lead = document.createElement('p');
    lead.className = 'dashboard-cluster-sublead';
    lead.textContent = 'Postfächer, Verteiler und Elternkommunikation.';
    const grid = document.createElement('div');
    grid.className = 'grid';
    grid.setAttribute('data-cluster-grid', 'kommunikation');
    grid.setAttribute('role', 'list');
    grid.setAttribute('aria-label', 'Kommunikation');
    cluster.appendChild(h4);
    cluster.appendChild(lead);
    cluster.appendChild(grid);
}

function gridForTool(catalog, toolId, clusterId) {
    const panelId = panelIdForCluster(clusterId);
    const panel = catalog.querySelector('#dash-panel-' + panelId);
    if (!panel) return null;

    if (clusterId === 'kommunikation') {
        ensureIntranetKommunikationSection(catalog);
        return panel.querySelector('[data-cluster-grid="kommunikation"]');
    }

    if (clusterId === 'intranet') {
        const card = catalog.querySelector('.choice[data-tool-id="' + toolId + '"]');
        const existing = card ? card.closest('[data-cluster-grid^="intranet"]') : null;
        if (existing && panel.contains(existing)) return existing;
        return panel.querySelector('[data-cluster-grid="intranet"]');
    }

    if (clusterId === 'schulapps') {
        const nutzenIds = [
            'projektwochen',
            'schularbeiten-planer',
            'schulaktivitaeten-planer',
            'freistellung-planer',
            'lehrer-freistellung-planer'
        ];
        const gridKey = nutzenIds.indexOf(toolId) >= 0 ? 'schulapps-nutzen' : 'schulapps';
        const card = catalog.querySelector('.choice[data-tool-id="' + toolId + '"]');
        const existing = card ? card.closest('[data-cluster-grid^="schulapps"]') : null;
        if (existing && panel.contains(existing)) return existing;
        return panel.querySelector('[data-cluster-grid="' + gridKey + '"]');
    }

    let grid = panel.querySelector('[data-cluster-grid="' + clusterId + '"]');
    if (!grid) grid = panel.querySelector('[data-cluster-grid]');
    return grid;
}

/**
 * @param {HTMLElement} catalog
 */
export function relocateCatalogTools(catalog) {
    const cards = Array.from(catalog.querySelectorAll('.choice[data-tool-id]'));
    for (const card of cards) {
        const toolId = String(card.getAttribute('data-tool-id') || '').trim();
        if (!toolId) continue;
        const clusterId = DASHBOARD_TOOL_CLUSTER[toolId] || 'hygiene';
        const grid = gridForTool(catalog, toolId, clusterId);
        if (grid && card.parentElement !== grid) {
            grid.appendChild(card);
        }
        card.setAttribute('data-cluster', clusterId);
    }
}

/**
 * @param {HTMLElement} catalog
 */
export function applyCatalogTabLabels(catalog) {
    DASHBOARD_CATALOG_TAB_IDS.forEach(function (panelId) {
        const btn = catalog.querySelector('[data-dashboard-tab="' + panelId + '"]');
        if (!btn) return;
        const icon = TAB_ICON_BY_PANEL[panelId] || 'bi-grid';
        const label = TAB_LABEL_BY_PANEL[panelId] || panelId;
        btn.innerHTML = '<i class="bi ' + icon + '" aria-hidden="true"></i>' + label;
    });
}

/**
 * @param {HTMLElement} catalog
 */
export function removeLegacyUebersichtCatalogTab(catalog) {
    const tabBtn = catalog.querySelector('[data-dashboard-tab="uebersicht"]');
    if (tabBtn) tabBtn.remove();
    const panel = catalog.querySelector('#dash-panel-uebersicht');
    if (panel) panel.remove();
}

/**
 * @param {HTMLElement} catalog
 */
/** @type {Record<string, { title: string, description: string }>} */
const MINIMAL_CATALOG_CARD_COPY = {
    'pa-termine-sync': {
        title: 'Termine',
        description: 'Schultermine zwischen SharePoint-Listen und Outlook-Kalendern synchronisieren – in Entwicklung.'
    },
    'pa-schularbeiten-mail': {
        title: 'SA Mail',
        description: 'Automatische E-Mail-Benachrichtigung bei Schularbeitsterminen – in Entwicklung.'
    },
    'pa-projektwochen-mail': {
        title: 'PW Mail',
        description: 'Status-E-Mails für Projektwochen-Anmeldungen – in Entwicklung.'
    },
    'pa-antraege': {
        title: 'Anträge',
        description: 'Microsoft-Forms-Anträge automatisch weiterverarbeiten – in Entwicklung.'
    },
    'pa-seminar': {
        title: 'Seminar',
        description: 'Fortbildungs- und Seminar-Anmeldungen per Flow – in Entwicklung.'
    }
};

/**
 * @param {HTMLElement} catalog
 */
export function enrichMinimalCatalogCards(catalog) {
    if (!catalog) return;
    catalog.querySelectorAll('.choice[data-tool-id]').forEach(function (card) {
        if (card.querySelector('h2, .card-title')) return;
        const toolId = String(card.getAttribute('data-tool-id') || '').trim();
        const link = card.querySelector('a[href]');
        if (!link) return;
        const copy = MINIMAL_CATALOG_CARD_COPY[toolId];
        const title = (copy && copy.title) || String(link.textContent || toolId).trim();
        const description =
            (copy && copy.description) ||
            'Geplantes Werkzeug – Inhalt folgt in einer der nächsten Versionen.';
        const href = link.getAttribute('href') || '#';
        card.classList.add('tool-card--coming-soon');
        card.setAttribute('aria-label', title);
        card.innerHTML =
            '<h2><i class="bi bi-hourglass-split" aria-hidden="true"></i>' +
            title +
            '</h2>' +
            '<p>' +
            description +
            '</p>' +
            '<span class="card-coming-soon-badge">In Entwicklung</span>' +
            '<a class="btn" href="' +
            href +
            '" tabindex="-1" aria-hidden="true" hidden>Öffnen</a>';
    });
}

export function initDashboardCatalogLayout(catalog) {
    if (!catalog) return;
    ensureIntranetKommunikationSection(catalog);
    relocateCatalogTools(catalog);
    enrichMinimalCatalogCards(catalog);
    enhanceCatalogToolCards(catalog);
    applyCatalogTabLabels(catalog);
    removeLegacyUebersichtCatalogTab(catalog);
}

export { DASHBOARD_CLUSTER_ORDER };
