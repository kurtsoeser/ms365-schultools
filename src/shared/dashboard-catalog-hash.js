/**
 * UI-Spec §7.1 – Kategorie-Hash für Aufgaben- und Katalog-Ansicht.
 */
import { normalizeDashboardCatalogTab } from './dashboard-catalog-layout.js';

/** @type {Record<string, string>} */
export const CATALOG_PANEL_HASH = {
    gruppen: 'mitgliedschaften',
    unterricht: 'klassen',
    personen: 'personen',
    schuljahr: 'schuljahresstart',
    intranet: 'intranet',
    kommunikation: 'aufraeumen',
    automationen: 'automationen'
};

/** @type {Record<string, string>} */
const HASH_TO_PANEL = (function () {
    const map = {};
    Object.keys(CATALOG_PANEL_HASH).forEach(function (panel) {
        map[CATALOG_PANEL_HASH[panel]] = panel;
    });
    map.import = 'gruppen';
    return map;
})();

/** @type {Record<string, string>} */
const TASK_ID_BY_PANEL = {
    gruppen: 'dashTaskGruppen',
    unterricht: 'dashTaskUnterricht',
    personen: 'dashTaskPersonen',
    schuljahr: 'dashTaskSchuljahr',
    intranet: 'dashTaskIntranet',
    kommunikation: 'dashTaskOrdnung'
};

/**
 * @param {string} [hash]
 */
export function catalogPanelFromHash(hash) {
    const h = String(hash != null ? hash : window.location.hash || '')
        .replace(/^#/, '')
        .trim()
        .toLowerCase();
    if (!h) return '';
    const panel = HASH_TO_PANEL[h];
    return panel ? normalizeDashboardCatalogTab(panel) : '';
}

/**
 * @param {string} panelId
 */
export function hashForCatalogPanel(panelId) {
    const p = normalizeDashboardCatalogTab(panelId);
    return CATALOG_PANEL_HASH[p] || '';
}

function readMainSection() {
    try {
        return document.documentElement.getAttribute('data-dash-main-section') || 'tasks';
    } catch {
        return 'tasks';
    }
}

/**
 * @param {string} hash
 * @param {{ skipHash?: boolean }} [opts]
 */
export function applyDashboardCategoryHash(hash, opts) {
    const panel = catalogPanelFromHash(hash);
    if (!panel) return false;
    const section = readMainSection();
    const options = opts && typeof opts === 'object' ? opts : {};

    if (section === 'catalog' && typeof window.__ms365DashCatalogSetActiveTab === 'function') {
        window.__ms365DashCatalogSetActiveTab(panel, { fromHash: true });
        return true;
    }
    if (section === 'tasks') {
        const importFirst = panel === 'gruppen' && String(hash || '').replace(/^#/, '') === 'import';
        const taskId = importFirst ? 'dashTaskImportVerknuepfen' : TASK_ID_BY_PANEL[panel];
        if (taskId && window.ms365DashTasksSplit && typeof window.ms365DashTasksSplit.select === 'function') {
            window.ms365DashTasksSplit.select(taskId, { skipHash: !!options.skipHash, pulse: false });
            return true;
        }
    }
    return false;
}

/**
 * @param {{
 *   setActiveTab: (tab: string, opts?: object) => void,
 *   isKnownTab: (tab: string) => boolean
 * }} bind
 */
export function wireCatalogHashNavigation(bind) {
    if (!bind || typeof bind.setActiveTab !== 'function') return;
    if (window.__ms365DashCatalogHashWired) return;
    window.__ms365DashCatalogHashWired = true;

    const baseSet = bind.setActiveTab;

    function setActiveTab(tab, opts) {
        const normalized = normalizeDashboardCatalogTab(tab);
        if (!bind.isKnownTab(normalized)) {
            baseSet(tab, opts);
            return;
        }
        baseSet(normalized, opts);
        const fromHash = opts && opts.fromHash;
        if (fromHash || readMainSection() !== 'catalog') return;
        const want = hashForCatalogPanel(normalized);
        if (!want) return;
        try {
            if (window.location.hash.replace(/^#/, '') !== want) {
                window.history.replaceState(null, '', '#' + want);
            }
        } catch {
            /* ignore */
        }
    }

    window.__ms365DashCatalogSetActiveTab = setActiveTab;

    function onHashOrSection() {
        applyDashboardCategoryHash(window.location.hash, { skipHash: true });
    }

    window.addEventListener('hashchange', onHashOrSection);
    document.addEventListener('ms365-dash-main-section-changed', onHashOrSection);

    const initial = catalogPanelFromHash(window.location.hash);
    if (initial && bind.isKnownTab(initial) && readMainSection() === 'catalog') {
        setActiveTab(initial, { fromHash: true });
    }
}

if (typeof document !== 'undefined') {
    function tryWire() {
        if (typeof window.__ms365DashCatalogBindHash !== 'function') return false;
        wireCatalogHashNavigation(window.__ms365DashCatalogBindHash());
        return true;
    }
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function () {
            if (!tryWire()) {
                document.addEventListener('ms365-dash-catalog-ready', tryWire);
            }
        });
    } else if (!tryWire()) {
        document.addEventListener('ms365-dash-catalog-ready', tryWire);
    }
}
