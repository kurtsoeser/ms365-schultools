/**
 * Dashboard „Alle Werkzeuge“ – gleiche Master-Detail-Shell wie „Was möchten Sie tun?“.
 */

import {
    TAB_ICON_BY_PANEL,
    TAB_LABEL_BY_PANEL
} from './dashboard-catalog-layout.js';
import { wireDashSplitNavResize } from './dashboard-tasks-split.js';

export const DASH_CATALOG_SPLIT_NAV_WIDTH_KEY = 'ms365-dash-catalog-split-nav-width-v1';

/**
 * @param {HTMLElement} btn
 * @param {string} panelId
 */
function enhanceCatalogTabButton(btn, panelId) {
    const icon = TAB_ICON_BY_PANEL[panelId] || 'bi-grid';
    const fullLabel = TAB_LABEL_BY_PANEL[panelId] || panelId;
    let indexText = '';
    let labelText = fullLabel;
    const numbered = /^(\d{2})\s+(.+)$/.exec(fullLabel);
    if (numbered) {
        indexText = numbered[1];
        labelText = numbered[2];
    }

    btn.classList.remove('dashboard-tab');
    btn.classList.add('dash-tasks-split__nav-item');

    btn.innerHTML =
        '<span class="dash-tasks-split__nav-row">' +
        '<span class="dash-tasks-split__nav-icon" aria-hidden="true"><i class="bi ' +
        icon +
        '"></i></span>' +
        '<span class="dash-tasks-split__nav-body">' +
        '<span class="dash-tasks-split__nav-line">' +
        (indexText
            ? '<span class="dash-tasks-split__nav-index">' + indexText + '</span>'
            : '<span class="dash-tasks-split__nav-index" hidden></span>') +
        '<span class="dash-tasks-split__nav-label">' +
        labelText +
        '</span>' +
        '</span></span></span>';
}

/** Sync aria-current auf Sidebar-Buttons (parallel zu aria-selected). */
export function syncCatalogSplitNavCurrent() {
    document.querySelectorAll('#dash-catalog [data-dashboard-tab]').forEach(function (btn) {
        const on = btn.getAttribute('aria-selected') === 'true';
        if (on) btn.setAttribute('aria-current', 'true');
        else btn.removeAttribute('aria-current');
    });
}

function relocateCatalogPersonalTools(catalog) {
    const personal = document.getElementById('dashPersonalTools');
    const nav = catalog.querySelector('.dash-tasks-split__nav');
    if (!personal || !nav || personal.parentElement === catalog) return;
    catalog.insertBefore(personal, nav);
    personal.classList.add('dash-catalog-split__personal');
}

export function mountDashboardCatalogSplit() {
    const catalog = document.getElementById('dash-catalog');
    if (!catalog) return false;

    if (catalog.dataset.catalogSplit === '1') {
        relocateCatalogPersonalTools(catalog);
        syncCatalogSplitNavCurrent();
        return true;
    }

    const tablist = catalog.querySelector('.dashboard-tablist');
    const panels = catalog.querySelector('.dashboard-catalog-panels');
    if (!tablist || !panels) return false;

    catalog.classList.add('dash-tasks-split', 'dash-catalog-split');
    catalog.dataset.catalogSplit = '1';

    tablist.classList.remove('dashboard-tablist');
    tablist.classList.add('dash-tasks-split__nav');

    panels.classList.add('dash-tasks-split__detail');

    tablist.querySelectorAll('[data-dashboard-tab]').forEach(function (btn) {
        const panelId = btn.getAttribute('data-dashboard-tab') || '';
        enhanceCatalogTabButton(btn, panelId);
    });

    wireDashSplitNavResize(catalog, {
        storageKey: DASH_CATALOG_SPLIT_NAV_WIDTH_KEY,
        ariaLabel: 'Breite der Katalog-Navigation anpassen'
    });

    relocateCatalogPersonalTools(catalog);

    document.body.classList.add('dash-catalog-split-active');
    syncCatalogSplitNavCurrent();
    return true;
}

if (typeof document !== 'undefined') {
    document.addEventListener('ms365-dash-catalog-ready', function () {
        mountDashboardCatalogSplit();
    });
}
