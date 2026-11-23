/**
 * Dashboard „Alle Werkzeuge“ – gleiche Master-Detail-Shell wie „Was möchten Sie tun?“.
 */

import {
    TAB_ICON_BY_PANEL,
    TAB_LABEL_BY_PANEL
} from './dashboard-catalog-layout.js';
import { wireDashSplitNavResize } from './dashboard-tasks-split.js';
import {
    applyToolsViewMode,
    ensureCatalogInsightTile,
    mountToolsViewToggle,
    readToolsViewMode
} from './dashboard-tools-view.js';
import {
    isDashboardItShellActive,
    persistItWorkspaceNav
} from './it-workspace-nav.js';

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

/**
 * @param {HTMLElement} catalog
 * @param {HTMLElement} tablist
 */
function mountCatalogNavSubtrees(catalog, tablist) {
    if (tablist.dataset.catalogSubs === '1') return;
    tablist.dataset.catalogSubs = '1';

    const tree = [];
    const buttons = Array.from(tablist.querySelectorAll('[data-dashboard-tab]'));
    buttons.forEach(function (btn) {
        const panelId = btn.getAttribute('data-dashboard-tab') || '';
        const panel = catalog.querySelector('#dash-panel-' + panelId);
        const block = document.createElement('div');
        block.className = 'dash-tasks-split__nav-block';
        block.dataset.catalogPanel = panelId;
        btn.parentNode.insertBefore(block, btn);
        block.appendChild(btn);

        const tools = [];
        if (panel) {
            panel.querySelectorAll('.choice[data-tool-id]').forEach(function (card) {
                const link = card.querySelector('a.btn[href], a.card-btn[href], a[href*=".html"]');
                if (!link) return;
                const href = String(link.getAttribute('href') || '').trim();
                if (!href) return;
                const titleEl = card.querySelector('h2, .card-title');
                let label = '';
                let icon = 'bi bi-box-arrow-up-right';
                if (titleEl) {
                    const ic = titleEl.querySelector('i.bi, .bi');
                    if (ic && ic.className) icon = String(ic.className);
                    const clone = titleEl.cloneNode(true);
                    clone.querySelectorAll('i, .bi').forEach(function (n) {
                        n.remove();
                    });
                    label = String(clone.textContent || '')
                        .replace(/\s+/g, ' ')
                        .trim();
                }
                if (!label) label = card.getAttribute('aria-label') || href;
                tools.push({ href: href, label: label, icon: icon });
            });
        }

        tree.push({
            id: 'catalog-' + panelId,
            label: TAB_LABEL_BY_PANEL[panelId] || panelId,
            icon: 'bi ' + (TAB_ICON_BY_PANEL[panelId] || 'bi-grid'),
            tools: tools
        });

        if (!tools.length) return;
        const sub = document.createElement('div');
        sub.className = 'dash-tasks-split__nav-sub';
        sub.hidden = btn.getAttribute('aria-selected') !== 'true';
        tools.forEach(function (tool) {
            const a = document.createElement('a');
            a.className = 'dash-tasks-split__nav-subitem';
            a.href = tool.href;
            a.innerHTML =
                '<span class="dash-tasks-split__nav-subicon" aria-hidden="true"><i class="' +
                tool.icon +
                '"></i></span>' +
                '<span class="dash-tasks-split__nav-sublabel"></span>';
            a.querySelector('.dash-tasks-split__nav-sublabel').textContent = tool.label;
            a.addEventListener('click', function () {
                if (isDashboardItShellActive()) {
                    persistItWorkspaceNav(tree, 'catalog-' + panelId);
                }
            });
            sub.appendChild(a);
        });
        block.appendChild(sub);
    });

    catalog.__ms365CatalogNavTree = tree;
}

/** Sync aria-current auf Sidebar-Buttons (parallel zu aria-selected). */
export function syncCatalogSplitNavCurrent() {
    document.querySelectorAll('#dash-catalog [data-dashboard-tab]').forEach(function (btn) {
        const on = btn.getAttribute('aria-selected') === 'true';
        if (on) btn.setAttribute('aria-current', 'true');
        else btn.removeAttribute('aria-current');
        const block = btn.closest('.dash-tasks-split__nav-block');
        const sub = block && block.querySelector('.dash-tasks-split__nav-sub');
        if (sub) sub.hidden = !on;
        if (block) block.classList.toggle('is-active', on);
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
        const detailReady = catalog.querySelector('.dash-tasks-split__detail');
        if (detailReady) mountToolsViewToggle(detailReady, { ariaLabel: 'Darstellung des Werkzeugkatalogs' });
        applyToolsViewMode(readToolsViewMode());
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
    mountCatalogNavSubtrees(catalog, tablist);

    wireDashSplitNavResize(catalog, {
        storageKey: DASH_CATALOG_SPLIT_NAV_WIDTH_KEY,
        ariaLabel: 'Breite der Katalog-Navigation anpassen'
    });

    relocateCatalogPersonalTools(catalog);

    const detail = catalog.querySelector('.dash-tasks-split__detail');
    if (detail) {
        mountToolsViewToggle(detail, { ariaLabel: 'Darstellung des Werkzeugkatalogs' });
    }
    catalog.querySelectorAll('.dashboard-tabpanel').forEach(function (panel) {
        ensureCatalogInsightTile(panel);
    });
    applyToolsViewMode(readToolsViewMode());

    document.body.classList.add('dash-catalog-split-active');
    syncCatalogSplitNavCurrent();

    catalog.addEventListener('click', function (ev) {
        const tab = ev.target && ev.target.closest ? ev.target.closest('[data-dashboard-tab]') : null;
        if (tab) {
            window.setTimeout(syncCatalogSplitNavCurrent, 0);
        }
        const open = ev.target && ev.target.closest ? ev.target.closest('a.btn[href], a.card-btn[href]') : null;
        if (open && open.closest('.choice[data-tool-id]') && isDashboardItShellActive()) {
            const panelEl = open.closest('.dashboard-tabpanel');
            const panelId = panelEl && panelEl.id ? panelEl.id.replace(/^dash-panel-/, '') : '';
            const tree = catalog.__ms365CatalogNavTree;
            if (Array.isArray(tree) && tree.length) {
                persistItWorkspaceNav(tree, panelId ? 'catalog-' + panelId : '');
            }
        }
    });
    return true;
}

if (typeof document !== 'undefined') {
    document.addEventListener('ms365-dash-catalog-ready', function () {
        mountDashboardCatalogSplit();
    });
}
