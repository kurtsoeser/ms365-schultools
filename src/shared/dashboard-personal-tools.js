/**
 * Persönliche Leiste im Werkzeugkatalog: Favoriten (Meine Tools) und Zuletzt verwendet.
 */
import { toolLabel } from './dashboard-audience-catalog.js';
import { catalogToolCardActionLink } from './dashboard-tool-card.js';

export const DASHBOARD_FAVORITES_KEY = 'ms365-dashboard-favorites-v1';
export const DASHBOARD_RECENT_TOOLS_KEY = 'ms365-dashboard-recent-tools-v1';

export const DASHBOARD_RECENT_MAX = 8;

/**
 * @param {string} href
 */
export function normalizeDashboardToolHref(href) {
    const raw = String(href || '').trim();
    if (!raw) return '';
    try {
        const u = new URL(raw, 'https://local.invalid/');
        let path = u.pathname.replace(/^\//, '');
        path = path.replace(/^(\.\.\/)+/, '').replace(/^\.\//, '');
        return path.split('#')[0].split('?')[0].toLowerCase();
    } catch {
        return raw
            .replace(/^[./]+/, '')
            .split('#')[0]
            .split('?')[0]
            .toLowerCase();
    }
}

/**
 * @param {string[]} ids
 */
export function expandLegacyCatalogToolIds(ids) {
    const out = [];
    (ids || []).forEach(function (id) {
        if (id === 'slg') {
            out.push('slg-schueler', 'slg-lehrer');
        } else {
            out.push(id);
        }
    });
    return out;
}

/**
 * @param {Storage} storage
 */
export function loadDashboardFavorites(storage = localStorage) {
    try {
        const raw = storage.getItem(DASHBOARD_FAVORITES_KEY);
        if (!raw) return [];
        const o = JSON.parse(raw);
        return Array.isArray(o) ? expandLegacyCatalogToolIds(o.map(String)) : [];
    } catch {
        return [];
    }
}

/**
 * @param {Storage} storage
 */
export function loadRecentDashboardToolIds(storage = localStorage) {
    try {
        const raw = storage.getItem(DASHBOARD_RECENT_TOOLS_KEY);
        if (!raw) return [];
        const o = JSON.parse(raw);
        return Array.isArray(o) ? o.map(String) : [];
    } catch {
        return [];
    }
}

/**
 * @param {string} toolId
 * @param {{ storage?: Storage, max?: number }} [opts]
 */
export function recordDashboardToolVisit(toolId, opts) {
    const storage = (opts && opts.storage) || localStorage;
    const max = (opts && opts.max) || DASHBOARD_RECENT_MAX;
    const id = String(toolId || '').trim();
    if (!id) return;

    let list = loadRecentDashboardToolIds(storage).filter(function (x) {
        return x !== id;
    });
    list.unshift(id);
    if (list.length > max) list = list.slice(0, max);
    try {
        storage.setItem(DASHBOARD_RECENT_TOOLS_KEY, JSON.stringify(list));
    } catch {
        /* ignore */
    }
}

/**
 * @param {ParentNode} catalog
 */
export function buildDashboardToolHrefMap(catalog) {
    /** @type {Map<string, string>} */
    const map = new Map();
    if (!catalog) return map;
    catalog.querySelectorAll('.choice[data-tool-id]').forEach(function (card) {
        const toolId = card.getAttribute('data-tool-id');
        if (!toolId) return;
        const link = catalogToolCardActionLink(card);
        const href = link ? link.getAttribute('href') : '';
        const key = normalizeDashboardToolHref(href);
        if (key) map.set(key, toolId);
        map.set(toolId, toolId);
    });
    return map;
}

/**
 * @param {string} url
 * @param {Map<string, string>} hrefMap
 */
export function toolIdFromToolPageUrl(url, hrefMap) {
    const key = normalizeDashboardToolHref(url);
    if (!key) return '';
    if (hrefMap.has(key)) return hrefMap.get(key) || '';
    const file = key.replace(/^tools\//, '').replace(/\.html$/, '');
    if (hrefMap.has('tools/' + file + '.html')) {
        return hrefMap.get('tools/' + file + '.html') || '';
    }
    return file.replace(/\//g, '-');
}

/**
 * @param {Element} el
 * @param {ParentNode} catalog
 * @param {Map<string, string>} hrefMap
 */
export function resolveDashboardToolIdFromElement(el, catalog, hrefMap) {
    if (!el || !el.closest) return '';
    const card = el.closest('.choice[data-tool-id]');
    if (card) return String(card.getAttribute('data-tool-id') || '').trim();

    const hygiene = el.closest('[data-dash-hygiene-id]');
    if (hygiene) return String(hygiene.getAttribute('data-dash-hygiene-id') || '').trim();

    const link = el.closest('a[href]');
    if (link) {
        const href = link.getAttribute('href') || '';
        if (href.indexOf('tools/') !== -1 || href.indexOf('/tools/') !== -1) {
            const key = normalizeDashboardToolHref(href);
            if (hrefMap.has(key)) return hrefMap.get(key) || '';
        }
    }
    return '';
}

/**
 * @param {ParentNode} catalog
 * @param {string} toolId
 */
export function resolveDashboardToolHref(catalog, toolId) {
    if (!catalog || !toolId) return '';
    const card = catalog.querySelector('.choice[data-tool-id="' + toolId + '"]');
    if (!card) return '';
    const link = catalogToolCardActionLink(card);
    return link ? String(link.getAttribute('href') || '') : '';
}

/**
 * @param {HTMLElement} mountBefore
 * @param {{
 *   catalog: ParentNode,
 *   isToolAllowed?: (toolId: string) => boolean,
 *   onChange?: () => void
 * }} opts
 */
export function mountDashboardPersonalToolsBar(mountBefore, opts) {
    const catalog = opts.catalog;
    const isToolAllowed =
        opts.isToolAllowed ||
        function () {
            return true;
        };

    let hrefMap = buildDashboardToolHrefMap(catalog);

    const bar = document.createElement('div');
    bar.className = 'dash-personal-tools quick-access-bar';
    bar.id = 'dashPersonalTools';
    bar.setAttribute('aria-label', 'Meine Werkzeuge');
    bar.innerHTML =
        '<div class="dash-personal-tools__row quick-access-bar__inner">' +
        '<div class="dash-personal-tools__group dash-personal-tools__group--pinned quick-zone pinned">' +
        '<span class="dash-personal-tools__label quick-zone-label"><i class="bi bi-star-fill" aria-hidden="true"></i> Meine Tools</span>' +
        '<div class="dash-personal-tools__chips quick-items" id="dashPersonalPinned"></div>' +
        '</div>' +
        '<div class="quick-divider" role="presentation" aria-hidden="true"></div>' +
        '<div class="dash-personal-tools__group dash-personal-tools__group--recent quick-zone recent">' +
        '<span class="dash-personal-tools__label quick-zone-label"><i class="bi bi-clock-history" aria-hidden="true"></i> Zuletzt</span>' +
        '<div class="dash-personal-tools__chips quick-items" id="dashPersonalRecent"></div>' +
        '</div>' +
        '</div>';

    if (mountBefore && mountBefore.parentElement) {
        mountBefore.parentElement.insertBefore(bar, mountBefore);
    }

    const pinnedEl = bar.querySelector('#dashPersonalPinned');
    const recentEl = bar.querySelector('#dashPersonalRecent');

    function renderChip(toolId) {
        const href = resolveDashboardToolHref(catalog, toolId);
        if (!href) return null;
        const a = document.createElement('a');
        a.className = 'dash-personal-tools__chip quick-item';
        a.href = href;
        a.setAttribute('data-tool-id', toolId);
        a.textContent = toolLabel(toolId) || toolId;
        return a;
    }

    function renderGroup(container, ids, emptyText) {
        if (!container) return;
        container.textContent = '';
        const visible = ids.filter(function (id) {
            return isToolAllowed(id) && resolveDashboardToolHref(catalog, id);
        });
        if (!visible.length) {
            const span = document.createElement('span');
            span.className = 'dash-personal-tools__empty';
            span.textContent = emptyText;
            container.appendChild(span);
            return;
        }
        visible.forEach(function (id) {
            const chip = renderChip(id);
            if (chip) container.appendChild(chip);
        });
    }

    function refresh() {
        hrefMap = buildDashboardToolHrefMap(catalog);
        const favs = loadDashboardFavorites().filter(function (id, i, arr) {
            return arr.indexOf(id) === i;
        });
        const recent = loadRecentDashboardToolIds().filter(function (id) {
            return favs.indexOf(id) === -1;
        });

        renderGroup(
            pinnedEl,
            favs,
            'Noch keine Favoriten – Stern im Katalog setzen'
        );
        renderGroup(recentEl, recent, 'Noch keine zuletzt geöffneten Werkzeuge');

        const hasPinned = pinnedEl && pinnedEl.querySelector('.dash-personal-tools__chip');
        const hasRecent = recentEl && recentEl.querySelector('.dash-personal-tools__chip');
        bar.classList.toggle('dash-personal-tools--has-pinned', !!hasPinned);
        bar.classList.toggle('dash-personal-tools--has-recent', !!hasRecent);
    }

    function noteVisit(toolId) {
        if (!toolId || !isToolAllowed(toolId)) return;
        recordDashboardToolVisit(toolId);
        refresh();
        if (opts.onChange) opts.onChange();
    }

    function recordFromReferrer() {
        try {
            const ref = document.referrer;
            if (!ref) return;
            const id = toolIdFromToolPageUrl(ref, hrefMap);
            if (id && resolveDashboardToolHref(catalog, id)) noteVisit(id);
        } catch {
            /* ignore */
        }
    }

    bar.addEventListener('click', function (e) {
        const chip = e.target.closest('.dash-personal-tools__chip[data-tool-id]');
        if (!chip) return;
        noteVisit(chip.getAttribute('data-tool-id') || '');
    });

    function attachDashboardClickTracking(root) {
        if (!root) return;
        root.addEventListener(
            'click',
            function (e) {
                const id = resolveDashboardToolIdFromElement(e.target, catalog, hrefMap);
                if (id) noteVisit(id);
            },
            true
        );
    }

    attachDashboardClickTracking(document.getElementById('dashboard-tools'));
    attachDashboardClickTracking(document.getElementById('dashboard-tasks'));

    recordFromReferrer();
    refresh();

    return {
        refresh,
        noteVisit,
        rebuildHrefMap: function () {
            hrefMap = buildDashboardToolHrefMap(catalog);
            refresh();
        }
    };
}
