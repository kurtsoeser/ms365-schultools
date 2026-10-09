/**
 * Dashboard: Playbooks – Kategorien-Split + Kacheln (wie Aufgaben/Katalog).
 */
import { loadPlaybookState } from './playbook-store.js';
import {
    DASHBOARD_PLAYBOOKS,
    PLAYBOOK_CATEGORIES,
    computePlaybookProgress,
    formatPlaybookMetaLine,
    formatPlaybookProgressLabel,
    playbookBiClass,
    playbooksInCategory
} from './dashboard-playbooks-catalog.js';
import { wireDashSplitNavResize } from './dashboard-tasks-split.js';

const CATEGORY_STORAGE_KEY = 'ms365-dash-playbooks-category-v1';
const NAV_WIDTH_KEY = 'ms365-dash-playbooks-split-nav-width-v1';

/**
 * @param {import('./dashboard-playbooks-catalog.js').DashboardPlaybookDef} def
 * @param {Storage} [storage]
 */
export function getDashboardPlaybookProgress(def, storage) {
    const st = loadPlaybookState(def.storageKey, storage);
    return computePlaybookProgress(st, def.stepIds);
}

/**
 * @param {string} s
 */
function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * @param {import('./dashboard-playbooks-catalog.js').DashboardPlaybookDef} def
 * @param {ReturnType<typeof computePlaybookProgress>} progress
 */
function renderPlaybookTile(def, progress) {
    const label = formatPlaybookProgressLabel(progress);
    const pct = progress.total ? Math.round(progress.ratio * 100) : 0;
    const complete = progress.status === 'complete' && progress.total > 0;
    const notStarted = progress.status === 'not-started' || progress.done === 0;
    const ctaLabel = complete ? 'Öffnen' : notStarted ? 'Starten' : 'Weiter';
    const statusTone =
        complete ? 'ok' : progress.status === 'in-progress' ? 'warn' : 'pending';
    const cardClass =
        'dash-split-tool-card dash-playbook-card dash-task-row' +
        (complete ? ' dash-playbook-card--complete' : '') +
        (progress.status === 'in-progress' ? ' dash-task-row--primary' : '');

    const fraction =
        progress.total > 0
            ? progress.done + ' / ' + progress.total
            : '';

    return (
        '<a class="' +
        cardClass +
        '" href="' +
        escapeHtml(def.href) +
        '" data-playbook-id="' +
        escapeHtml(def.id) +
        '">' +
        '<div class="dash-task-row__body">' +
        '<div class="dash-playbook-card__top">' +
        '<span class="dash-playbook-card__icon" aria-hidden="true"><i class="' +
        escapeHtml(playbookBiClass(def.icon)) +
        '"></i></span>' +
        '<div class="dash-playbook-card__intro">' +
        '<span class="dash-task-row__title">' +
        escapeHtml(def.title) +
        '</span>' +
        '<span class="dash-playbook-card__meta">' +
        escapeHtml(formatPlaybookMetaLine(def)) +
        '</span></div></div>' +
        (def.blurb ? '<p class="dash-task-row__desc">' + escapeHtml(def.blurb) + '</p>' : '') +
        '<div class="dash-playbook-card__progress" role="group" aria-label="Fortschritt">' +
        '<div class="dash-playbook-card__bar-row">' +
        '<div class="dash-playbook-card__bar" aria-hidden="true">' +
        '<span class="dash-playbook-card__bar-fill" style="width:' +
        pct +
        '%"></span></div>' +
        (fraction
            ? '<span class="dash-playbook-card__fraction" aria-hidden="true">' +
              escapeHtml(fraction) +
              '</span>'
            : '') +
        '</div></div>' +
        '<div class="dash-task-row__foot">' +
        '<span class="dash-task-row__status" data-tone="' +
        statusTone +
        '">' +
        escapeHtml(label) +
        '</span>' +
        '<span class="dash-playbook-card__cta" aria-hidden="true">' +
        escapeHtml(ctaLabel) +
        ' <i class="bi bi-arrow-right-short" aria-hidden="true"></i></span>' +
        '</div></div></a>'
    );
}

/**
 * @param {string} categoryId
 */
function renderCategoryPanel(categoryId) {
    const cat = PLAYBOOK_CATEGORIES.find(function (c) {
        return c.id === categoryId;
    });
    const items = playbooksInCategory(categoryId);
    const tiles = items
        .map(function (def) {
            return renderPlaybookTile(def, getDashboardPlaybookProgress(def));
        })
        .join('');

    return (
        '<div class="dash-tasks-split__panel dash-playbooks-split__panel" data-playbook-category="' +
        escapeHtml(categoryId) +
        '" role="region" aria-labelledby="dash-playbook-cat-' +
        escapeHtml(categoryId) +
        '">' +
        '<h3 id="dash-playbook-cat-' +
        escapeHtml(categoryId) +
        '">' +
        escapeHtml(cat ? cat.label : categoryId) +
        '</h3>' +
        '<p class="dash-playbooks-split__lead">Geführte Checklisten – Fortschritt aus den Haken auf der Playbook-Seite.</p>' +
        '<div class="dash-playbooks-split__grid" role="list">' +
        tiles +
        '</div></div>'
    );
}

function readStoredCategory() {
    try {
        const v = localStorage.getItem(CATEGORY_STORAGE_KEY);
        if (v && PLAYBOOK_CATEGORIES.some(function (c) { return c.id === v; })) return v;
    } catch {
        /* ignore */
    }
    return PLAYBOOK_CATEGORIES[0]?.id || 'grundlagen';
}

function storeCategory(id) {
    try {
        localStorage.setItem(CATEGORY_STORAGE_KEY, id);
    } catch {
        /* ignore */
    }
}

/**
 * @param {HTMLElement} split
 * @param {string} categoryId
 */
function activateCategory(split, categoryId) {
    split.querySelectorAll('[data-playbook-category]').forEach(function (panel) {
        const on = panel.getAttribute('data-playbook-category') === categoryId;
        panel.classList.toggle('is-active', on);
        panel.hidden = !on;
    });
    split.querySelectorAll('[data-playbook-nav-cat]').forEach(function (btn) {
        const on = btn.getAttribute('data-playbook-nav-cat') === categoryId;
        if (on) btn.setAttribute('aria-current', 'true');
        else btn.removeAttribute('aria-current');
    });
}

function buildSplitShell(mount) {
    mount.innerHTML =
        '<div class="dash-tasks-split dash-playbooks-split" id="dashPlaybooksSplit">' +
        '<nav class="dash-tasks-split__nav" id="dashPlaybooksNav" aria-label="Playbook-Kategorien"></nav>' +
        '<div class="dash-tasks-split__detail" id="dashPlaybooksDetail"></div>' +
        '</div>';

    const split = mount.querySelector('#dashPlaybooksSplit');
    const nav = mount.querySelector('#dashPlaybooksNav');
    const detail = mount.querySelector('#dashPlaybooksDetail');
    if (!split || !nav || !detail) return null;

    PLAYBOOK_CATEGORIES.forEach(function (cat, idx) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'dash-tasks-split__nav-item';
        btn.setAttribute('data-playbook-nav-cat', cat.id);
        const indexText = String(idx + 1).padStart(2, '0');
        btn.innerHTML =
            '<span class="dash-tasks-split__nav-row">' +
            '<span class="dash-tasks-split__nav-icon" aria-hidden="true"><i class="' +
            escapeHtml(playbookBiClass(cat.icon)) +
            '"></i></span>' +
            '<span class="dash-tasks-split__nav-body">' +
            '<span class="dash-tasks-split__nav-line">' +
            '<span class="dash-tasks-split__nav-index">' +
            indexText +
            '</span>' +
            '<span class="dash-tasks-split__nav-label">' +
            escapeHtml(cat.label) +
            '</span></span></span></span>';
        btn.addEventListener('click', function () {
            storeCategory(cat.id);
            activateCategory(split, cat.id);
        });
        nav.appendChild(btn);
    });

    detail.innerHTML = PLAYBOOK_CATEGORIES.map(function (c) {
        return renderCategoryPanel(c.id);
    }).join('');

    wireDashSplitNavResize(split, {
        storageKey: NAV_WIDTH_KEY,
        ariaLabel: 'Breite der Playbook-Navigation anpassen'
    });

    const initial = readStoredCategory();
    activateCategory(split, initial);
    document.body.classList.add('dash-playbooks-split-active');

    return split;
}

function refreshTiles(split) {
    PLAYBOOK_CATEGORIES.forEach(function (cat) {
        const panel = split.querySelector('[data-playbook-category="' + cat.id + '"]');
        if (!panel) return;
        const grid = panel.querySelector('.dash-playbooks-split__grid');
        if (!grid) return;
        grid.innerHTML = playbooksInCategory(cat.id)
            .map(function (def) {
                return renderPlaybookTile(def, getDashboardPlaybookProgress(def));
            })
            .join('');
    });
}

export function mountDashboardPlaybooksFlows() {
    const mount = document.getElementById('dashPlaybooksMount');
    if (!mount || mount.dataset.mounted === '1') return;
    mount.dataset.mounted = '1';

    const split = buildSplitShell(mount);
    if (!split) return;

    function refresh() {
        refreshTiles(split);
    }

    window.addEventListener('ms365-app-local-data-changed', function (e) {
        const d = e && e.detail;
        if (!d || d.source === 'playbook') refresh();
    });
    window.addEventListener('storage', refresh);

    return { refresh };
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', mountDashboardPlaybooksFlows);
    } else {
        mountDashboardPlaybooksFlows();
    }
}
