/**
 * Ansichtsmodus für IT-Tool-Flächen: Liste | Kacheln (persistiert).
 * Gilt für „Was möchten Sie tun?“ und „Alle Werkzeuge“.
 */

export const DASH_TOOLS_VIEW_KEY = 'ms365-dash-tools-view-v1';
export const DASH_TOOLS_VIEW_EVENT = 'ms365-dash-tools-view-changed';

/** @typedef {'list'|'tiles'} DashToolsViewMode */

/**
 * @returns {DashToolsViewMode}
 */
export function readToolsViewMode() {
    try {
        const v = String(localStorage.getItem(DASH_TOOLS_VIEW_KEY) || '')
            .trim()
            .toLowerCase();
        if (v === 'tiles' || v === 'kacheln' || v === 'grid') return 'tiles';
    } catch {
        /* ignore */
    }
    return 'list';
}

/**
 * @param {DashToolsViewMode} mode
 */
export function writeToolsViewMode(mode) {
    const next = mode === 'tiles' ? 'tiles' : 'list';
    try {
        localStorage.setItem(DASH_TOOLS_VIEW_KEY, next);
    } catch {
        /* ignore */
    }
    return next;
}

/**
 * @param {DashToolsViewMode} [mode]
 */
export function applyToolsViewMode(mode) {
    const next = mode === 'tiles' || mode === 'list' ? mode : readToolsViewMode();
    if (typeof document === 'undefined') return next;
    document.documentElement.setAttribute('data-dash-tools-view', next);
    document.querySelectorAll('.dash-tasks-split, #dash-catalog, #dashTasksSplit').forEach(function (el) {
        el.setAttribute('data-dash-tools-view', next);
    });
    document.querySelectorAll('.dash-view-seg[data-dash-tools-view-toggle]').forEach(function (seg) {
        syncToggleButtons(seg, next);
    });
    return next;
}

/**
 * @param {HTMLElement} seg
 * @param {DashToolsViewMode} mode
 */
function syncToggleButtons(seg, mode) {
    seg.querySelectorAll('[data-dash-tools-view]').forEach(function (btn) {
        const on = btn.getAttribute('data-dash-tools-view') === mode;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
}

/**
 * @param {DashToolsViewMode} mode
 */
export function setToolsViewMode(mode) {
    const next = writeToolsViewMode(mode);
    applyToolsViewMode(next);
    try {
        window.dispatchEvent(new CustomEvent(DASH_TOOLS_VIEW_EVENT, { detail: { mode: next } }));
    } catch {
        /* ignore */
    }
    return next;
}

/**
 * Segment-Umschalter Liste | Kacheln.
 * @param {HTMLElement} host
 * @param {{ ariaLabel?: string }} [opts]
 * @returns {HTMLElement|null}
 */
export function mountToolsViewToggle(host, opts) {
    if (!host) return null;
    let bar = host.querySelector('.dash-tools-view-bar');
    if (bar) {
        applyToolsViewMode(readToolsViewMode());
        return bar;
    }

    bar = document.createElement('div');
    bar.className = 'dash-tools-view-bar';
    bar.innerHTML =
        '<div class="dash-view-seg" data-dash-tools-view-toggle role="group" aria-label="' +
        (opts && opts.ariaLabel ? opts.ariaLabel : 'Darstellung der Werkzeuge') +
        '">' +
        '<button type="button" class="dash-view-seg__btn" data-dash-tools-view="list" aria-pressed="false">' +
        '<i class="bi bi-list-ul" aria-hidden="true"></i><span>Liste</span>' +
        '</button>' +
        '<button type="button" class="dash-view-seg__btn" data-dash-tools-view="tiles" aria-pressed="false">' +
        '<i class="bi bi-grid-3x3-gap" aria-hidden="true"></i><span>Kacheln</span>' +
        '</button>' +
        '</div>';

    const seg = bar.querySelector('.dash-view-seg');
    seg.querySelectorAll('[data-dash-tools-view]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            setToolsViewMode(btn.getAttribute('data-dash-tools-view') === 'tiles' ? 'tiles' : 'list');
        });
    });

    host.insertBefore(bar, host.firstChild);
    applyToolsViewMode(readToolsViewMode());
    return bar;
}

/**
 * @param {HTMLElement|null} el
 * @returns {string}
 */
function plainInsightText(el) {
    if (!el) return '';
    const clone = el.cloneNode(true);
    clone
        .querySelectorAll(
            '.context-tip, .featured-info, .bi, button, [hidden], .dash-task-index, script, style'
        )
        .forEach(function (n) {
            n.remove();
        });
    return String(clone.textContent || '')
        .replace(/\s+/g, ' ')
        .trim();
}

/**
 * @param {HTMLElement} task
 * @returns {{ title: string, lead: string, progText: string, tone: string, toolCount: number }}
 */
function readGroupInsightData(task) {
    const titleEl = task.querySelector('h3');
    const leadEl = task.querySelector(
        ':scope > p:not(.dash-task-progress):not(.dash-task-scan-meta), .dash-split-section--intro p'
    );
    const prog = task.querySelector('.dash-task-progress');
    return {
        title: plainInsightText(titleEl),
        lead: plainInsightText(leadEl),
        progText: plainInsightText(prog),
        tone: prog ? String(prog.getAttribute('data-tone') || '').trim() : '',
        toolCount: task.querySelectorAll('.dash-split-tool-card, .dash-task-row').length
    };
}

/**
 * @param {HTMLElement} aside
 * @param {{ title: string, lead: string, progText: string, tone: string, toolCount: number }} data
 */
function paintGroupInsightAside(aside, data) {
    if (data.tone) aside.setAttribute('data-tone', data.tone);
    else aside.removeAttribute('data-tone');

    aside.innerHTML = '';
    const kicker = document.createElement('div');
    kicker.className = 'dash-group-insight__kicker';
    kicker.textContent = 'Übersicht';
    aside.appendChild(kicker);

    if (data.title) {
        const h = document.createElement('div');
        h.className = 'dash-group-insight__title';
        h.textContent = data.title;
        aside.appendChild(h);
    }
    if (data.lead) {
        const p = document.createElement('p');
        p.className = 'dash-group-insight__lead';
        p.textContent = data.lead;
        aside.appendChild(p);
    }

    const stats = document.createElement('div');
    stats.className = 'dash-group-insight__stats';
    if (data.progText) {
        const s = document.createElement('div');
        s.className = 'dash-group-insight__stat';
        s.innerHTML = '<span class="dash-group-insight__stat-label">Stand</span>';
        const val = document.createElement('span');
        val.className = 'dash-group-insight__stat-value';
        val.textContent = data.progText;
        if (data.tone) val.setAttribute('data-tone', data.tone);
        s.appendChild(val);
        stats.appendChild(s);
    }
    if (data.toolCount > 0) {
        const s = document.createElement('div');
        s.className = 'dash-group-insight__stat';
        s.innerHTML =
            '<span class="dash-group-insight__stat-label">Werkzeuge</span>' +
            '<span class="dash-group-insight__stat-value">' +
            String(data.toolCount) +
            '</span>';
        stats.appendChild(s);
    }
    if (stats.childElementCount) aside.appendChild(stats);
}

/**
 * Info-Kachel zur aktuellen Gruppierung (Status, Kurztext, Kennzahl).
 * @param {HTMLElement} task
 */
export function ensureGroupInsightTile(task) {
    if (!task) return;
    const data = readGroupInsightData(task);
    if (!data.title && !data.lead && !data.progText) return;

    let aside = task.querySelector('.dash-group-insight');
    if (!aside) {
        aside = document.createElement('aside');
        aside.className = 'dash-group-insight';
        const toolsSection = task.querySelector('.dash-split-section--tools');
        const statusSection = task.querySelector('.dash-split-section--status');
        const anchor = toolsSection || statusSection || task.querySelector('.dash-task-links--stack');
        if (anchor && anchor.parentElement) anchor.parentElement.insertBefore(aside, anchor);
        else task.appendChild(aside);
    }
    paintGroupInsightAside(aside, data);
    const progRow = task.querySelector('.dash-task-progress-row');
    if (progRow && data.progText) progRow.setAttribute('data-insight-mirrored', '1');
}

/**
 * @param {HTMLElement} task
 */
export function refreshGroupInsightTile(task) {
    if (!task) return;
    const aside = task.querySelector('.dash-group-insight');
    if (!aside) {
        ensureGroupInsightTile(task);
        return;
    }
    const data = readGroupInsightData(task);
    paintGroupInsightAside(aside, data);
    const progRow = task.querySelector('.dash-task-progress-row');
    if (progRow && data.progText) progRow.setAttribute('data-insight-mirrored', '1');
}

/**
 * Katalog: kurze Info-Kachel über dem Cluster-Grid.
 * @param {HTMLElement} panel
 */
export function ensureCatalogInsightTile(panel) {
    if (!panel || panel.querySelector('.dash-group-insight')) return;
    const head = panel.querySelector('.dashboard-cluster-head');
    if (!head) return;
    const titleEl = head.querySelector('h3');
    const leadEl = head.querySelector('p');
    const title = titleEl
        ? String(titleEl.textContent || '')
              .replace(/\s+/g, ' ')
              .trim()
        : '';
    const lead = leadEl
        ? String(leadEl.textContent || '')
              .replace(/\s+/g, ' ')
              .trim()
        : '';
    const count = panel.querySelectorAll('.choice[data-tool-id]').length;
    if (!title && !lead) return;

    const aside = document.createElement('aside');
    aside.className = 'dash-group-insight dash-group-insight--catalog';
    aside.innerHTML =
        '<div class="dash-group-insight__kicker">Bereich</div>' +
        (title ? '<div class="dash-group-insight__title"></div>' : '') +
        (lead ? '<p class="dash-group-insight__lead"></p>' : '') +
        '<div class="dash-group-insight__stats">' +
        '<div class="dash-group-insight__stat">' +
        '<span class="dash-group-insight__stat-label">Einträge</span>' +
        '<span class="dash-group-insight__stat-value">' +
        String(count) +
        '</span></div></div>';
    const t = aside.querySelector('.dash-group-insight__title');
    if (t) t.textContent = title;
    const p = aside.querySelector('.dash-group-insight__lead');
    if (p) p.textContent = lead;

    const cluster = panel.querySelector('.dashboard-cluster');
    const grid = panel.querySelector('.grid');
    if (grid && grid.parentElement) grid.parentElement.insertBefore(aside, grid);
    else if (cluster) cluster.insertBefore(aside, cluster.firstChild);
    else panel.insertBefore(aside, panel.firstChild);
}

if (typeof document !== 'undefined') {
    applyToolsViewMode(readToolsViewMode());
    document.addEventListener('DOMContentLoaded', function () {
        applyToolsViewMode(readToolsViewMode());
    });
}
