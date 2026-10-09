/**
 * Dashboard „Was möchten Sie tun?“ – Master-Detail (linke Navigation, Detail rechts).
 */

import {
    enrichTaskToolRows,
    mountTaskScenarioGuide
} from './dashboard-task-tool-copy.js';

const STORAGE_KEY = 'ms365-dash-tasks-split-v1';
export const DASH_TASKS_SPLIT_NAV_WIDTH_KEY = 'ms365-dash-tasks-split-nav-width-v1';
export const DASH_TASKS_SPLIT_NAV_WIDTH_MIN = 220;
export const DASH_TASKS_SPLIT_NAV_WIDTH_MAX = 520;
export const DASH_TASKS_SPLIT_NAV_WIDTH_DEFAULT = 300;

/** @param {number} px */
export function clampDashTasksSplitNavWidth(px) {
    const n = Number(px);
    if (!Number.isFinite(n)) return DASH_TASKS_SPLIT_NAV_WIDTH_DEFAULT;
    return Math.max(
        DASH_TASKS_SPLIT_NAV_WIDTH_MIN,
        Math.min(DASH_TASKS_SPLIT_NAV_WIDTH_MAX, Math.round(n))
    );
}

/**
 * @param {Pick<Storage, 'getItem'>} storage
 * @param {string} [storageKey]
 */
export function readStoredDashSplitNavWidth(storage, storageKey) {
    const key = storageKey || DASH_TASKS_SPLIT_NAV_WIDTH_KEY;
    try {
        const raw = storage.getItem(key);
        if (raw != null && raw !== '') return clampDashTasksSplitNavWidth(parseInt(raw, 10));
    } catch (_e) {
        /* ignore */
    }
    return DASH_TASKS_SPLIT_NAV_WIDTH_DEFAULT;
}

/**
 * @param {Pick<Storage, 'getItem'>} storage
 */
export function readStoredDashTasksSplitNavWidth(storage) {
    return readStoredDashSplitNavWidth(storage, DASH_TASKS_SPLIT_NAV_WIDTH_KEY);
}

/** @param {HTMLElement} split @param {number} px @param {string} [cssVar] */
function applyDashSplitNavWidth(split, px, cssVar) {
    const w = clampDashTasksSplitNavWidth(px);
    split.style.setProperty(cssVar || '--dash-tasks-split-nav-width', w + 'px');
    return w;
}

/**
 * @param {HTMLElement} split
 * @param {{
 *   storageKey?: string,
 *   cssVar?: string,
 *   ariaLabel?: string,
 *   resizingClass?: string
 * }} [options]
 */
export function wireDashSplitNavResize(split, options) {
    const opts = options && typeof options === 'object' ? options : {};
    const storageKey = opts.storageKey || DASH_TASKS_SPLIT_NAV_WIDTH_KEY;
    const cssVar = opts.cssVar || '--dash-tasks-split-nav-width';
    const ariaLabel = opts.ariaLabel || 'Breite der Navigation anpassen';
    const resizingClass = opts.resizingClass || 'is-resizing-nav';

    const detail = split.querySelector('.dash-tasks-split__detail');
    if (!detail || split.dataset.navResize === '1') return;
    split.dataset.navResize = '1';

    let lastWidth = applyDashSplitNavWidth(
        split,
        readStoredDashSplitNavWidth(localStorage, storageKey),
        cssVar
    );

    const handle = document.createElement('button');
    handle.type = 'button';
    handle.className = 'dash-tasks-split__resize';
    handle.setAttribute('role', 'separator');
    handle.setAttribute('aria-orientation', 'vertical');
    handle.setAttribute('aria-valuemin', String(DASH_TASKS_SPLIT_NAV_WIDTH_MIN));
    handle.setAttribute('aria-valuemax', String(DASH_TASKS_SPLIT_NAV_WIDTH_MAX));
    handle.setAttribute('aria-valuenow', String(lastWidth));
    handle.setAttribute('aria-label', ariaLabel);
    split.insertBefore(handle, detail);

    let dragging = false;

    function persistWidth() {
        try {
            localStorage.setItem(storageKey, String(lastWidth));
        } catch (_e) {
            /* ignore */
        }
    }

    function setFromPointer(clientX) {
        const rect = split.getBoundingClientRect();
        lastWidth = applyDashSplitNavWidth(split, clientX - rect.left, cssVar);
        handle.setAttribute('aria-valuenow', String(lastWidth));
    }

    handle.addEventListener('pointerdown', function (e) {
        if (e.button !== 0) return;
        dragging = true;
        split.classList.add(resizingClass);
        handle.classList.add('is-dragging');
        handle.setPointerCapture(e.pointerId);
        setFromPointer(e.clientX);
        e.preventDefault();
    });

    handle.addEventListener('pointermove', function (e) {
        if (!dragging) return;
        setFromPointer(e.clientX);
    });

    function endDrag(e) {
        if (!dragging) return;
        dragging = false;
        split.classList.remove(resizingClass);
        handle.classList.remove('is-dragging');
        try {
            handle.releasePointerCapture(e.pointerId);
        } catch (_e) {
            /* ignore */
        }
        persistWidth();
    }

    handle.addEventListener('pointerup', endDrag);
    handle.addEventListener('pointercancel', endDrag);

    handle.addEventListener('keydown', function (e) {
        const step = e.shiftKey ? 32 : 16;
        if (e.key === 'ArrowLeft') {
            lastWidth = applyDashSplitNavWidth(split, lastWidth - step, cssVar);
            handle.setAttribute('aria-valuenow', String(lastWidth));
            persistWidth();
            e.preventDefault();
        } else if (e.key === 'ArrowRight') {
            lastWidth = applyDashSplitNavWidth(split, lastWidth + step, cssVar);
            handle.setAttribute('aria-valuenow', String(lastWidth));
            persistWidth();
            e.preventDefault();
        }
    });
}

/** @param {HTMLElement} split */
function wireDashTasksSplitNavResize(split) {
    wireDashSplitNavResize(split, {
        storageKey: DASH_TASKS_SPLIT_NAV_WIDTH_KEY,
        ariaLabel: 'Breite der Aufgaben-Navigation anpassen'
    });
}

const NAV_LABELS = {
    dashTaskImportVerknuepfen: 'Daten importieren',
    dashTaskGruppen: 'Mitgliedschaften',
    dashTaskUnterricht: 'Klassen & Unterricht',
    dashTaskPersonen: 'Personen',
    dashTaskSchuljahr: 'Schuljahresstart',
    dashTaskIntranet: 'Intranet',
    dashTaskSchulApps: 'Schul-Apps',
    dashTaskOrdnung: 'Aufräumen'
};

const PROGRESS_BY_TASK = {
    dashTaskGruppen: 'dashProgGruppen',
    dashTaskUnterricht: 'dashProgUnterricht',
    dashTaskSchuljahr: 'dashProgSchuljahr'
};

/** @type {Record<string, string>} */
const TASK_HASH_BY_ID = {
    dashTaskImportVerknuepfen: 'import',
    dashTaskGruppen: 'mitgliedschaften',
    dashTaskUnterricht: 'klassen',
    dashTaskPersonen: 'personen',
    dashTaskSchuljahr: 'schuljahresstart',
    dashTaskIntranet: 'intranet',
    dashTaskSchulApps: 'schulapps',
    dashTaskOrdnung: 'aufraeumen'
};

const NAV_ICONS = {
    dashTaskImportVerknuepfen: 'bi bi-box-arrow-in-down',
    dashTaskGruppen: 'bi bi-people',
    dashTaskUnterricht: 'bi bi-mortarboard',
    dashTaskPersonen: 'bi bi-person-badge',
    dashTaskSchuljahr: 'bi bi-calendar2-range',
    dashTaskIntranet: 'bi bi-house-door',
    dashTaskSchulApps: 'bi bi-window-stack',
    dashTaskOrdnung: 'bi bi-ui-checks-grid'
};

/** @param {string} id */
function readStoredTaskId(id) {
    try {
        const v = localStorage.getItem(STORAGE_KEY);
        if (v && document.getElementById(v)) return v;
    } catch {
        /* ignore */
    }
    return id;
}

/** @param {string} taskId */
function writeStoredTaskId(taskId) {
    try {
        localStorage.setItem(STORAGE_KEY, taskId);
    } catch {
        /* ignore */
    }
}

/** @param {HTMLElement} task */
function navLabelForTask(task) {
    const id = task.id || '';
    if (NAV_LABELS[id]) return NAV_LABELS[id];
    const custom = task.getAttribute('data-dash-nav-label');
    if (custom) return custom.trim();
    const h3 = task.querySelector('h3');
    return h3 ? String(h3.textContent || '').trim() : 'Bereich';
}

/** @param {HTMLElement} task */
function flattenExpertTools(task) {
    const stack = task.querySelector('.dash-task-links--stack');
    if (!stack) return;
    stack.querySelectorAll('details.dash-task-expert').forEach(function (details) {
        const rows = Array.from(details.querySelectorAll('.dash-task-row'));
        rows.forEach(function (row) {
            row.classList.remove('dash-task-row--primary');
            stack.appendChild(row);
        });
        details.remove();
    });
}

/** @param {string} title */
function createSplitSectionHead(title) {
    const head = document.createElement('div');
    head.className = 'dash-split-section-head';
    const text = document.createElement('span');
    text.className = 'dash-split-section-head__text';
    text.textContent = title;
    head.appendChild(text);
    return head;
}

/**
 * @param {string} mod
 * @param {string} title
 * @param {HTMLElement} anchor
 */
function wrapSplitSection(mod, title, anchor, fillBody) {
    if (!anchor || !anchor.parentElement) return null;
    const parent = anchor.parentElement;
    const section = document.createElement('section');
    section.className = 'dash-split-section dash-split-section--' + mod;
    section.appendChild(createSplitSectionHead(title));
    const body = document.createElement('div');
    body.className = 'dash-split-section__body';
    section.appendChild(body);
    // Zuerst einhängen – fillBody verschiebt ggf. anchor aus parent (insertBefore braucht anchor noch im parent).
    parent.insertBefore(section, anchor);
    fillBody(body);
    return section;
}

/** @param {HTMLElement} task */
function collectTaskDeviationMessages(task) {
    const msgs = [];
    task.querySelectorAll('.dash-task-progress').forEach(function (el) {
        const tone = String(el.getAttribute('data-tone') || '').trim();
        const text = String(el.textContent || '').trim();
        if (!text) return;
        if (tone === 'warn' || tone === 'mismatch' || tone === 'unmatched') {
            msgs.push(text);
        }
    });
    return msgs;
}

/** @param {HTMLElement} task */
function ensureAlertSection(task) {
    const existing = task.querySelector('.alert-section');
    const msgs = collectTaskDeviationMessages(task);
    if (!msgs.length) {
        if (existing) existing.remove();
        return;
    }
    const summary = msgs.slice(0, 2).join(' · ');
    if (existing) {
        const content = existing.querySelector('.alert-content');
        if (content) {
            content.innerHTML =
                '<strong>Handlungsbedarf:</strong> ' +
                summary +
                '. Das empfohlene Werkzeug ist unten hervorgehoben.';
        }
        return;
    }
    const alert = document.createElement('div');
    alert.className = 'alert-section';
    alert.setAttribute('role', 'alert');
    alert.innerHTML =
        '<div class="alert-icon" aria-hidden="true">⚠️</div>' +
        '<div class="alert-content"><strong>Handlungsbedarf:</strong> ' +
        summary +
        '. Das empfohlene Werkzeug ist unten hervorgehoben.</div>';
    const top = task.querySelector('.dash-task-top');
    const anchor = top ? top.nextElementSibling : task.querySelector('h3');
    if (anchor && anchor.parentElement) {
        anchor.parentElement.insertBefore(alert, anchor.nextSibling);
    } else {
        task.insertBefore(alert, task.firstChild);
    }
}

/** @param {HTMLElement} task */
function structureDetailSections(task) {
    if (task.dataset.splitStructured === '1') return;

    ensureAlertSection(task);

    const progRow = task.querySelector('.dash-task-progress-row');
    const scanRow = task.querySelector('.dash-task-scan-row');
    const links = task.querySelector('.dash-task-links--stack');

    const statusAnchor = progRow || scanRow;
    if (statusAnchor) {
        wrapSplitSection('status', 'Aktueller Stand', statusAnchor, function (body) {
            if (progRow) body.appendChild(progRow);
            if (scanRow) {
                const actions = document.createElement('div');
                actions.className = 'dash-split-detail__actions';
                actions.appendChild(scanRow);
                body.appendChild(actions);
            }
        });
    }

    if (links) {
        layoutToolsGrid(task);
        wrapSplitSection('tools', '', links, function (body) {
            body.appendChild(links);
        });
    }

    enrichTaskToolRows(task);
    task.dataset.splitStructured = '1';
}

/** @param {string} title @param {HTMLElement[]} rows */
function appendToolsBlock(parent, title, rows) {
    if (!rows.length) return;
    const block = document.createElement('div');
    block.className = 'dash-split-tools-block';
    if (title) {
        const head = createSplitSectionHead(title);
        head.classList.add('dash-split-tools-block__head');
        block.appendChild(head);
    }
    const grid = document.createElement('div');
    grid.className = 'dash-split-tools-grid';
    rows.forEach(function (row) {
        row.classList.add('dash-split-tool-card');
        grid.appendChild(row);
    });
    block.appendChild(grid);
    parent.appendChild(block);
}

/** @param {HTMLElement} task */
function layoutToolsGrid(task) {
    const stack = task.querySelector('.dash-task-links--stack');
    if (!stack || stack.dataset.splitTools === '1') return;
    const rows = Array.from(stack.querySelectorAll(':scope > .dash-task-row'));
    if (!rows.length) return;

    const primary = [];
    const more = [];
    rows.forEach(function (row) {
        if (row.classList.contains('dash-task-row--primary')) primary.push(row);
        else more.push(row);
    });

    stack.textContent = '';
    stack.dataset.splitTools = '1';
    stack.classList.add('dash-task-links--split');

    appendToolsBlock(stack, 'Empfohlene Werkzeuge', primary);
    appendToolsBlock(stack, 'Weitere Werkzeuge', more);
}

/** @param {HTMLElement} task */
function navIconClassForTask(task) {
    const id = task.id || '';
    if (NAV_ICONS[id]) return NAV_ICONS[id];
    const icon = task.querySelector('.dash-task-icon i');
    if (icon && icon.className) return icon.className;
    return 'bi bi-circle';
}

/**
 * @param {string} tone
 * @param {string} text
 * @returns {'ok'|'warn'|'muted'|'none'}
 */
export function navStatusDotTone(tone, text) {
    const t = String(tone || '').trim();
    const tx = String(text || '').trim();
    if (t === 'ok') return 'ok';
    if (t === 'warn' || t === 'mismatch') return 'warn';
    if (t === 'unmatched') return 'muted';
    if (tx) {
        if (/schritt/i.test(tx)) return 'warn';
        return 'muted';
    }
    return 'none';
}

/** @param {string} tone @param {string} text */
export function navBadgeFromProgress(tone, text) {
    const t = String(tone || '').trim();
    const tx = String(text || '').trim();
    if (!tx) {
        if (t === 'ok') return { text: 'Konsistent', tone: 'ok' };
        return { text: '', tone: '' };
    }
    if (t === 'ok') return { text: 'Konsistent', tone: 'ok' };
    if (t === 'warn' || t === 'mismatch') return { text: 'Abweichung', tone: 'warn' };
    if (t === 'unmatched') return { text: 'Offen', tone: 'unmatched' };
    if (/schritt/i.test(tx)) {
        const m = tx.match(/(\d+)\s*\/\s*(\d+)/);
        if (m) return { text: m[1] + '/' + m[2] + ' Schritte', tone: t || 'warn' };
    }
    return { text: tx.length > 42 ? tx.slice(0, 40) + '…' : tx, tone: t || '' };
}

/** @param {string} taskId */
function badgeForTaskId(taskId) {
    const progId = PROGRESS_BY_TASK[taskId];
    if (progId) {
        const el = document.getElementById(progId);
        if (el && el.textContent && el.textContent.trim()) {
            return navBadgeFromProgress(el.getAttribute('data-tone'), el.textContent);
        }
        if (el && el.getAttribute('data-tone') === 'ok') {
            return navBadgeFromProgress('ok', '');
        }
    }
    if (taskId === 'dashTaskOrdnung') {
        const h = window.ms365MembershipHygiene;
        if (h && typeof h.loadHygieneScanCache === 'function') {
            const c = h.loadHygieneScanCache();
            if (c && c.counts) {
                if ((c.mismatch || 0) > 0 || (c.emptyList || 0) > 0) {
                    return { text: 'Datenhygiene', tone: 'warn' };
                }
                if ((c.ok || 0) > 0) return { text: 'Konsistent', tone: 'ok' };
            }
        }
    }
    return { text: '', tone: '' };
}

/**
 * @param {HTMLElement} btn
 * @param {string} taskId
 */
function paintNavBadge(btn, taskId) {
    const dot = btn.querySelector('.dash-tasks-split__nav-dot');
    if (!dot) return;
    const b = badgeForTaskId(taskId);
    const status = navStatusDotTone(b.tone, b.text);
    dot.setAttribute('data-status', status);
    if (b.text) dot.setAttribute('title', b.text);
    else dot.removeAttribute('title');
}

function rowMatchesSearch(row, needle) {
    if (!needle) return true;
    const extra = row.getAttribute('data-search') || '';
    const href = row.getAttribute('href') || '';
    const slug = href.replace(/^[./]*tools\//, '').replace(/\.html.*$/, '').replace(/[#?].*$/, '');
    const hay = ((row.textContent || '') + ' ' + extra + ' ' + slug.replace(/-/g, ' ')).toLowerCase();
    return hay.indexOf(needle) !== -1;
}

export function mountDashboardTasksSplit() {
    if (document.getElementById('dashTasksSplit')) return window.ms365DashTasksSplit || null;
    const grid = document.querySelector('.dash-task-grid.dash-main-tasks-only');
    const tasksRoot = document.getElementById('dashboard-tasks');
    if (!grid || !tasksRoot) return null;

    const allTasks = Array.from(grid.querySelectorAll('.dash-task'));
    const navTasks = allTasks.filter(function (t) {
        return t.id;
    });
    navTasks.sort(function (a, b) {
        if (a.id === 'dashTaskImportVerknuepfen') return -1;
        if (b.id === 'dashTaskImportVerknuepfen') return 1;
        return 0;
    });
    if (!navTasks.length) return null;

    const split = document.createElement('div');
    split.className = 'dash-tasks-split dash-main-tasks-only';
    split.id = 'dashTasksSplit';

    const nav = document.createElement('nav');
    nav.className = 'dash-tasks-split__nav';
    nav.setAttribute('aria-label', 'Aufgaben-Bereiche');

    const detail = document.createElement('div');
    detail.className = 'dash-tasks-split__detail';
    detail.id = 'dashTasksSplitDetail';

    const navButtons = [];

    navTasks.forEach(function (task) {
        flattenExpertTools(task);
        structureDetailSections(task);
        task.classList.add('dash-tasks-split__panel');
        task.hidden = true;

        const indexEl = task.querySelector('.dash-task-index');
        const indexText = indexEl ? String(indexEl.textContent || '').trim() : '';

        if (task.id === 'dashTaskGruppen') {
            const sep = document.createElement('div');
            sep.className = 'dash-tasks-split__nav-sep';
            sep.setAttribute('role', 'presentation');
            nav.appendChild(sep);
        }

        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'dash-tasks-split__nav-item';
        if (task.id === 'dashTaskImportVerknuepfen') btn.classList.add('dash-tasks-split__nav-item--start');
        btn.setAttribute('data-task-id', task.id);
        btn.setAttribute('aria-controls', task.id);
        btn.innerHTML =
            '<span class="dash-tasks-split__nav-row">' +
            '<span class="dash-tasks-split__nav-icon" aria-hidden="true"><i></i></span>' +
            '<span class="dash-tasks-split__nav-body">' +
            '<span class="dash-tasks-split__nav-line">' +
            '<span class="dash-tasks-split__nav-index"></span>' +
            '<span class="dash-tasks-split__nav-dot" data-status="none" aria-hidden="true"></span>' +
            '<span class="dash-tasks-split__nav-label"></span>' +
            '</span></span></span>';
        const iconEl = btn.querySelector('.dash-tasks-split__nav-icon i');
        if (iconEl) iconEl.className = navIconClassForTask(task);
        const indexSpan = btn.querySelector('.dash-tasks-split__nav-index');
        if (indexSpan) {
            if (indexText && indexText !== 'Start') indexSpan.textContent = indexText;
            else indexSpan.hidden = true;
        }
        btn.querySelector('.dash-tasks-split__nav-label').textContent = navLabelForTask(task);
        paintNavBadge(btn, task.id);

        btn.addEventListener('click', function () {
            selectTask(task.id);
        });

        nav.appendChild(btn);
        detail.appendChild(task);
        navButtons.push({ btn, task });
    });

    split.appendChild(nav);
    split.appendChild(detail);
    wireDashTasksSplitNavResize(split);
    grid.replaceWith(split);

    split.dataset.splitMounted = '1';
    document.body.classList.add('dash-tasks-split-active');

    function selectTask(taskId, opts) {
        const id = String(taskId || '').trim();
        const options = opts && typeof opts === 'object' ? opts : {};
        let found = false;
        navButtons.forEach(function (item) {
            const on = item.task.id === id;
            item.task.classList.toggle('is-active', on);
            item.task.hidden = !on;
            item.btn.setAttribute('aria-current', on ? 'true' : 'false');
            if (on) found = true;
        });
        if (!found && navButtons[0]) {
            selectTask(navButtons[0].task.id, options);
            return;
        }
        if (found) writeStoredTaskId(id);
        if (found && !options.skipHash) {
            const hash = TASK_HASH_BY_ID[id] || id;
            try {
                if (window.location.hash.replace(/^#/, '') !== hash) {
                    window.history.replaceState(null, '', '#' + hash);
                }
            } catch {
                /* ignore */
            }
        }
        if (options.pulse) {
            const el = document.getElementById(id);
            if (el) {
                el.classList.add('dash-task--pulse');
                window.setTimeout(function () {
                    el.classList.remove('dash-task--pulse');
                }, 1200);
            }
        }
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-dash-tasks-split-changed', { detail: { taskId: id } })
            );
        } catch {
            /* ignore */
        }
    }

    function refreshNavBadges() {
        navButtons.forEach(function (item) {
            paintNavBadge(item.btn, item.task.id);
            ensureAlertSection(item.task);
        });
    }

    /** @param {string} needle */
    function applySearch(needle) {
        const q = String(needle || '').trim().toLowerCase();
        const searching = q.length > 0;
        let hits = 0;
        let firstHit = '';

        navButtons.forEach(function (item) {
            const task = item.task;
            const sectionHay =
                navLabelForTask(task).toLowerCase() +
                ' ' +
                String(task.getAttribute('data-dash-task-search') || '').toLowerCase();
            const sectionMatch = !searching || sectionHay.indexOf(q) !== -1;
            let anyRow = false;
            task.querySelectorAll('.dash-task-row').forEach(function (row) {
                const rowMatch = sectionMatch || rowMatchesSearch(row, q);
                if (searching) row.hidden = !rowMatch;
                else row.hidden = false;
                if (rowMatch) anyRow = true;
            });
            const navVisible = !searching || sectionMatch || anyRow;
            item.btn.hidden = !navVisible;
            if (navVisible && searching) {
                hits += 1;
                if (!firstHit) firstHit = task.id;
            }
        });

        if (searching) {
            if (firstHit) selectTask(firstHit);
            if (hits > 0) tasksRoot.setAttribute('data-dash-task-search-visible', '1');
            else tasksRoot.removeAttribute('data-dash-task-search-visible');
        } else {
            tasksRoot.setAttribute('data-dash-task-search-visible', '1');
            navButtons.forEach(function (item) {
                item.btn.hidden = false;
            });
        }
        return hits;
    }

    const defaultId =
        (navTasks.find(function (t) {
            return t.id === 'dashTaskGruppen';
        }) || navTasks[0]).id;

    function taskIdFromLocationHash() {
        const h = String(window.location.hash || '').replace(/^#/, '').trim();
        if (!h) return '';
        for (const taskId of Object.keys(TASK_HASH_BY_ID)) {
            if (TASK_HASH_BY_ID[taskId] === h && document.getElementById(taskId)) return taskId;
        }
        return '';
    }

    const initialId = taskIdFromLocationHash() || readStoredTaskId(defaultId);
    selectTask(initialId, { skipHash: !taskIdFromLocationHash() });

    function refreshToolCards() {
        navButtons.forEach(function (item) {
            enrichTaskToolRows(item.task);
        });
    }

    mountTaskScenarioGuide(split, { select: selectTask });

    const api = {
        select: selectTask,
        refreshNavBadges,
        refreshToolCards,
        applySearch
    };
    window.ms365DashTasksSplit = api;
    window.__ms365DashTasksSplitRefreshBadges = refreshNavBadges;
    window.__ms365DashTasksSplitRefreshToolCards = refreshToolCards;

    return api;
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', mountDashboardTasksSplit);
    } else {
        mountDashboardTasksSplit();
    }
}
