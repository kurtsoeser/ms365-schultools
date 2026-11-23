/**
 * Dashboard IT-Bereiche – Master-Detail (linke Gruppen-Navigation, Detail rechts).
 */

import {
    enrichTaskToolRows,
    mountTaskScenarioGuide
} from './dashboard-task-tool-copy.js';
import {
    applyToolsViewMode,
    ensureGroupInsightTile,
    mountToolsViewToggle,
    readToolsViewMode,
    refreshGroupInsightTile
} from './dashboard-tools-view.js';
import {
    buildNavTreeFromDashboard,
    clearItWorkspaceNav,
    deactivateItWorkspace,
    extractToolsFromTask,
    IT_WORKSPACE_STAMMDATEN_TASK_ID,
    isDashboardItShellActive,
    persistItWorkspaceNav
} from './it-workspace-nav.js';
import { mergeStammdatenWorkspaceTools } from './stammdaten-nav-catalog.js';

const STORAGE_KEY = 'ms365-dash-tasks-split-v1';
/** Offen/zu je Bereich (Mitgliedschaften, Aufräumen, …) – nicht mehr Abgleich/Betrieb. */
const NAV_GROUPS_STORAGE_KEY = 'ms365-dash-tasks-nav-groups-v2';
/** „Mehr“-Overflow je Bereich in der Sidebar. */
const NAV_MORE_STORAGE_KEY = 'ms365-dash-tasks-nav-more-v1';
export const DASH_TASKS_SPLIT_NAV_WIDTH_KEY = 'ms365-dash-tasks-split-nav-width-v1';
export const DASH_TASKS_SPLIT_NAV_WIDTH_MIN = 220;
export const DASH_TASKS_SPLIT_NAV_WIDTH_MAX = 520;
export const DASH_TASKS_SPLIT_NAV_WIDTH_DEFAULT = 256;

/** Reihenfolge der Bereichs-Gruppen in der IT-Leiste (Stammdaten zuerst, wie die anderen aufklappbar). */
const NAV_AREA_TASK_IDS = [
    'dashTaskImportVerknuepfen',
    'dashTaskGruppen',
    'dashTaskUnterricht',
    'dashTaskPersonen',
    'dashTaskSchuljahr',
    'dashTaskIntranet',
    'dashTaskSchulApps',
    'dashTaskOrdnung'
];

/** @returns {Record<string, boolean>} */
function readNavGroupOpenState() {
    try {
        const raw = JSON.parse(localStorage.getItem(NAV_GROUPS_STORAGE_KEY) || '{}');
        return raw && typeof raw === 'object' ? raw : {};
    } catch {
        return {};
    }
}

/** @param {Record<string, boolean>} state */
function writeNavGroupOpenState(state) {
    try {
        localStorage.setItem(NAV_GROUPS_STORAGE_KEY, JSON.stringify(state || {}));
    } catch {
        /* ignore */
    }
}

/** @returns {Record<string, boolean>} */
function readNavMoreOpenState() {
    try {
        const raw = JSON.parse(localStorage.getItem(NAV_MORE_STORAGE_KEY) || '{}');
        return raw && typeof raw === 'object' ? raw : {};
    } catch {
        return {};
    }
}

/** @param {Record<string, boolean>} state */
function writeNavMoreOpenState(state) {
    try {
        localStorage.setItem(NAV_MORE_STORAGE_KEY, JSON.stringify(state || {}));
    } catch {
        /* ignore */
    }
}

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

/** Wenn schon Schuldaten im Browser liegen: Einstieg als „Stammdaten“ statt nur Import. */
function hasMeaningfulStammdaten() {
    try {
        const settings =
            typeof window.ms365TenantSettingsLoad === 'function' ? window.ms365TenantSettingsLoad() : null;
        if (!settings || typeof settings !== 'object') return false;
        const domain = String(settings.domain || '').trim();
        const schoolName = String(settings.schoolName || '').trim();
        const subjects = Array.isArray(settings.subjects) ? settings.subjects.length : 0;
        const teachers = Array.isArray(settings.teachers) ? settings.teachers.length : 0;
        const students = Array.isArray(settings.students) ? settings.students.length : 0;
        const classes = Array.isArray(settings.classes) ? settings.classes.length : 0;
        return !!(domain || schoolName || subjects || teachers || students || classes);
    } catch {
        return false;
    }
}

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
    if (id === 'dashTaskImportVerknuepfen' && hasMeaningfulStammdaten()) return 'Stammdaten';
    if (NAV_LABELS[id]) return NAV_LABELS[id];
    const custom = task.getAttribute('data-dash-nav-label');
    if (custom) return custom.trim();
    const h3 = task.querySelector('h3');
    return h3 ? String(h3.textContent || '').trim() : 'Bereich';
}

/**
 * Import-Bereich: ohne Daten = Einstieg „importieren“, mit Daten = „Stammdaten“.
 * @param {HTMLElement} task
 * @param {HTMLButtonElement} [btn]
 */
function applyImportTaskPersona(task, btn) {
    if (!task || task.id !== 'dashTaskImportVerknuepfen') return;
    const hasData = hasMeaningfulStammdaten();
    task.classList.toggle('dash-task--has-stammdaten', hasData);
    const h3 = task.querySelector('h3');
    const lead = task.querySelector(':scope > p, .dash-split-section--intro p');
    const iconWrap = task.querySelector('.dash-task-icon i');
    const indexEl = task.querySelector('.dash-task-index');
    if (h3) {
        h3.textContent = hasData ? 'Stammdaten pflegen & importieren' : 'Erste Einrichtung (Blaupause)';
    }
    if (lead) {
        lead.textContent = hasData
            ? 'Listen prüfen, bei Bedarf neu importieren und Microsoft 365-Gruppen zuordnen.'
            : 'Export einlesen, in den Stammdaten prüfen und Microsoft 365-Gruppen Schritt für Schritt zuordnen.';
    }
    if (iconWrap) {
        iconWrap.className = hasData ? 'bi bi-database' : 'bi bi-box-arrow-in-down';
    }
    if (indexEl) {
        if (hasData) {
            indexEl.textContent = '';
            indexEl.hidden = true;
        } else {
            indexEl.textContent = 'Start';
            indexEl.hidden = false;
        }
    }
    const stack = task.querySelector('.dash-task-links--stack');
    if (stack) {
        let stammdatenRow = null;
        stack.querySelectorAll('a.dash-task-row[href*="tenant.html"]').forEach(function (a) {
            const href = String(a.getAttribute('href') || '');
            if (/organigramm/i.test(href)) return;
            if (!stammdatenRow) stammdatenRow = a;
        });
        if (stammdatenRow) {
            const label = stammdatenRow.querySelector('.dash-task-row__label');
            if (label) {
                label.innerHTML = hasData
                    ? '<i class="bi bi-database" aria-hidden="true"></i>Stammdaten öffnen'
                    : '<i class="bi bi-database" aria-hidden="true"></i>Stammdaten';
            }
            if (hasData) {
                stammdatenRow.classList.add('dash-task-row--primary');
                stack.insertBefore(stammdatenRow, stack.firstChild);
            }
        }
    }
    if (btn) {
        const navIcon = btn.querySelector('.dash-tasks-split__nav-icon i');
        if (navIcon) navIcon.className = hasData ? 'bi bi-database' : 'bi bi-box-arrow-in-down';
        const navLabel = btn.querySelector('.dash-tasks-split__nav-label');
        if (navLabel) navLabel.textContent = navLabelForTask(task);
        const navIndex = btn.querySelector('.dash-tasks-split__nav-index');
        if (navIndex) {
            if (hasData) {
                navIndex.textContent = '';
                navIndex.hidden = true;
            } else {
                navIndex.textContent = 'Start';
                navIndex.hidden = false;
            }
        }
        btn.classList.toggle('dash-tasks-split__nav-item--start', !hasData);
        btn.classList.toggle('dash-tasks-split__nav-item--stammdaten', hasData);
    }
}

/** Expert-Block behalten; Summary auf „Mehr …“ vereinheitlichen. */
function prepareExpertTools(task) {
    const stack = task.querySelector('.dash-task-links--stack');
    if (!stack) return;
    stack.querySelectorAll('details.dash-task-expert').forEach(function (details) {
        const summary = details.querySelector('summary');
        if (!summary) return;
        const icon = summary.querySelector('i.bi');
        const iconHtml = icon
            ? '<i class="' + icon.className + '" aria-hidden="true"></i>'
            : '<i class="bi bi-three-dots" aria-hidden="true"></i>';
        summary.innerHTML = iconHtml + 'Mehr …';
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
    ensureGroupInsightTile(task);
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
    const expertBlocks = Array.from(stack.querySelectorAll(':scope > details.dash-task-expert'));
    if (!rows.length && !expertBlocks.length) return;

    const primary = [];
    const looseMore = [];
    rows.forEach(function (row) {
        if (row.classList.contains('dash-task-row--primary')) primary.push(row);
        else looseMore.push(row);
    });

    stack.textContent = '';
    stack.dataset.splitTools = '1';
    stack.classList.add('dash-task-links--split');

    appendToolsBlock(stack, '', primary.concat(looseMore));
    expertBlocks.forEach(function (details) {
        stack.appendChild(details);
    });
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

/** @param {string} taskId */
function countForTaskId(taskId) {
    const progId = PROGRESS_BY_TASK[taskId];
    if (progId) {
        const el = document.getElementById(progId);
        const tone = el ? String(el.getAttribute('data-tone') || '').trim() : '';
        const tx = el ? String(el.textContent || '').trim() : '';
        if (tone === 'warn' || tone === 'mismatch') {
            const dev = tx.match(/(\d+)\s+Abweichung/i);
            if (dev) return Number(dev[1]) || 1;
            const frac = tx.match(/(\d+)\s*\/\s*(\d+)/);
            if (frac) {
                const diff = Number(frac[2]) - Number(frac[1]);
                if (diff > 0) return diff;
            }
            return 1;
        }
    }
    if (taskId === 'dashTaskOrdnung') {
        const h = window.ms365MembershipHygiene;
        if (h && typeof h.loadHygieneScanCache === 'function') {
            const c = h.loadHygieneScanCache();
            if (c && c.counts) {
                const n = Number(c.counts.mismatch || 0) + Number(c.counts.emptyList || 0);
                if (n > 0) return n;
            }
        }
    }
    return 0;
}

/**
 * @param {HTMLElement} btn
 * @param {string} taskId
 */
function paintNavBadge(btn, taskId) {
    const dot = btn.querySelector('.dash-tasks-split__nav-dot');
    const b = badgeForTaskId(taskId);
    const status = navStatusDotTone(b.tone, b.text);
    if (dot) {
        dot.setAttribute('data-status', status);
        if (b.text) dot.setAttribute('title', b.text);
        else dot.removeAttribute('title');
    }
    const countEl = btn.querySelector('.dash-tasks-split__nav-count');
    if (countEl) {
        const count = countForTaskId(taskId);
        if (count > 0) {
            countEl.textContent = String(count);
            countEl.hidden = false;
        } else {
            countEl.textContent = '';
            countEl.hidden = true;
        }
    }
}

function rowMatchesSearch(row, needle) {
    if (!needle) return true;
    const extra = row.getAttribute('data-search') || '';
    const href = row.getAttribute('href') || '';
    const slug = href.replace(/^[./]*tools\//, '').replace(/\.html.*$/, '').replace(/[#?].*$/, '');
    const labelEl = row.querySelector('.dash-task-row__label, .dash-task-row__title');
    let label = '';
    if (labelEl) {
        const clone = labelEl.cloneNode(true);
        clone.querySelectorAll('i, .bi').forEach(function (n) {
            n.remove();
        });
        label = String(clone.textContent || '').replace(/\s+/g, ' ').trim();
    }
    const hay = (label + ' ' + extra + ' ' + slug.replace(/-/g, ' ')).toLowerCase();
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
    /** @type {{ id: string, el: HTMLElement, toggle: HTMLButtonElement, body: HTMLElement, taskIds: string[] }[]} */
    const navGroups = [];
    const groupOpenState = readNavGroupOpenState();
    const taskById = {};
    navTasks.forEach(function (task) {
        taskById[task.id] = task;
    });

    function currentSelectedTaskId() {
        const current = navButtons.find(function (item) {
            return item.btn.getAttribute('aria-current') === 'true';
        });
        return current && current.task ? current.task.id : '';
    }

    function persistWorkspaceTree(activeTaskId) {
        if (!isDashboardItShellActive()) return;
        const tree = buildNavTreeFromDashboard(navButtons, navLabelForTask, navIconClassForTask);
        // Overview-Hashes ergänzen
        tree.forEach(function (node) {
            node.overviewHash = TASK_HASH_BY_ID[node.id] || '';
        });
        persistItWorkspaceNav(tree, activeTaskId || '');
    }

    const moreOpenState = readNavMoreOpenState();

    /**
     * @param {{ href: string, label: string, icon: string }} tool
     * @param {string} taskId
     */
    function createNavSubLink(tool, taskId) {
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
            persistWorkspaceTree(taskId);
        });
        return a;
    }

    /**
     * @param {HTMLElement} task
     * @returns {HTMLElement | null}
     */
    function buildNavSubForTask(task) {
        const extracted = extractToolsFromTask(task);
        const tools =
            task && task.id === IT_WORKSPACE_STAMMDATEN_TASK_ID
                ? mergeStammdatenWorkspaceTools(extracted)
                : extracted;
        if (!tools.length) return null;
        const primary = tools.filter(function (t) {
            return !t.expert;
        });
        const expert = tools.filter(function (t) {
            return !!t.expert;
        });
        const sub = document.createElement('div');
        sub.className = 'dash-tasks-split__nav-sub';
        // Flache Unterliste wie bei Mitgliedschaften – ein IT-Menü, gleiche Optik
        primary.forEach(function (tool) {
            sub.appendChild(createNavSubLink(tool, task.id));
        });
        if (expert.length) {
            const moreWrap = document.createElement('div');
            moreWrap.className = 'dash-tasks-split__nav-more';
            const moreBtn = document.createElement('button');
            moreBtn.type = 'button';
            moreBtn.className = 'dash-tasks-split__nav-more-btn';
            moreBtn.setAttribute('aria-expanded', 'false');
            moreBtn.innerHTML =
                '<span class="dash-tasks-split__nav-more-line" aria-hidden="true"></span>' +
                '<span class="dash-tasks-split__nav-more-label">Mehr …</span>' +
                '<span class="dash-tasks-split__nav-more-count">' +
                expert.length +
                '</span>';
            const moreBody = document.createElement('div');
            moreBody.className = 'dash-tasks-split__nav-more-body';
            moreBody.hidden = true;
            expert.forEach(function (tool) {
                moreBody.appendChild(createNavSubLink(tool, task.id));
            });
            function setMoreOpen(open, persist) {
                const next = !!open;
                moreBtn.setAttribute('aria-expanded', next ? 'true' : 'false');
                moreBody.hidden = !next;
                moreWrap.classList.toggle('is-open', next);
                moreBtn.querySelector('.dash-tasks-split__nav-more-label').textContent = next
                    ? 'Weniger'
                    : 'Mehr …';
                if (persist !== false) {
                    moreOpenState[task.id] = next;
                    writeNavMoreOpenState(moreOpenState);
                }
            }
            moreBtn.addEventListener('click', function (e) {
                e.preventDefault();
                e.stopPropagation();
                setMoreOpen(moreBtn.getAttribute('aria-expanded') !== 'true', true);
            });
            moreWrap.appendChild(moreBtn);
            moreWrap.appendChild(moreBody);
            sub.appendChild(moreWrap);
            setMoreOpen(moreOpenState[task.id] === true, false);
        }
        return sub;
    }

    function createNavItem(task, opts) {
        const options = opts && typeof opts === 'object' ? opts : {};
        const accordion = !!options.accordion;

        prepareExpertTools(task);
        structureDetailSections(task);
        task.classList.add('dash-tasks-split__panel');
        task.hidden = true;

        const indexEl = task.querySelector('.dash-task-index');
        const indexText = indexEl ? String(indexEl.textContent || '').trim() : '';

        const block = document.createElement('div');
        block.className = 'dash-tasks-split__nav-block';
        block.dataset.taskId = task.id;

        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'dash-tasks-split__nav-item';
        if (accordion) btn.classList.add('dash-tasks-split__nav-group-toggle');
        if (task.id === 'dashTaskImportVerknuepfen') btn.classList.add('dash-tasks-split__nav-item--start');
        btn.setAttribute('data-task-id', task.id);
        btn.setAttribute('aria-controls', task.id);
        btn.innerHTML =
            '<span class="dash-tasks-split__nav-row">' +
            '<span class="dash-tasks-split__nav-icon" aria-hidden="true"><i></i></span>' +
            '<span class="dash-tasks-split__nav-body">' +
            '<span class="dash-tasks-split__nav-line">' +
            '<span class="dash-tasks-split__nav-index" hidden></span>' +
            '<span class="dash-tasks-split__nav-dot" data-status="none" aria-hidden="true" hidden></span>' +
            '<span class="dash-tasks-split__nav-label"></span>' +
            '<span class="dash-tasks-split__nav-count" hidden></span>' +
            (accordion
                ? '<i class="bi bi-chevron-down dash-tasks-split__nav-group-chevron" aria-hidden="true"></i>'
                : '') +
            '</span></span></span>';
        const iconEl = btn.querySelector('.dash-tasks-split__nav-icon i');
        if (iconEl) iconEl.className = navIconClassForTask(task);
        const indexSpan = btn.querySelector('.dash-tasks-split__nav-index');
        if (indexSpan) {
            if (indexText && indexText !== 'Start') indexSpan.textContent = indexText;
            indexSpan.hidden = true;
        }
        btn.querySelector('.dash-tasks-split__nav-label').textContent = navLabelForTask(task);
        paintNavBadge(btn, task.id);
        if (task.id === 'dashTaskImportVerknuepfen') applyImportTaskPersona(task, btn);

        const sub = buildNavSubForTask(task);

        detail.appendChild(task);

        if (accordion) {
            const groupEl = document.createElement('div');
            groupEl.className = 'dash-tasks-split__nav-group';
            groupEl.dataset.navGroup = task.id;
            groupEl.dataset.taskId = task.id;

            btn.id = 'dashNavGroupToggle-' + task.id;
            btn.setAttribute('aria-controls', 'dashNavGroupBody-' + task.id);

            const body = document.createElement('div');
            body.className = 'dash-tasks-split__nav-group-body';
            body.id = 'dashNavGroupBody-' + task.id;
            body.setAttribute('role', 'group');
            body.setAttribute('aria-labelledby', btn.id);
            if (sub) body.appendChild(sub);
            else body.hidden = true;

            groupEl.appendChild(btn);
            groupEl.appendChild(body);
            block.appendChild(groupEl);

            const group = {
                id: task.id,
                el: groupEl,
                toggle: btn,
                body: body,
                taskIds: [task.id]
            };
            navGroups.push(group);

            const defaultOpen = groupOpenState[task.id] === true;
            setGroupOpen(group, defaultOpen, false);

            btn.addEventListener('click', function () {
                const isCurrent = btn.getAttribute('aria-current') === 'true';
                const isOpen = btn.getAttribute('aria-expanded') === 'true';
                if (isCurrent && isOpen) {
                    setGroupOpen(group, false, true);
                    return;
                }
                selectTask(task.id);
            });

            navButtons.push({ btn: btn, task: task, block: block, sub: sub, group: group });
            return block;
        }

        btn.addEventListener('click', function () {
            selectTask(task.id);
        });
        block.appendChild(btn);
        if (sub) {
            sub.hidden = true;
            block.appendChild(sub);
        }
        navButtons.push({ btn: btn, task: task, block: block, sub: sub, group: null });
        return block;
    }

    function setGroupOpen(group, open, persist) {
        const next = !!open;
        group.el.classList.toggle('is-collapsed', !next);
        group.toggle.setAttribute('aria-expanded', next ? 'true' : 'false');
        group.body.hidden = !next;
        if (persist !== false) {
            groupOpenState[group.id] = next;
            writeNavGroupOpenState(groupOpenState);
        }
    }

    function expandGroupForTask(taskId) {
        navGroups.forEach(function (group) {
            if (group.taskIds.indexOf(taskId) !== -1) setGroupOpen(group, true, true);
        });
    }

    // Bereiche = aufklappbare Gruppen (Stammdaten, Mitgliedschaften, …)
    NAV_AREA_TASK_IDS.forEach(function (id) {
        const task = taskById[id];
        if (!task) return;
        nav.appendChild(createNavItem(task, { accordion: true }));
    });

    // Unbekannte Tasks (falls später ergänzt) ans Ende
    navTasks.forEach(function (task) {
        if (NAV_AREA_TASK_IDS.indexOf(task.id) !== -1) return;
        nav.appendChild(createNavItem(task, { accordion: true }));
    });

    split.appendChild(nav);
    split.appendChild(detail);
    mountToolsViewToggle(detail, { ariaLabel: 'Darstellung der Aufgaben-Werkzeuge' });
    applyToolsViewMode(readToolsViewMode());
    wireDashTasksSplitNavResize(split);
    grid.replaceWith(split);

    split.dataset.splitMounted = '1';
    document.body.classList.add('dash-tasks-split-active');

    // Tool-Klicks in der Detailfläche: Nav-Baum für die bleibende IT-Leiste merken
    detail.addEventListener('click', function (ev) {
        const a = ev.target && ev.target.closest ? ev.target.closest('a.dash-task-row[href]') : null;
        if (!a) return;
        const panel = a.closest('.dash-tasks-split__panel');
        persistWorkspaceTree(panel && panel.id ? panel.id : '');
    });

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
        if (found) {
            writeStoredTaskId(id);
            expandGroupForTask(id);
            navButtons.forEach(function (item) {
                const on = item.task.id === id;
                if (item.block) item.block.classList.toggle('is-active', on);
            });
            persistWorkspaceTree(id);
        }
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
            refreshGroupInsightTile(item.task);
            if (item.task.id === 'dashTaskImportVerknuepfen') {
                applyImportTaskPersona(item.task, item.btn);
            }
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
            let expertHit = false;
            task.querySelectorAll('.dash-task-row').forEach(function (row) {
                const rowMatch = sectionMatch || rowMatchesSearch(row, q);
                if (searching) row.hidden = !rowMatch;
                else row.hidden = false;
                if (rowMatch) anyRow = true;
                const details = row.closest('details.dash-task-expert');
                if (details && rowMatch && searching) {
                    details.open = true;
                    expertHit = true;
                }
            });
            const navVisible = !searching || sectionMatch || anyRow;
            item.btn.hidden = !navVisible;
            if (item.block) item.block.hidden = !navVisible;
            if (item.sub && searching && (sectionMatch || expertHit || anyRow)) {
                const moreWrap = item.sub.querySelector('.dash-tasks-split__nav-more');
                if (moreWrap && (expertHit || sectionMatch)) {
                    const moreBtn = moreWrap.querySelector('.dash-tasks-split__nav-more-btn');
                    const moreBody = moreWrap.querySelector('.dash-tasks-split__nav-more-body');
                    if (moreBtn && moreBody) {
                        moreBtn.setAttribute('aria-expanded', 'true');
                        moreBody.hidden = false;
                        moreWrap.classList.add('is-open');
                        const label = moreBtn.querySelector('.dash-tasks-split__nav-more-label');
                        if (label) label.textContent = 'Weniger';
                    }
                }
            }
            if (navVisible && searching) {
                hits += 1;
                if (!firstHit) firstHit = task.id;
            }
        });

        navGroups.forEach(function (group) {
            const anyVisible = group.taskIds.some(function (tid) {
                const item = navButtons.find(function (n) {
                    return n.task.id === tid;
                });
                return item && !item.btn.hidden;
            });
            group.el.hidden = searching && !anyVisible;
            if (searching && anyVisible) setGroupOpen(group, true, false);
        });

        if (searching) {
            if (firstHit) selectTask(firstHit);
            if (hits > 0) tasksRoot.setAttribute('data-dash-task-search-visible', '1');
            else tasksRoot.removeAttribute('data-dash-task-search-visible');
        } else {
            tasksRoot.setAttribute('data-dash-task-search-visible', '1');
            navButtons.forEach(function (item) {
                item.btn.hidden = false;
                if (item.block) item.block.hidden = false;
            });
            navGroups.forEach(function (group) {
                group.el.hidden = false;
                setGroupOpen(group, groupOpenState[group.id] === true, false);
            });
            const current = navButtons.find(function (item) {
                return item.btn.getAttribute('aria-current') === 'true';
            });
            if (current) expandGroupForTask(current.task.id);
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

    mountTaskScenarioGuide(detail, { select: selectTask });

    const api = {
        select: selectTask,
        refreshNavBadges,
        refreshToolCards,
        applySearch
    };
    window.ms365DashTasksSplit = api;
    window.__ms365DashTasksSplitRefreshBadges = refreshNavBadges;
    window.__ms365DashTasksSplitRefreshToolCards = refreshToolCards;

    ['ms365-school-data-written', 'ms365-browser-backup-status', 'ms365-auth-state-changed'].forEach(
        function (evName) {
            try {
                window.addEventListener(evName, function () {
                    refreshNavBadges();
                });
            } catch {
                /* ignore */
            }
        }
    );

    function syncWorkspaceSessionWithDashboardMode(ev) {
        if (isDashboardItShellActive()) {
            persistWorkspaceTree(currentSelectedTaskId());
            return;
        }
        const mode = ev && ev.detail && ev.detail.layoutMode ? String(ev.detail.layoutMode) : '';
        if (/^apps-/.test(mode)) {
            clearItWorkspaceNav();
        } else {
            deactivateItWorkspace();
        }
    }

    ['ms365-dash-app-only-changed', 'ms365-it-shell-changed'].forEach(function (evName) {
        try {
            window.addEventListener(evName, syncWorkspaceSessionWithDashboardMode);
        } catch {
            /* ignore */
        }
    });

    return api;
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', mountDashboardTasksSplit);
    } else {
        mountDashboardTasksSplit();
    }
}
