/**
 * Bleibende IT-Linke-Leiste auf Werkzeugseiten (Admin-Center-Muster).
 * Darf ausschließlich im bewussten IT-Modus (`isIt` + Shell `full`) erscheinen –
 * niemals für Lehrkräfte, Schüler/innen oder in der Portal-Ansicht.
 */
import { readItShellChoice } from './dashboard-layout-mode.js';
import { showItAppChrome } from './frontend-planner-chrome-policy.js';
import {
    clearItWorkspaceNav,
    deactivateItWorkspace,
    ensureItWorkspaceNavReady,
    IT_WORKSPACE_FLAG_KEY,
    IT_WORKSPACE_STAMMDATEN_TASK_ID,
    isItWorkspaceActive,
    isTenantWorkspaceHref,
    loadItWorkspaceNav,
    normalizeWorkspaceHref,
    persistItWorkspaceNav,
    resolveWorkspaceHrefForPage,
    sanitizeItWorkspaceTree,
    workspaceHrefsMatch
} from './it-workspace-nav.js';
import {
    bootItWorkspaceStammdatenEmbed,
    unmountItWorkspaceStammdatenEmbed
} from './it-workspace-stammdaten-embed.js';

const SHELL_CSS_HREF = 'src/shared/it-workspace-shell.css?v=6';
const PERSONA_CACHE_KEY = 'ms365-dashboard-persona-cache-v2';
const EXCLUDED_PAGE_RE =
    /\/(?:index|welcome|admin|elternbrief-lesen|einsammler-antwort|krankmeldung-formular|schulsprecherwahl-stimme|kollegium-umfragen-antwort)\.html(?:\?|#|$)/i;

let lifecycleWired = false;

function ensureShellStylesheet() {
    if (typeof document === 'undefined') return;
    const id = 'ms365-it-workspace-shell-css';
    if (document.getElementById(id)) return;
    const inTools = /\/tools\//i.test(String(location.pathname || ''));
    const link = document.createElement('link');
    link.id = id;
    link.rel = 'stylesheet';
    link.href = inTools ? '../' + SHELL_CSS_HREF : SHELL_CSS_HREF;
    document.head.appendChild(link);
}

function currentPageHref() {
    const path = String(location.pathname || '').replace(/\\/g, '/');
    const leaf = path.split('/').filter(Boolean).pop() || '';
    let rel = leaf;
    if (/\/tools\//i.test(path)) rel = 'tools/' + leaf;
    else if (/^(tenant|index|dashboard-werkzeug-zugriff)\.html$/i.test(leaf)) rel = leaf;
    return normalizeWorkspaceHref(rel + (location.search || '') + (location.hash || ''));
}

function overviewHrefForTask(taskId) {
    const map = {
        dashTaskImportVerknuepfen: 'import',
        dashTaskGruppen: 'mitgliedschaften',
        dashTaskUnterricht: 'klassen',
        dashTaskPersonen: 'personen',
        dashTaskSchuljahr: 'schuljahresstart',
        dashTaskIntranet: 'intranet',
        dashTaskSchulApps: 'schulapps',
        dashTaskOrdnung: 'aufraeumen'
    };
    const hash = map[taskId] || '';
    const inTools = /\/tools\//i.test(String(location.pathname || ''));
    const base = inTools ? '../index.html' : 'index.html';
    return hash ? base + '#' + hash : base;
}

/**
 * Stammdaten / Portal vorbereiten: IT-Backend-Seiten – IT-Leiste auch behalten,
 * wenn die Kopfzeile gerade auf „Portal“ steht (sonst wirkt die Seite losgelöst).
 */
function isItBackendPageNeedingRail() {
    try {
        const path = String(location.pathname || '').replace(/\\/g, '/');
        const leaf = path.split('/').filter(Boolean).pop() || '';
        return /^(tenant|dashboard-werkzeug-zugriff)\.html$/i.test(leaf);
    } catch {
        return false;
    }
}

/**
 * Erkennt synchron, ob das aktuelle Konto / die aktuelle Ansicht sicher NICHT
 * im echten IT-Modus läuft (z. B. Portal-Ansicht, Lehrkraft, Schüler/in, abgemeldet).
 * @returns {boolean}
 */
function isDisallowedForItWorkspace() {
    try {
        if (readItShellChoice() !== 'full') {
            // Ausnahme: IT-Backend-Seiten brauchen die globale IT-Leiste trotzdem.
            if (!isItBackendPageNeedingRail()) {
                deactivateItWorkspace();
                return true;
            }
        }
    } catch {
        return true;
    }

    function isItPersona(p) {
        return !!(p && (p.isIt || p.globalAdmin || p.designatedSchoolIt));
    }

    try {
        const last = typeof window !== 'undefined' ? window.__ms365DashboardPersonasLast : null;
        if (last && last.loggedIn) {
            const allowBackend = isItBackendPageNeedingRail() && isItPersona(last);
            if (!showItAppChrome(last) && !allowBackend) {
                if (!isItPersona(last)) clearItWorkspaceNav();
                else deactivateItWorkspace();
                return true;
            }
        }
    } catch {
        /* ignore */
    }

    try {
        const raw = sessionStorage.getItem(PERSONA_CACHE_KEY);
        if (raw) {
            const parsed = JSON.parse(raw);
            const p = parsed && parsed.personas;
            if (p && p.loggedIn) {
                if (!isItPersona(p)) {
                    clearItWorkspaceNav();
                    return true;
                }
            }
        }
    } catch {
        /* ignore */
    }

    return false;
}

/**
 * @returns {boolean}
 */
export function shouldMountItWorkspaceShell() {
    if (typeof document === 'undefined' || !document.body) return false;
    const path = String(location.pathname || '').replace(/\\/g, '/');
    if (EXCLUDED_PAGE_RE.test(path) || /\/$/.test(path)) return false;
    if (/\/landing\//i.test(path)) return false;
    if (isDisallowedForItWorkspace()) return false;
    // Nav ggf. aus Fallback (Stammdaten / Portal vorbereiten ohne Dashboard-Umweg).
    try {
        if (readItShellChoice() === 'full' || isItBackendPageNeedingRail()) {
            const ready = ensureItWorkspaceNavReady();
            return !!(ready && ready.tree && ready.tree.length);
        }
    } catch {
        /* fall through */
    }
    if (!isItWorkspaceActive()) return false;
    return !!loadItWorkspaceNav();
}

/**
 * Bestimmt den aktiven Bereich anhand der aktuellen Seiten-URL.
 * Markiert niemals blind einen veralteten Bereich (z. B. Mitgliedschaften auf Elternbrief).
 * @param {any[]} tree
 * @param {string} current
 * @returns {string}
 */
export function resolveActiveWorkspaceTaskId(tree, current) {
    if (!Array.isArray(tree) || !current) return '';
    // Stammdaten ist eigene Gruppe: tenant.html immer dort, nie unter Schuljahresstart.
    if (isTenantWorkspaceHref(current)) {
        const hasStammdaten = tree.some(function (n) {
            return n && String(n.id || '') === IT_WORKSPACE_STAMMDATEN_TASK_ID;
        });
        if (hasStammdaten) return IT_WORKSPACE_STAMMDATEN_TASK_ID;
    }
    for (let i = 0; i < tree.length; i++) {
        const node = tree[i];
        const tools = (node && node.tools) || [];
        for (let j = 0; j < tools.length; j++) {
            if (workspaceHrefsMatch(tools[j].href, current)) {
                return String(node.id || '');
            }
        }
    }
    for (let i = 0; i < tree.length; i++) {
        const node = tree[i];
        const tools = (node && node.tools) || [];
        for (let j = 0; j < tools.length; j++) {
            if (workspaceHrefsMatch(tools[j].href, current, { looseFile: true })) {
                return String(node.id || '');
            }
        }
    }
    return '';
}

/**
 * Entfernt die IT-Linke-Leiste sauber wieder (z. B. nach Wechsel auf Portal-Ansicht oder Nicht-IT-Konto).
 * @returns {boolean}
 */
export function unmountItWorkspaceShell() {
    if (typeof document === 'undefined' || !document.body) return false;
    const shell = document.getElementById('itWorkspaceShell');
    if (shell) {
        const main = shell.querySelector('#itWorkspaceMain');
        const parent = shell.parentElement;
        if (main && parent) {
            while (main.firstChild) {
                parent.insertBefore(main.firstChild, shell);
            }
        }
        shell.remove();
    }
    document.body.classList.remove('it-workspace-shell');
    unmountItWorkspaceStammdatenEmbed();
    restoreDashboardBackLinks();
    return !!shell;
}

/**
 * @returns {boolean}
 */
export function mountItWorkspaceShell() {
    wireItWorkspaceLifecycle();
    if (typeof document === 'undefined' || !document.body) return false;
    if (!shouldMountItWorkspaceShell()) {
        unmountItWorkspaceShell();
        return false;
    }
    if (document.getElementById('itWorkspaceShell')) {
        try {
            bootItWorkspaceStammdatenEmbed();
        } catch {
            /* ignore */
        }
        return true;
    }

    const stored = loadItWorkspaceNav();
    if (!stored || !stored.tree || !stored.tree.length) return false;
    stored.tree = sanitizeItWorkspaceTree(stored.tree);

    ensureShellStylesheet();

    const inTools = /\/tools\//i.test(String(location.pathname || ''));
    const current = currentPageHref();
    const activeTaskId = resolveActiveWorkspaceTaskId(stored.tree, current);

    const shell = document.createElement('div');
    shell.id = 'itWorkspaceShell';
    shell.className = 'it-workspace-shell';
    shell.innerHTML =
        '<aside class="it-workspace-shell__rail" aria-label="IT-Bereiche">' +
        '<div class="it-workspace-shell__brand">' +
        '<a class="it-workspace-shell__home" href="' +
        resolveWorkspaceHrefForPage('index.html', inTools) +
        '"><i class="bi bi-grid-1x2" aria-hidden="true"></i><span>IT-Übersicht</span></a>' +
        '</div>' +
        '<nav class="it-workspace-shell__nav" id="itWorkspaceNav"></nav>' +
        '</aside>' +
        '<div class="it-workspace-shell__main" id="itWorkspaceMain"></div>';

    const nav = shell.querySelector('#itWorkspaceNav');
    const main = shell.querySelector('#itWorkspaceMain');

    stored.tree.forEach(function (node) {
        const block = document.createElement('div');
        block.className = 'it-workspace-shell__block';
        block.dataset.taskId = node.id;
        const isActiveBlock = !!activeTaskId && node.id === activeTaskId;

        const parent = document.createElement('a');
        parent.className = 'it-workspace-shell__item';
        parent.href = overviewHrefForTask(node.id);
        const tools = node.tools || [];
        const exactSubMatch = tools.some(function (t) {
            return workspaceHrefsMatch(t.href, current);
        });
        const looseSubMatch =
            !exactSubMatch &&
            tools.some(function (t) {
                return workspaceHrefsMatch(t.href, current, { looseFile: true });
            });
        if (isActiveBlock && !exactSubMatch && !looseSubMatch) {
            parent.setAttribute('aria-current', 'page');
        }
        parent.innerHTML =
            '<span class="it-workspace-shell__icon" aria-hidden="true"><i class="' +
            (node.icon || 'bi bi-circle') +
            '"></i></span>' +
            '<span class="it-workspace-shell__label"></span>' +
            (tools.length
                ? '<i class="bi bi-chevron-down it-workspace-shell__chevron" aria-hidden="true"></i>'
                : '');
        parent.querySelector('.it-workspace-shell__label').textContent = node.label || node.id;
        if (tools.length) {
            parent.setAttribute('aria-expanded', isActiveBlock ? 'true' : 'false');
        }
        block.appendChild(parent);

        if (tools.length) {
            const sub = document.createElement('div');
            sub.className = 'it-workspace-shell__sub';
            // Nur der aktive Bereich ist aufgeklappt – sonst verliert man die Gruppenzugehörigkeit.
            sub.hidden = !isActiveBlock;
            parent.addEventListener('click', function (ev) {
                const liveSub = block.querySelector(':scope > .it-workspace-shell__sub');
                if (!liveSub) return;
                const open = !liveSub.hidden;
                // Aktiver Bereich: Klick klappt zu/auf, ohne die Seite zu verlassen.
                if (isActiveBlock) {
                    ev.preventDefault();
                    liveSub.hidden = open;
                    parent.setAttribute('aria-expanded', open ? 'false' : 'true');
                    block.classList.toggle('is-collapsed', open);
                    return;
                }
                // Anderer Bereich: Untermenü kurz öffnen, Navigation zur Übersicht bleibt.
                if (!open) liveSub.hidden = false;
                parent.setAttribute('aria-expanded', 'true');
            });
            let looseMarked = false;
            /** @type {HTMLElement | null} */
            let moreBody = null;
            tools.forEach(function (tool) {
                const isExpert = !!tool.expert;
                if (isExpert && !moreBody) {
                    const moreWrap = document.createElement('div');
                    moreWrap.className = 'it-workspace-shell__more';
                    const moreBtn = document.createElement('button');
                    moreBtn.type = 'button';
                    moreBtn.className = 'it-workspace-shell__more-btn';
                    moreBtn.setAttribute('aria-expanded', 'false');
                    moreBtn.textContent = 'Mehr …';
                    moreBody = document.createElement('div');
                    moreBody.className = 'it-workspace-shell__more-body';
                    moreBody.hidden = true;
                    moreBtn.addEventListener('click', function (ev) {
                        ev.preventDefault();
                        ev.stopPropagation();
                        const open = moreBtn.getAttribute('aria-expanded') !== 'true';
                        moreBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
                        moreBody.hidden = !open;
                        moreBtn.textContent = open ? 'Weniger' : 'Mehr …';
                    });
                    moreWrap.appendChild(moreBtn);
                    moreWrap.appendChild(moreBody);
                    sub.appendChild(moreWrap);
                }
                const a = document.createElement('a');
                a.className = 'it-workspace-shell__subitem';
                a.href = resolveWorkspaceHrefForPage(tool.href, inTools);
                const isExact = isActiveBlock && workspaceHrefsMatch(tool.href, current);
                const isLoose =
                    isActiveBlock &&
                    !exactSubMatch &&
                    !looseMarked &&
                    workspaceHrefsMatch(tool.href, current, { looseFile: true });
                if (isExact || isLoose) {
                    if (isLoose) looseMarked = true;
                    a.setAttribute('aria-current', 'page');
                }
                a.innerHTML =
                    '<span class="it-workspace-shell__icon" aria-hidden="true"><i class="' +
                    (tool.icon || 'bi bi-dot') +
                    '"></i></span>' +
                    '<span class="it-workspace-shell__label"></span>';
                a.querySelector('.it-workspace-shell__label').textContent = tool.label;
                a.addEventListener('click', function () {
                    persistItWorkspaceNav(stored.tree, node.id);
                });
                if (isExpert && moreBody) moreBody.appendChild(a);
                else sub.appendChild(a);
            });
            block.appendChild(sub);
        }

        nav.appendChild(block);
    });

    const container =
        document.querySelector('.container.page-card') || document.querySelector('.container');
    if (!container || !container.parentElement) return false;

    const parent = container.parentElement;
    parent.insertBefore(shell, container);
    main.appendChild(container);

    document.body.classList.add('it-workspace-shell', 'app-shell-chrome');
    hideDashboardBackLinks();

    try {
        sessionStorage.setItem(IT_WORKSPACE_FLAG_KEY, '1');
    } catch {
        /* ignore */
    }

    try {
        bootItWorkspaceStammdatenEmbed();
    } catch {
        /* ignore */
    }

    return true;
}

function hideDashboardBackLinks() {
    document.querySelectorAll('a[href]').forEach(function (a) {
        const href = String(a.getAttribute('href') || '').replace(/\\/g, '/');
        if (!/(?:^|\/)index\.html(?:\?|#|$)/i.test(href) && !/\.\.\/index\.html/i.test(href)) return;
        const label = String(a.textContent || '')
            .replace(/\s+/g, ' ')
            .trim();
        if (!/dashboard/i.test(label)) return;
        a.hidden = true;
        a.setAttribute('data-it-workspace-hidden-back', '1');
    });
}

function restoreDashboardBackLinks() {
    document.querySelectorAll('a[data-it-workspace-hidden-back="1"]').forEach(function (a) {
        a.hidden = false;
        a.removeAttribute('data-it-workspace-hidden-back');
    });
}

function syncItWorkspaceShellState() {
    if (!shouldMountItWorkspaceShell()) {
        unmountItWorkspaceShell();
        return;
    }
    mountItWorkspaceShell();
}

function wireItWorkspaceLifecycle() {
    if (lifecycleWired || typeof window === 'undefined') return;
    lifecycleWired = true;

    window.addEventListener('ms365-auth-state-changed', function () {
        if (typeof window.ms365AuthEnsureInitialized === 'function') {
            Promise.resolve(window.ms365AuthEnsureInitialized())
                .then(function () {
                    if (
                        typeof window.ms365AuthIsLoggedIn === 'function' &&
                        !window.ms365AuthIsLoggedIn()
                    ) {
                        deactivateItWorkspace();
                    }
                    syncItWorkspaceShellState();
                })
                .catch(function () {
                    syncItWorkspaceShellState();
                });
            return;
        }
        syncItWorkspaceShellState();
    });

    window.addEventListener('ms365-dashboard-persona-ready', function () {
        syncItWorkspaceShellState();
    });

    window.addEventListener('ms365-it-shell-changed', function (ev) {
        const mode = ev && ev.detail ? ev.detail.mode : '';
        if (mode === 'full') {
            try {
                if (loadItWorkspaceNav()) sessionStorage.setItem(IT_WORKSPACE_FLAG_KEY, '1');
            } catch {
                /* ignore */
            }
        } else {
            deactivateItWorkspace();
        }
        syncItWorkspaceShellState();
    });
}

export function bootItWorkspaceShell() {
    try {
        wireItWorkspaceLifecycle();
        return mountItWorkspaceShell();
    } catch {
        return false;
    }
}

if (typeof window !== 'undefined') {
    window.ms365BootItWorkspaceShell = bootItWorkspaceShell;
}
