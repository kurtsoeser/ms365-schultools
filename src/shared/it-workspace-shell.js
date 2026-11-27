/**
 * Bleibende IT-Linke-Leiste auf Werkzeugseiten (Admin-Center-Muster).
 * Darf ausschließlich im bewussten IT-Modus (`isIt` + Shell `full`) erscheinen –
 * niemals für Lehrkräfte, Schüler/innen oder in der Portal-Ansicht.
 */
import { readItShellChoice } from './dashboard-layout-mode.js';
import { showItAppChrome } from './frontend-planner-chrome-policy.js';
import {
    itNavGroupShouldOpen,
    readItNavGroupOpenState,
    readItNavMoreOpenState,
    rememberItNavGroupOpen,
    rememberItNavMoreOpen
} from './it-nav-open-state.js';
import {
    clearItWorkspaceNav,
    deactivateItWorkspace,
    defaultItWorkspaceTree,
    ensureItWorkspaceNavReady,
    IT_WORKSPACE_FLAG_KEY,
    IT_WORKSPACE_STAMMDATEN_TASK_ID,
    isItWorkspaceActive,
    isTenantWorkspaceHref,
    loadItWorkspaceNav,
    canonicalizeWorkspaceHref,
    isItWorkspaceEmbedMode,
    normalizeWorkspaceHref,
    persistItWorkspaceNav,
    resolveWorkspaceHrefForPage,
    sanitizeItWorkspaceTree,
    setItWorkspaceEmbedMode,
    workspaceHrefsMatch
} from './it-workspace-nav.js';
import {
    bootItWorkspaceStammdatenEmbed,
    unmountItWorkspaceStammdatenEmbed
} from './it-workspace-stammdaten-embed.js';
import { parseRegisterHash } from './stammdaten-register-mode.js';

/** @type {ReturnType<typeof setTimeout>|0} */
let mainSkelSafetyTimer = 0;

function leafFromPathname(pathname) {
    return String(pathname || '')
        .replace(/\\/g, '/')
        .split('/')
        .filter(Boolean)
        .pop() || '';
}

/** Ohne Import-Zyklus zu portal-workspace-shell. */
function dismissPortalWorkspaceShell() {
    try {
        const rail = document.getElementById('portalWorkspaceRail');
        if (rail) rail.remove();
        document.body.classList.remove('portal-workspace-shell');
        document.querySelectorAll('[data-portal-workspace-shell-css]').forEach(function (n) {
            /* stylesheets may stay */
            void n;
        });
    } catch {
        /* ignore */
    }
}

const SHELL_CSS_HREF = 'src/shared/it-workspace-shell.css?v=10';
const PERSONA_CACHE_KEY = 'ms365-dashboard-persona-cache-v2';
const EXCLUDED_PAGE_RE =
    /\/(?:index|welcome|admin|elternbrief-lesen|einsammler-antwort|krankmeldung-formular|schulsprecherwahl-stimme|kollegium-umfragen-antwort)\.html(?:\?|#|$)/i;

/** Startseite der App (nicht die Landingpage). Dort ersetzt die IT-Leiste die Aufgaben-Spalte. */
export function isItDashboardHomePath() {
    try {
        const path = String(location.pathname || '').replace(/\\/g, '/');
        if (/\/landing\//i.test(path)) return false;
        const parts = path.split('/').filter(Boolean);
        if (!parts.length) return true;
        return /^index\.html$/i.test(parts[parts.length - 1] || '');
    } catch {
        return false;
    }
}

const DASHBOARD_HASH_TASK = {
    import: 'dashTaskImportVerknuepfen',
    mitgliedschaften: 'dashTaskGruppen',
    klassen: 'dashTaskUnterricht',
    personen: 'dashTaskPersonen',
    schuljahresstart: 'dashTaskSchuljahr',
    intranet: 'dashTaskIntranet',
    schulapps: 'dashTaskSchulApps',
    aufraeumen: 'dashTaskOrdnung',
    sonstiges: 'dashTaskSchulApps'
};

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

/**
 * Markiert das passende Untertool in der bereits gemounteten Rail (ohne Remount).
 * @param {string} targetHref
 */
function markItWorkspaceSubitemCurrent(targetHref) {
    const nav = document.getElementById('itWorkspaceNav');
    if (!nav) return;
    nav.querySelectorAll('.it-workspace-shell__subitem[aria-current]').forEach(function (el) {
        el.removeAttribute('aria-current');
    });
    nav.querySelectorAll('a.it-workspace-shell__subitem').forEach(function (a) {
        const href = a.getAttribute('href') || '';
        if (workspaceHrefsMatch(href, targetHref)) {
            a.setAttribute('aria-current', 'page');
        }
    });
}

/**
 * Mitgliedschaften / Aufräumen auf derselben HTML-Seite ohne Full-Reload
 * (z. B. verwaltung?kind=, tenant.html#…, leere-gruppen-report?problem=).
 * @param {string} targetHref
 * @returns {boolean}
 */
export function trySoftItWorkspaceNavigate(targetHref) {
    try {
        const current = currentPageHref();
        const target = normalizeWorkspaceHref(targetHref);
        if (!current || !target) return false;
        const fileOf = function (s) {
            return String(canonicalizeWorkspaceHref(s) || s)
                .replace(/#.*$/, '')
                .split('?')[0]
                .toLowerCase();
        };
        if (fileOf(current) !== fileOf(target)) return false;
        if (canonicalizeWorkspaceHref(current) === canonicalizeWorkspaceHref(target) && current === target) {
            return false;
        }

        const u = new URL(String(targetHref || ''), location.href);
        const leaf = fileOf(target);

        if (/verwaltung\.html$/i.test(leaf)) {
            const kind = String(u.searchParams.get('kind') || '');
            if (kind !== 'schulleitung' && kind !== 'verwaltung') return false;
            history.pushState({ ms365ItSoftNav: true }, '', u.pathname + u.search + (u.hash || ''));
            if (typeof window.ms365VerwaltungNavigateKind === 'function') {
                window.ms365VerwaltungNavigateKind(kind);
            } else {
                const btn = document.querySelector(
                    '#slgListItems [data-vw-kind="' + kind + '"]'
                );
                if (!btn) return false;
                btn.click();
            }
            markItWorkspaceSubitemCurrent(target);
            return true;
        }

        // Stammdaten-Tabs: nur Hash wechseln, Seite bleibt.
        if (/^tenant\.html$/i.test(leaf) && u.hash) {
            if (location.hash === u.hash) {
                markItWorkspaceSubitemCurrent(target);
                return true;
            }
            history.pushState({ ms365ItSoftNav: true }, '', u.pathname + u.search + u.hash);
            try {
                window.dispatchEvent(new HashChangeEvent('hashchange'));
            } catch {
                window.dispatchEvent(new Event('hashchange'));
            }
            markItWorkspaceSubitemCurrent(target);
            return true;
        }

        // Leere-Gruppen-Report: Problem-Filter per Query – Seite neu laden wäre Flash;
        // wenn die Seite einen Soft-Handler anbietet, nutzen.
        if (/leere-gruppen-report\.html$/i.test(leaf)) {
            const problem = String(u.searchParams.get('problem') || '');
            if (
                problem &&
                typeof window.ms365LeereGruppenNavigateProblem === 'function'
            ) {
                history.pushState({ ms365ItSoftNav: true }, '', u.pathname + u.search + (u.hash || ''));
                window.ms365LeereGruppenNavigateProblem(problem);
                markItWorkspaceSubitemCurrent(target);
                return true;
            }
        }

        return false;
    } catch {
        return false;
    }
}

function showItMainSkeleton() {
    const main = document.getElementById('itWorkspaceMain');
    if (!main) return;
    let sk = document.getElementById('ms365ItMainSkeleton');
    if (!sk) {
        sk = document.createElement('div');
        sk.id = 'ms365ItMainSkeleton';
        sk.className = 'it-main-skel';
        sk.setAttribute('role', 'status');
        sk.setAttribute('aria-busy', 'true');
        sk.setAttribute('aria-label', 'Werkzeug wird geladen');
        sk.innerHTML =
            '<span class="it-shell-skel__bar it-shell-skel__bar--title"></span>' +
            '<span class="it-shell-skel__bar it-shell-skel__bar--lead"></span>' +
            '<span class="it-shell-skel__bar it-shell-skel__bar--row"></span>' +
            '<span class="it-shell-skel__bar it-shell-skel__bar--row"></span>' +
            '<span class="it-shell-skel__bar it-shell-skel__bar--row is-short"></span>';
        main.appendChild(sk);
    }
    sk.hidden = false;
    if (mainSkelSafetyTimer) clearTimeout(mainSkelSafetyTimer);
    mainSkelSafetyTimer = setTimeout(function () {
        mainSkelSafetyTimer = 0;
        hideItMainSkeleton();
    }, 10000);
}

function hideItMainSkeleton() {
    if (mainSkelSafetyTimer) {
        clearTimeout(mainSkelSafetyTimer);
        mainSkelSafetyTimer = 0;
    }
    const sk = document.getElementById('ms365ItMainSkeleton');
    if (sk) sk.hidden = true;
}

/**
 * Stammdaten schon im iframe: nur Tab/Hash umschalten – kein Reload, kein Skeleton.
 * @param {HTMLIFrameElement} frame
 * @param {URL} abs
 * @returns {boolean}
 */
function tryActivateTenantTabInExistingFrame(frame, abs) {
    if (!frame || !/^tenant\.html$/i.test(leafFromPathname(abs.pathname))) return false;
    let win;
    let loc;
    try {
        win = frame.contentWindow;
        loc = win && win.location;
        if (!loc || !/^tenant\.html$/i.test(leafFromPathname(loc.pathname))) return false;
    } catch {
        return false;
    }
    try {
        const doc = frame.contentDocument;
        quietEmbedDocument(doc);
        const nextHash = abs.hash || '';
        if ((loc.hash || '') !== nextHash) {
            try {
                win.history.replaceState(null, '', abs.pathname + abs.search + nextHash);
            } catch {
                win.location.hash = String(nextHash).replace(/^#/, '');
            }
        }
        const parsed = parseRegisterHash(nextHash);
        if (parsed.tabBtnId) {
            try {
                win.dispatchEvent(
                    new CustomEvent('ms365-register-tab-request', {
                        detail: { tabBtnId: parsed.tabBtnId }
                    })
                );
            } catch {
                /* ignore */
            }
            if (typeof win.ms365TenantTabActivate === 'function') {
                win.ms365TenantTabActivate(parsed.tabBtnId);
            }
        } else {
            try {
                win.dispatchEvent(new HashChangeEvent('hashchange'));
            } catch {
                win.dispatchEvent(new Event('hashchange'));
            }
        }
        return true;
    } catch {
        return false;
    }
}

function quietEmbedCssText() {
    return (
        'html[data-ms365-it-embed] #dashCompactHeader,' +
        'html[data-ms365-it-embed] #itWorkspaceShell,' +
        'html[data-ms365-it-embed] #portalWorkspaceRail,' +
        'html[data-ms365-it-embed] #ms365ItShellSkeleton,' +
        'html[data-ms365-it-embed] #ms365AuthWidget,' +
        'html[data-ms365-it-embed] #ms365FixedFooter,' +
        'html[data-ms365-it-embed] #siteCreditKurtrocks,' +
        'html[data-ms365-it-embed] .app-footer-help,' +
        'html[data-ms365-it-embed] #ms365LicenseGateOverlay,' +
        'html[data-ms365-it-embed] .ms365-license-gate{display:none!important;}' +
        'html[data-ms365-it-embed] body{padding:0!important;}' +
        'html[data-ms365-it-embed] body > .container,' +
        'html[data-ms365-it-embed] body > .container.page-card{' +
        'max-width:none!important;margin:0!important;border-radius:0!important;box-shadow:none!important;}' +
        'html[data-ms365-it-embed] body > .container > .header:not(.dash-compact-header){' +
        'background:transparent!important;color:var(--heading)!important;box-shadow:none!important;}' +
        'html[data-ms365-it-embed] .dash-tasks-split__nav,' +
        'html[data-ms365-it-embed] .dash-tasks-split__resize{display:none!important;}' +
        'html[data-ms365-it-embed] .dash-tasks-split{display:block!important;grid-template-columns:minmax(0,1fr)!important;}' +
        'html[data-ms365-it-embed] .dash-tasks-split__detail{grid-column:1!important;}' +
        'html[data-ms365-it-embed] .ts-admin-rail{display:none!important;}' +
        'html[data-ms365-it-embed] .ts-register-shell.ts-admin-shell{grid-template-columns:minmax(0,1fr)!important;}'
    );
}

function quietEmbedDocument(doc) {
    if (!doc || !doc.documentElement) return;
    try {
        doc.documentElement.setAttribute('data-ms365-it-embed', '');
        if (doc.title) document.title = doc.title;
        let style = doc.getElementById('ms365-it-embed-quiet-css');
        const css = quietEmbedCssText();
        if (style) {
            const cur = String(style.textContent || '');
            if (cur.indexOf('.ts-admin-rail') === -1 || cur.indexOf('.app-footer-help') === -1) {
                style.textContent = css;
            }
            return;
        }
        style = doc.createElement('style');
        style.id = 'ms365-it-embed-quiet-css';
        style.textContent = css;
        doc.head.appendChild(style);
    } catch {
        /* ignore */
    }
}

/**
 * Startseite und Werkzeuge tauschen nur den Inhalt. Willkommen, Admin und Formulare laden neu.
 * @param {string} href
 * @returns {boolean}
 */
export function isItContentSwapHref(href) {
    try {
        const abs = new URL(String(href || ''), 'https://app.ms365.schule');
        const path = abs.pathname.replace(/\\/g, '/');
        if (/\/landing\//i.test(path)) return false;
        const leaf = (path.split('/').filter(Boolean).pop() || '').toLowerCase();
        if (leaf === 'welcome.html' || leaf === 'admin.html') return false;
        if (
            /^(elternbrief-lesen|einsammler-antwort|krankmeldung-formular|schulsprecherwahl-stimme|kollegium-umfragen-antwort)\.html$/i.test(
                leaf
            )
        ) {
            return false;
        }
        return true;
    } catch {
        return false;
    }
}

/**
 * Hebt den passenden Eintrag hervor, ohne die Leiste neu zu bauen.
 * Andere Bereiche bleiben so offen oder zu, wie sie sind.
 * @param {string} nextUrl
 */
function syncRailToWorkspaceHref(nextUrl) {
    const nav = document.getElementById('itWorkspaceNav');
    if (!nav) return;
    const norm = normalizeWorkspaceHref(nextUrl);
    const treeNow = defaultItWorkspaceTree();
    let taskId = resolveActiveWorkspaceTaskId(treeNow, norm);
    if (!taskId) {
        const hash = String(nextUrl || '').split('#')[1] || '';
        taskId = DASHBOARD_HASH_TASK[hash.trim()] || '';
    }
    nav.querySelectorAll('.it-workspace-shell__item[aria-current="page"]').forEach(function (el) {
        el.removeAttribute('aria-current');
    });
    markItWorkspaceSubitemCurrent(norm);
    if (taskId) persistItWorkspaceNav(treeNow, taskId);
    const block = taskId
        ? nav.querySelector('.it-workspace-shell__block[data-task-id="' + taskId + '"]')
        : null;
    if (!block) return;
    const parent = block.querySelector(':scope > a.it-workspace-shell__item');
    const subCurrent = block.querySelector('.it-workspace-shell__subitem[aria-current]');
    if (!subCurrent && parent) parent.setAttribute('aria-current', 'page');
    const sub = block.querySelector(':scope > .it-workspace-shell__sub');
    if (sub && sub.hidden) {
        sub.hidden = false;
        if (parent) parent.setAttribute('aria-expanded', 'true');
        block.classList.remove('is-collapsed');
        rememberItNavGroupOpen(taskId, true);
    }
}

/**
 * Inhalt wechseln, Leiste stehen lassen. Startseite und Werkzeuge laufen im iframe.
 * @param {string} targetHref
 * @returns {boolean}
 */
export function trySoftItWorkspaceFrameNavigate(targetHref) {
    try {
        const main = document.getElementById('itWorkspaceMain');
        const rail = document.getElementById('itWorkspaceNav');
        if (!main || !document.getElementById('itWorkspaceShell') || !rail) return false;

        const abs = new URL(String(targetHref || ''), location.href);
        if (abs.origin !== location.origin) return false;
        if (!isItContentSwapHref(abs.href)) return false;

        dismissPortalWorkspaceShell();
        setItWorkspaceEmbedMode(true);

        let frame = document.getElementById('itWorkspaceFrame');
        if (!frame) {
            while (main.firstChild) main.removeChild(main.firstChild);
            frame = document.createElement('iframe');
            frame.id = 'itWorkspaceFrame';
            frame.className = 'it-workspace-shell__frame';
            frame.title = 'IT-Werkzeug';
            frame.setAttribute('aria-label', 'IT-Werkzeug');
            main.appendChild(frame);
            frame.addEventListener('load', function () {
                hideItMainSkeleton();
                try {
                    quietEmbedDocument(frame.contentDocument);
                    const inner = frame.contentWindow && frame.contentWindow.location;
                    const next = inner ? inner.pathname + inner.search + inner.hash : '';
                    const here = location.pathname + location.search + location.hash;
                    if (next && next !== here) {
                        history.pushState({ ms365ItFrame: true, href: next }, '', next);
                    }
                    syncRailToWorkspaceHref(next || here);
                } catch {
                    /* ignore */
                }
            });
        }

        const nextUrl = abs.pathname + abs.search + abs.hash;
        history.pushState({ ms365ItFrame: true, href: nextUrl }, '', nextUrl);
        syncRailToWorkspaceHref(nextUrl);

        /* Stammdaten: Hash/Tab im bestehenden Frame – sonst hängt Skeleton ohne load */
        if (tryActivateTenantTabInExistingFrame(frame, abs)) {
            hideItMainSkeleton();
            return true;
        }

        let sameLoc = false;
        try {
            const loc = frame.contentWindow && frame.contentWindow.location;
            if (loc) {
                sameLoc =
                    loc.pathname === abs.pathname &&
                    loc.search === abs.search &&
                    loc.hash === abs.hash;
            }
        } catch {
            sameLoc = false;
        }
        if (!sameLoc) {
            try {
                const cur = new URL(frame.getAttribute('src') || '', location.href);
                sameLoc =
                    cur.pathname === abs.pathname &&
                    cur.search === abs.search &&
                    cur.hash === abs.hash;
            } catch {
                sameLoc = false;
            }
        }

        if (sameLoc) {
            hideItMainSkeleton();
            try {
                quietEmbedDocument(frame.contentDocument);
            } catch {
                /* ignore */
            }
        } else {
            showItMainSkeleton();
            frame.src = abs.href;
        }
        return true;
    } catch {
        setItWorkspaceEmbedMode(false);
        return false;
    }
}

function wireItWorkspaceFramePopstate() {
    if (typeof window === 'undefined' || window.__ms365ItFramePopstate) return;
    window.__ms365ItFramePopstate = true;
    window.addEventListener('popstate', function (ev) {
        const st = ev && ev.state;
        if (!st || !st.ms365ItFrame) return;
        const frame = document.getElementById('itWorkspaceFrame');
        if (!frame) return;
        const href = st.href || location.pathname + location.search + location.hash;
        setItWorkspaceEmbedMode(true);
        markItWorkspaceSubitemCurrent(normalizeWorkspaceHref(href));
        let abs;
        try {
            abs = new URL(href, location.href);
        } catch {
            return;
        }
        if (tryActivateTenantTabInExistingFrame(frame, abs)) {
            hideItMainSkeleton();
            return;
        }
        showItMainSkeleton();
        frame.onload = function () {
            hideItMainSkeleton();
            try {
                quietEmbedDocument(frame.contentDocument);
            } catch {
                /* ignore */
            }
        };
        frame.src = abs.href;
    });
}

function appRootHref(href) {
    let raw = String(href || '').trim();
    if (!raw || /^(https?:|mailto:|javascript:)/i.test(raw)) return raw;
    while (raw.startsWith('../')) raw = raw.slice(3);
    if (raw.startsWith('/')) return raw;
    return '/' + raw;
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
    const base = '/index.html';
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

function hasItWorkspaceSessionFlag() {
    try {
        return sessionStorage.getItem(IT_WORKSPACE_FLAG_KEY) === '1';
    } catch {
        return false;
    }
}

/** Theme-boot / laufende Shell: kurz „IT halten“, auch wenn Choice noch portal liest. */
function isItWorkspaceTransitionPending() {
    try {
        if (document.getElementById('itWorkspaceShell')) return true;
        if (document.documentElement.hasAttribute('data-ms365-it-fouc')) return true;
        if (
            hasItWorkspaceSessionFlag() &&
            document.documentElement.hasAttribute('data-dash-it-shell')
        ) {
            return true;
        }
    } catch {
        /* ignore */
    }
    return false;
}

/**
 * Erkennt synchron, ob das aktuelle Konto / die aktuelle Ansicht sicher NICHT
 * im echten IT-Modus läuft (z. B. Portal-Ansicht, Lehrkraft, Schüler/in, abgemeldet).
 * Wichtig: pending-directory / unknownUser darf die IT-Leiste NICHT abreißen –
 * sonst Flash Lehrer-Portal + Anmelden beim Tool-Wechsel.
 * @returns {boolean}
 */
function isDisallowedForItWorkspace() {
    // Explizite IT-Arbeitsfläche nicht durch Lehrer-Persona oder Auth-Flicker abreißen.
    try {
        if (readItShellChoice() === 'full') return false;
    } catch {
        /* ignore */
    }

    function isItPersona(p) {
        return !!(p && (p.isIt || p.globalAdmin || p.designatedSchoolIt || p.operator));
    }

    /** Verzeichnis/Persona noch nicht entschieden – IT-Shell stehen lassen. */
    function isPersonaUnresolved(p) {
        if (!p || !p.loggedIn) return true;
        if (p.unknownUser) return true;
        const src = String(p.personaSource || '');
        if (/^pending/i.test(src)) return true;
        return false;
    }

    function isOperatorNow() {
        try {
            return !!(
                typeof window !== 'undefined' &&
                window.ms365OperatorAccess &&
                typeof window.ms365OperatorAccess.isCurrentUserOperator === 'function' &&
                window.ms365OperatorAccess.isCurrentUserOperator()
            );
        } catch {
            return false;
        }
    }

    try {
        if (readItShellChoice() !== 'full') {
            // Ausnahme: IT-Backend-Seiten brauchen die globale IT-Leiste trotzdem.
            if (isItBackendPageNeedingRail()) {
                /* keep going – Persona prüfen */
            } else if (isItWorkspaceTransitionPending()) {
                // Seitenwechsel in der IT-Arbeitsfläche: nicht unmounten (FOUC-/Auth-Race).
            } else {
                deactivateItWorkspace();
                return true;
            }
        }
    } catch {
        if (!isItWorkspaceTransitionPending()) return true;
    }

    if (isOperatorNow()) return false;

    try {
        const last = typeof window !== 'undefined' ? window.__ms365DashboardPersonasLast : null;
        if (last && last.loggedIn) {
            if (isPersonaUnresolved(last)) return false;
            const allowBackend = isItBackendPageNeedingRail() && isItPersona(last);
            if (!showItAppChrome(last) && !allowBackend) {
                // Bewusste Portal-Wahl bei IT-Persona: nur Flag, Nav behalten.
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
                if (isPersonaUnresolved(p)) return false;
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
    // Eingebettetes Tool in der IT-iframe: keine zweite Shell.
    if (isItWorkspaceEmbedMode()) return false;
    const path = String(location.pathname || '').replace(/\\/g, '/');
    const itHome = isItDashboardHomePath() && readItShellChoice() === 'full';
    if ((EXCLUDED_PAGE_RE.test(path) || /\/$/.test(path)) && !itHome) return false;
    if (/\/landing\//i.test(path)) return false;
    if (isDisallowedForItWorkspace()) return false;
    // Nav ggf. aus Fallback (Stammdaten / Portal vorbereiten ohne Dashboard-Umweg).
    try {
        if (
            readItShellChoice() === 'full' ||
            isItBackendPageNeedingRail() ||
            isItWorkspaceTransitionPending()
        ) {
            const ready = ensureItWorkspaceNavReady();
            return !!(ready && ready.tree && ready.tree.length);
        }
    } catch {
        /* fall through */
    }
    if (!isItWorkspaceActive()) return false;
    return true;
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
    try {
        if (typeof window.ms365ClearItShellSkeleton === 'function') {
            window.ms365ClearItShellSkeleton();
        } else {
            const sk = document.getElementById('ms365ItShellSkeleton');
            if (sk) sk.remove();
            document.documentElement.removeAttribute('data-ms365-it-fouc');
        }
    } catch {
        /* ignore */
    }
    unmountItWorkspaceStammdatenEmbed();
    restoreDashboardBackLinks();
    syncDashboardTasksRail();
    return !!shell;
}

/** Startseiten-Spalte ausblenden, sobald dieselbe IT-Leiste steht. */
function syncDashboardTasksRail() {
    const split = document.getElementById('dashTasksSplit');
    if (!split) return;
    const on = !!document.getElementById('itWorkspaceShell');
    split.classList.toggle('dash-tasks-split--shared-rail', on);
    const nav = split.querySelector(':scope > .dash-tasks-split__nav');
    if (nav) {
        nav.hidden = on;
        if (on) nav.setAttribute('aria-hidden', 'true');
        else nav.removeAttribute('aria-hidden');
    }
    const handle = split.querySelector(':scope > .dash-tasks-split__resize');
    if (handle) handle.hidden = on;
}

/**
 * @returns {boolean}
 */
export function mountItWorkspaceShell() {
    wireItWorkspaceLifecycle();
    wireItWorkspaceFramePopstate();
    if (typeof document === 'undefined' || !document.body) return false;
    // iframe-Inhalt: Flag behalten, keine zweite Shell.
    if (isItWorkspaceEmbedMode()) return false;
    if (!shouldMountItWorkspaceShell()) {
        unmountItWorkspaceShell();
        return false;
    }
    dismissPortalWorkspaceShell();
    if (document.getElementById('itWorkspaceShell')) {
        try {
            if (typeof window.ms365ClearItShellSkeleton === 'function') {
                window.ms365ClearItShellSkeleton();
            }
            bootItWorkspaceStammdatenEmbed();
        } catch {
            /* ignore */
        }
        return true;
    }

    const tree = sanitizeItWorkspaceTree(defaultItWorkspaceTree());
    if (!tree.length) return false;
    loadItWorkspaceNav();

    ensureShellStylesheet();

    const inTools = /\/tools\//i.test(String(location.pathname || ''));
    const current = currentPageHref();
    const hashTask = DASHBOARD_HASH_TASK[String(location.hash || '').replace(/^#/, '').trim()] || '';
    const activeTaskId =
        resolveActiveWorkspaceTaskId(tree, current) ||
        (isItDashboardHomePath() ? hashTask : '');

    const shell = document.createElement('div');
    shell.id = 'itWorkspaceShell';
    shell.className = 'it-workspace-shell';
    shell.innerHTML =
        '<aside class="it-workspace-shell__rail" aria-label="IT-Bereiche">' +
        '<div class="it-workspace-shell__brand">' +
        '<a class="it-workspace-shell__home" href="' +
        appRootHref(resolveWorkspaceHrefForPage('index.html', inTools)) +
        '"><i class="bi bi-grid-1x2" aria-hidden="true"></i><span>IT-Übersicht</span></a>' +
        '</div>' +
        '<nav class="it-workspace-shell__nav" id="itWorkspaceNav"></nav>' +
        '</aside>' +
        '<div class="it-workspace-shell__main" id="itWorkspaceMain"></div>';

    const nav = shell.querySelector('#itWorkspaceNav');
    const main = shell.querySelector('#itWorkspaceMain');
    const homeLink = shell.querySelector('.it-workspace-shell__home');
    if (homeLink) {
        homeLink.addEventListener('click', function (ev) {
            if (ev.button !== 0 || ev.metaKey || ev.ctrlKey || ev.shiftKey || ev.altKey) return;
            if (trySoftItWorkspaceFrameNavigate(homeLink.href)) ev.preventDefault();
        });
    }

    const groupOpenState = readItNavGroupOpenState();
    const moreOpenState = readItNavMoreOpenState();

    tree.forEach(function (node) {
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
        const sectionKnown = Object.prototype.hasOwnProperty.call(groupOpenState, node.id);
        const sectionOpen = tools.length && itNavGroupShouldOpen(groupOpenState, node.id, isActiveBlock);
        if (tools.length && !sectionKnown && sectionOpen) {
            groupOpenState[node.id] = true;
            rememberItNavGroupOpen(node.id, true);
        }
        if (tools.length) {
            parent.setAttribute('aria-expanded', sectionOpen ? 'true' : 'false');
            block.classList.toggle('is-collapsed', !sectionOpen);
        }
        block.appendChild(parent);

        if (tools.length) {
            const sub = document.createElement('div');
            sub.className = 'it-workspace-shell__sub';
            sub.hidden = !sectionOpen;
            parent.addEventListener('click', function (ev) {
                const liveSub = block.querySelector(':scope > .it-workspace-shell__sub');
                if (!liveSub) return;
                const currentlyOpen = !liveSub.hidden;
                if (currentlyOpen) {
                    ev.preventDefault();
                    liveSub.hidden = true;
                    parent.setAttribute('aria-expanded', 'false');
                    block.classList.add('is-collapsed');
                    groupOpenState[node.id] = false;
                    rememberItNavGroupOpen(node.id, false);
                    return;
                }
                liveSub.hidden = false;
                parent.setAttribute('aria-expanded', 'true');
                block.classList.remove('is-collapsed');
                groupOpenState[node.id] = true;
                rememberItNavGroupOpen(node.id, true);
            });
            let looseMarked = false;
            /** @type {HTMLElement | null} */
            let moreBody = null;
            const moreKnown = Object.prototype.hasOwnProperty.call(moreOpenState, node.id);
            const activeExpert = isActiveBlock && tools.some(function (tool) {
                return !!tool.expert && (
                    workspaceHrefsMatch(tool.href, current) ||
                    (!exactSubMatch && workspaceHrefsMatch(tool.href, current, { looseFile: true }))
                );
            });
            const moreOpen = moreKnown ? moreOpenState[node.id] === true : !!activeExpert;
            if (!moreKnown && moreOpen) {
                moreOpenState[node.id] = true;
                rememberItNavMoreOpen(node.id, true);
            }
            tools.forEach(function (tool) {
                const isExpert = !!tool.expert;
                if (isExpert && !moreBody) {
                    const moreWrap = document.createElement('div');
                    moreWrap.className = 'it-workspace-shell__more';
                    const moreBtn = document.createElement('button');
                    moreBtn.type = 'button';
                    moreBtn.className = 'it-workspace-shell__more-btn';
                    moreBtn.setAttribute('aria-expanded', moreOpen ? 'true' : 'false');
                    moreBtn.textContent = moreOpen ? 'Weniger' : 'Mehr …';
                    moreBody = document.createElement('div');
                    moreBody.className = 'it-workspace-shell__more-body';
                    moreBody.hidden = !moreOpen;
                    moreBtn.addEventListener('click', function (ev) {
                        ev.preventDefault();
                        ev.stopPropagation();
                        const open = moreBtn.getAttribute('aria-expanded') !== 'true';
                        moreBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
                        moreBody.hidden = !open;
                        moreBtn.textContent = open ? 'Weniger' : 'Mehr …';
                        moreOpenState[node.id] = open;
                        rememberItNavMoreOpen(node.id, open);
                    });
                    moreWrap.appendChild(moreBtn);
                    moreWrap.appendChild(moreBody);
                    sub.appendChild(moreWrap);
                }
                const a = document.createElement('a');
                a.className = 'it-workspace-shell__subitem';
                a.href = appRootHref(resolveWorkspaceHrefForPage(tool.href, inTools));
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
                a.addEventListener('click', function (ev) {
                    persistItWorkspaceNav(tree, node.id);
                    groupOpenState[node.id] = true;
                    rememberItNavGroupOpen(node.id, true);
                    if (
                        ev.defaultPrevented ||
                        ev.button !== 0 ||
                        ev.metaKey ||
                        ev.ctrlKey ||
                        ev.shiftKey ||
                        ev.altKey
                    ) {
                        return;
                    }
                    if (trySoftItWorkspaceNavigate(a.href)) {
                        ev.preventDefault();
                        return;
                    }
                    if (trySoftItWorkspaceFrameNavigate(a.href)) {
                        ev.preventDefault();
                    }
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
    try {
        if (typeof window.ms365ClearItShellSkeleton === 'function') {
            window.ms365ClearItShellSkeleton();
        } else {
            const sk = document.getElementById('ms365ItShellSkeleton');
            if (sk) sk.remove();
            document.documentElement.removeAttribute('data-ms365-it-fouc');
        }
    } catch {
        /* ignore */
    }
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

    syncDashboardTasksRail();
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

function installItSoftNavCapture() {
    if (typeof document === 'undefined' || typeof document.addEventListener !== 'function' || window.__ms365ItSoftNavCapture) {
        return;
    }
    window.__ms365ItSoftNavCapture = true;
    document.addEventListener(
        'click',
        function (ev) {
            if (ev.defaultPrevented || ev.button !== 0 || ev.metaKey || ev.ctrlKey || ev.shiftKey || ev.altKey) {
                return;
            }
            const shellOn =
                !!document.getElementById('itWorkspaceShell') ||
                document.documentElement.hasAttribute('data-dash-it-shell');
            if (!shellOn) return;
            const t = ev.target;
            const a = t && t.closest ? t.closest('a[href]') : null;
            if (!a || a.target === '_blank' || a.hasAttribute('download')) return;
            if (a.classList.contains('it-workspace-shell__item')) {
                const sub = a.parentElement && a.parentElement.querySelector(':scope > .it-workspace-shell__sub');
                if (sub && !sub.hidden) return;
            }
            let abs;
            try {
                abs = new URL(a.href, location.href);
            } catch {
                return;
            }
            if (abs.origin !== location.origin) return;
            const sameDoc = abs.pathname === location.pathname && abs.search === location.search;
            if (sameDoc) {
                const frame = document.getElementById('itWorkspaceFrame');
                let frameHref = '';
                try {
                    const inner = frame && frame.contentWindow && frame.contentWindow.location;
                    frameHref = inner ? inner.pathname + inner.search + inner.hash : '';
                } catch {
                    frameHref = '';
                }
                const target = abs.pathname + abs.search + abs.hash;
                if (!frame || frameHref === target) return;
            }
            if (trySoftItWorkspaceFrameNavigate(a.href)) {
                ev.preventDefault();
                ev.stopPropagation();
            }
        },
        true
    );
}

function wireItWorkspaceLifecycle() {
    if (lifecycleWired || typeof window === 'undefined') return;
    lifecycleWired = true;
    installItSoftNavCapture();

    window.addEventListener('ms365-auth-state-changed', function () {
        if (typeof window.ms365AuthEnsureInitialized === 'function') {
            Promise.resolve(window.ms365AuthEnsureInitialized())
                .then(function () {
                    // Nur bei klarem Logout die IT-Arbeitsfläche beenden – nicht während MSAL-Init.
                    if (
                        typeof window.ms365AuthIsLoggedIn === 'function' &&
                        !window.ms365AuthIsLoggedIn() &&
                        !hasItWorkspaceSessionFlag()
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

    window.addEventListener('ms365-dash-tasks-split-changed', function (ev) {
        if (!isItDashboardHomePath()) return;
        const taskId = ev && ev.detail ? ev.detail.taskId : '';
        const nav = document.getElementById('itWorkspaceNav');
        if (!nav || !taskId) return;
        nav.querySelectorAll('.it-workspace-shell__item[aria-current="page"]').forEach(function (el) {
            el.removeAttribute('aria-current');
        });
        const block = nav.querySelector('.it-workspace-shell__block[data-task-id="' + taskId + '"]');
        const parent = block && block.querySelector(':scope > a.it-workspace-shell__item');
        if (parent) parent.setAttribute('aria-current', 'page');
    });

    window.addEventListener('ms365-it-shell-changed', function (ev) {
        const mode = ev && ev.detail ? ev.detail.mode : '';
        if (mode === 'full') {
            try {
                loadItWorkspaceNav();
                sessionStorage.setItem(IT_WORKSPACE_FLAG_KEY, '1');
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
