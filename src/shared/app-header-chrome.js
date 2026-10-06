/**
 * Einheitlicher App-Sticky-Header (Dashboard + alle Werkzeug-Seiten).
 *
 * Zonen (Look-and-Feel):
 * A – Sticky-Leiste (#dashCompactHeader): Logo, Suche, Datensicherung, Schulregister, Konto
 * B – optional Begrüßung (nur Dashboard, #dashGreetingBand)
 * C – Dashboard-Inhalt / Katalog
 * D – Werkzeug-Subheader (.app-tool-chrome): ← Dashboard, Titel, Hilfe, Nav-Pills
 *
 * Werkzeug-Seiten: app-shell-boot.js + app-tool-chrome.js (via pin-gate).
 */

/** @param {string} [file] */
export function resolveAppRootHref(file) {
    const name = String(file || '').replace(/^\//, '');
    try {
        const p = String(window.location.pathname || '/').split('?')[0].split('#')[0];
        const norm = p.replace(/\\/g, '/');
        const iTools = norm.toLowerCase().indexOf('/tools/');
        if (iTools !== -1) return norm.slice(0, iTools) + '/' + name;
        const slash = norm.lastIndexOf('/');
        const dir = slash >= 0 ? norm.slice(0, slash + 1) : '/';
        return dir + name;
    } catch {
        return name;
    }
}

/** @typedef {'dashboard' | 'tool'} AppHeaderMode */

/**
 * @param {Document} [doc]
 */
export function detectAppHeaderMode(doc) {
    const d = doc || (typeof document !== 'undefined' ? document : null);
    if (!d || !d.body) return 'tool';
    return d.body.classList.contains('page-dashboard') ? 'dashboard' : 'tool';
}

/**
 * @param {AppHeaderMode} mode
 */
export function buildAppChromeHeaderHtml(mode) {
    const indexHref = resolveAppRootHref('index.html');
    const logoTeal = resolveAppRootHref('assets/schooltool-teal.png');
    const logoClassic = resolveAppRootHref('assets/schooltool-classic.png');
    const isDashboard = mode === 'dashboard';

    const searchLink =
        '<a class="dash-compact-header__search-dash-link" href="' +
        indexHref +
        '#werkzeuge" title="Zum Dashboard – Aufgaben und Werkzeuge suchen">' +
        '<i class="bi bi-search" aria-hidden="true"></i>' +
        '<span class="dash-compact-header__search-dash-link-text">Suchen …</span>' +
        '</a>';

    const navBlock = isDashboard
        ? '<div class="dash-compact-header__nav" id="dashHeaderNavMount" data-dash-section-audience="it" aria-label="Dashboard-Ansicht"></div>'
        : '<div class="dash-compact-header__nav dash-compact-header__nav--spacer" aria-hidden="true"></div>';

    const searchSlot = isDashboard
        ? '<div class="dash-compact-header__search-slot" id="dashHeaderSearchSlot" data-dash-section-audience="it"></div>'
        : '<div class="dash-compact-header__search-slot" id="dashHeaderSearchSlot" data-dash-section-audience="it">' +
          searchLink +
          '</div>';

    return (
        '<header class="header dash-compact-header" id="dashCompactHeader" data-app-header-mode="' +
        mode +
        '">' +
        '<div class="dash-compact-header__bar">' +
        '<a class="dash-compact-header__brand header-logo" href="' +
        indexHref +
        '" title="MS365-Schul-Tools – Start">' +
        '<img class="app-brand-logo__mark app-brand-logo__mark--teal" src="' +
        logoTeal +
        '" width="28" height="28" alt="" decoding="async" aria-hidden="true" />' +
        '<img class="app-brand-logo__mark app-brand-logo__mark--classic" src="' +
        logoClassic +
        '" width="28" height="28" alt="" decoding="async" aria-hidden="true" />' +
        '<span class="app-name">MS365-Schul-Tools</span></a>' +
        navBlock +
        '<div class="dash-compact-header__actions" id="adminAppTopActions">' +
        searchSlot +
        '<div id="dashRegisterToolbar" hidden aria-hidden="true">' +
        '<input type="file" id="browserBackupImportFile" class="ts-register-save-bar__file" accept="application/json,.json" hidden aria-hidden="true">' +
        '<button type="button" id="browserBackupExport" data-ms365-backup="export" hidden tabindex="-1"></button>' +
        '</div></div></div></header>'
    );
}

/**
 * Lädt Header-CSS app-weit (Kompakt-Header + UI-Spec Zone C).
 */
export function ensureAppChromeStylesheets() {
    if (typeof document === 'undefined' || !document.head) return;

    function linkCss(href, marker) {
        if (document.querySelector('link[' + marker + ']')) return;
        const existing = Array.prototype.slice.call(document.querySelectorAll('link[rel="stylesheet"]'));
        const base = href.split('?')[0];
        if (existing.some(function (el) { return String(el.getAttribute('href') || '').indexOf(base) !== -1; })) {
            return;
        }
        const link = document.createElement('link');
        link.rel = 'stylesheet';
        link.href = href;
        link.setAttribute(marker, '1');
        document.head.appendChild(link);
    }

    linkCss(resolveAppRootHref('src/shared/dashboard-compact-header.css') + '?v=8', 'data-dash-compact-header-css');
    linkCss(resolveAppRootHref('src/shared/dashboard-ui-spec.css') + '?v=6', 'data-dash-ui-spec-css');
    linkCss(resolveAppRootHref('src/shared/app-tool-chrome.css') + '?v=4', 'data-app-tool-chrome-css');
    linkCss(resolveAppRootHref('src/shared/dashboard-card-surfaces.css') + '?v=3', 'data-dash-card-surfaces-css');
}

/**
 * @param {HTMLElement|null} container
 */
export function normalizeAppHeaderActionsOrder(container) {
    if (!container) return;
    const search = container.querySelector('#dashHeaderSearchSlot');
    const backup = document.getElementById('ms365BackupHeader');
    const reg = document.getElementById('ms365HeaderSchulregister');
    const auth = document.getElementById('ms365AuthWidget');
    const toolbar = document.getElementById('dashRegisterToolbar');

    [search, backup, reg, auth, toolbar].forEach(function (el) {
        if (el && el.parentElement === container) {
            container.appendChild(el);
        }
    });
}

export function markToolSubheader() {
    const container = document.querySelector('.container.page-card, .container');
    if (!container) return;
    const legacy = container.querySelector(':scope > .header:not(.dash-compact-header)');
    if (legacy) legacy.classList.add('app-tool-subheader');
}
