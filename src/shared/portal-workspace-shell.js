/**
 * Portal-Leiste auf Werkzeugseiten (Lehrkraft / Schüler/in).
 * Die IT-Leiste bleibt unberührt und hat Vorrang.
 */
import { readItShellChoice } from './dashboard-layout-mode.js';
import { resolveWorkspaceHrefForPage } from './it-workspace-nav.js';
import { shouldMountItWorkspaceShell } from './it-workspace-shell.js';
import {
    buildDefaultPortalTree,
    loadPortalWorkspaceNav,
    rememberPortalSection,
    resolvePortalLocation,
    resolvePortalWorkspaceView
} from './portal-workspace-nav.js';

const SHELL_CSS_HREF = 'src/shared/portal-workspace-shell.css?v=1';
const PERSONA_CACHE_KEY = 'ms365-dashboard-persona-cache-v2';
const EXCLUDED_PAGE_RE =
    /\/(?:index|welcome|admin|elternbrief-lesen|einsammler-antwort|krankmeldung-formular|schulsprecherwahl-stimme|kollegium-umfragen-antwort)\.html(?:\?|#|$)/i;

let lifecycleWired = false;

function ensureShellStylesheet() {
    if (typeof document === 'undefined') return;
    const id = 'ms365-portal-workspace-shell-css';
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
    return rel + (location.search || '') + (location.hash || '');
}

function readCachedPersonas() {
    try {
        const raw = sessionStorage.getItem(PERSONA_CACHE_KEY);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        return parsed && parsed.personas ? parsed.personas : null;
    } catch {
        return null;
    }
}

function isItFullMode() {
    const personas = readCachedPersonas();
    if (personas && personas.loggedIn && personas.isIt && readItShellChoice() === 'full') return true;
    if (document.getElementById('itWorkspaceShell')) return true;
    try {
        if (shouldMountItWorkspaceShell()) return true;
    } catch {
        /* ignore */
    }
    return false;
}

/**
 * @param {string} current
 * @returns {{ view: string, tree: any[], loc: { section: any, tool: any } }|null}
 */
function pickTreeAndLocation(current) {
    const personaView = resolvePortalWorkspaceView(readCachedPersonas());
    const stored = loadPortalWorkspaceNav();
    /** @type {string[]} */
    const views = [];
    if (personaView) views.push(personaView);
    else {
        if (stored && (stored.view === 'lehrer' || stored.view === 'schueler')) views.push(stored.view);
        views.push('lehrer', 'schueler');
    }
    const seen = new Set();
    for (let i = 0; i < views.length; i++) {
        const view = views[i];
        if (seen.has(view)) continue;
        seen.add(view);
        const tree =
            stored && stored.view === view && stored.tree && stored.tree.length
                ? stored.tree
                : buildDefaultPortalTree(view);
        const loc = resolvePortalLocation(tree, current);
        if (loc) return { view: view, tree: tree, loc: loc };
    }
    return null;
}

function shouldMountPortalWorkspace() {
    if (typeof document === 'undefined' || !document.body) return false;
    const path = String(location.pathname || '').replace(/\\/g, '/');
    if (EXCLUDED_PAGE_RE.test(path) || /\/$/.test(path)) return false;
    if (/\/landing\//i.test(path)) return false;
    if (isItFullMode()) return false;
    return !!pickTreeAndLocation(currentPageHref());
}

function safeIconClass(icon) {
    const s = String(icon || '');
    return /^bi bi-[a-z0-9-]+$/i.test(s) ? s : 'bi bi-circle';
}

function safeToolHref(href, inTools) {
    const raw = String(href || '').trim();
    if (!/^(?:tools\/)?[a-z0-9-]+\.html(?:\?[^#]*)?$/i.test(raw)) return '';
    return resolveWorkspaceHrefForPage(raw, inTools);
}

function indexHref(inTools) {
    return resolveWorkspaceHrefForPage('index.html', inTools);
}

function hideDashboardBackLinks() {
    document.querySelectorAll('a[href]').forEach(function (a) {
        if (a.closest('.portal-workspace-shell__context, .portal-workspace-shell__rail')) return;
        const href = String(a.getAttribute('href') || '').replace(/\\/g, '/');
        if (!/(?:^|\/)index\.html(?:\?|#|$)/i.test(href) && !/\.\.\/index\.html/i.test(href)) return;
        const label = String(a.textContent || '')
            .replace(/\s+/g, ' ')
            .trim();
        if (!/dashboard/i.test(label)) return;
        a.hidden = true;
        a.setAttribute('data-portal-workspace-hidden-back', '1');
    });
}

function restoreDashboardBackLinks() {
    document.querySelectorAll('a[data-portal-workspace-hidden-back="1"]').forEach(function (a) {
        a.hidden = false;
        a.removeAttribute('data-portal-workspace-hidden-back');
    });
}

/**
 * @returns {boolean}
 */
export function unmountPortalWorkspaceShell() {
    if (typeof document === 'undefined' || !document.body) return false;
    const rail = document.getElementById('portalWorkspaceRail');
    const context = document.getElementById('portalWorkspaceContext');
    if (rail) rail.remove();
    if (context) context.remove();
    document.body.classList.remove('portal-workspace-shell');
    restoreDashboardBackLinks();
    return !!(rail || context);
}

/**
 * @param {HTMLElement} nav
 * @param {any} node
 * @param {boolean} isActive
 * @param {boolean} inTools
 * @param {{ section: any, tool: any }} loc
 */
function appendGroup(nav, node, isActive, inTools, loc) {
    const block = document.createElement('div');
    block.className = 'portal-workspace-shell__block';
    if (isActive) block.classList.add('is-current');

    const row = document.createElement('div');
    row.className = 'portal-workspace-shell__row';

    const parent = document.createElement('a');
    parent.className = 'portal-workspace-shell__item';
    parent.href = indexHref(inTools);
    parent.title = 'Übersicht ' + (node.label || '');
    const icon = document.createElement('span');
    icon.className = 'portal-workspace-shell__icon';
    icon.setAttribute('aria-hidden', 'true');
    const ic = document.createElement('i');
    ic.className = safeIconClass(node.icon);
    icon.appendChild(ic);
    const label = document.createElement('span');
    label.className = 'portal-workspace-shell__label';
    label.textContent = node.label || node.id;
    parent.appendChild(icon);
    parent.appendChild(label);
    parent.addEventListener('click', function () {
        rememberPortalSection(node.id);
    });
    row.appendChild(parent);

    const tools = node.tools || [];
    /** @type {HTMLElement|null} */
    let sub = null;
    if (tools.length) {
        const chev = document.createElement('button');
        chev.type = 'button';
        chev.className = 'portal-workspace-shell__chevron-btn';
        chev.setAttribute('aria-expanded', isActive ? 'true' : 'false');
        chev.setAttribute(
            'aria-label',
            (node.label || 'Bereich') + (isActive ? ' zuklappen' : ' aufklappen')
        );
        const chevIcon = document.createElement('i');
        chevIcon.className = 'bi bi-chevron-down';
        chevIcon.setAttribute('aria-hidden', 'true');
        chev.appendChild(chevIcon);
        sub = document.createElement('div');
        sub.className = 'portal-workspace-shell__sub';
        sub.hidden = !isActive;
        chev.addEventListener('click', function (ev) {
            ev.preventDefault();
            const open = !!sub && sub.hidden;
            if (sub) sub.hidden = !open;
            chev.setAttribute('aria-expanded', open ? 'true' : 'false');
            chev.setAttribute(
                'aria-label',
                (node.label || 'Bereich') + (open ? ' zuklappen' : ' aufklappen')
            );
            block.classList.toggle('is-collapsed', !open);
        });
        row.appendChild(chev);
    }
    block.appendChild(row);

    if (sub) {
        tools.forEach(function (tool) {
            const href = safeToolHref(tool.href, inTools);
            if (!href) return;
            const a = document.createElement('a');
            a.className = 'portal-workspace-shell__subitem';
            a.href = href;
            if (isActive && loc.tool && tool.id === loc.tool.id) {
                a.setAttribute('aria-current', 'page');
            }
            const subIcon = document.createElement('span');
            subIcon.className = 'portal-workspace-shell__icon';
            subIcon.setAttribute('aria-hidden', 'true');
            const subIc = document.createElement('i');
            subIc.className = safeIconClass(tool.icon);
            subIcon.appendChild(subIc);
            const subLabel = document.createElement('span');
            subLabel.className = 'portal-workspace-shell__label';
            subLabel.textContent = tool.label || tool.id;
            a.appendChild(subIcon);
            a.appendChild(subLabel);
            a.addEventListener('click', function () {
                rememberPortalSection(node.id);
            });
            sub.appendChild(a);
        });
        block.appendChild(sub);
    }

    nav.appendChild(block);
}

function buildContextBar(loc, inTools) {
    const bar = document.createElement('nav');
    bar.id = 'portalWorkspaceContext';
    bar.className = 'portal-workspace-shell__context';
    bar.setAttribute('aria-label', 'Portal-Pfad');

    const back = document.createElement('a');
    back.className = 'portal-workspace-shell__back';
    back.href = indexHref(inTools);
    const sectionLabel = (loc.section && loc.section.label) || 'Portal';
    back.setAttribute('aria-label', 'Zurück zur Übersicht ' + sectionLabel);
    const backIcon = document.createElement('i');
    backIcon.className = 'bi bi-arrow-left';
    backIcon.setAttribute('aria-hidden', 'true');
    const backText = document.createElement('span');
    backText.textContent = sectionLabel;
    back.appendChild(backIcon);
    back.appendChild(backText);
    back.addEventListener('click', function () {
        rememberPortalSection(loc.section && loc.section.id);
    });

    const sep = document.createElement('span');
    sep.className = 'portal-workspace-shell__sep';
    sep.setAttribute('aria-hidden', 'true');
    sep.textContent = '/';

    const current = document.createElement('span');
    current.className = 'portal-workspace-shell__current';
    current.setAttribute('aria-current', 'page');
    current.textContent = (loc.tool && loc.tool.label) || '';

    bar.appendChild(back);
    bar.appendChild(sep);
    bar.appendChild(current);
    return bar;
}

/**
 * @returns {boolean}
 */
export function mountPortalWorkspaceShell() {
    if (typeof document === 'undefined' || !document.body) return false;
    if (!shouldMountPortalWorkspace()) {
        unmountPortalWorkspaceShell();
        return false;
    }

    const picked = pickTreeAndLocation(currentPageHref());
    if (!picked) {
        unmountPortalWorkspaceShell();
        return false;
    }

    const existing = document.getElementById('portalWorkspaceRail');
    if (existing && existing.getAttribute('data-portal-view') === picked.view) {
        hideDashboardBackLinks();
        return true;
    }
    if (existing) unmountPortalWorkspaceShell();

    ensureShellStylesheet();

    const inTools = /\/tools\//i.test(String(location.pathname || ''));
    const container =
        document.querySelector('body > .container.page-card') ||
        document.querySelector('body > .container') ||
        document.querySelector('.container.page-card') ||
        document.querySelector('.container');
    if (!container) return false;

    const rail = document.createElement('aside');
    rail.id = 'portalWorkspaceRail';
    rail.className = 'portal-workspace-shell__rail';
    rail.setAttribute('data-portal-view', picked.view);
    rail.setAttribute('aria-label', 'Portal-Menü');

    const brand = document.createElement('div');
    brand.className = 'portal-workspace-shell__brand';
    const home = document.createElement('a');
    home.className = 'portal-workspace-shell__home';
    home.href = indexHref(inTools);
    const homeIcon = document.createElement('i');
    homeIcon.className = 'bi bi-grid-1x2';
    homeIcon.setAttribute('aria-hidden', 'true');
    const homeLabel = document.createElement('span');
    homeLabel.textContent = 'Portal';
    home.appendChild(homeIcon);
    home.appendChild(homeLabel);
    home.addEventListener('click', function () {
        rememberPortalSection('home');
    });
    brand.appendChild(home);
    rail.appendChild(brand);

    const nav = document.createElement('nav');
    nav.className = 'portal-workspace-shell__nav';
    nav.setAttribute('aria-label', 'Portal-Bereiche');
    picked.tree.forEach(function (node) {
        const isActive = !!(picked.loc.section && node.id === picked.loc.section.id);
        appendGroup(nav, node, isActive, inTools, picked.loc);
    });
    rail.appendChild(nav);

    const context = buildContextBar(picked.loc, inTools);
    const header = container.querySelector(':scope > .header');
    if (header && header.parentElement === container) {
        const after = header.nextSibling;
        container.insertBefore(rail, after);
        container.insertBefore(context, rail.nextSibling);
    } else {
        container.insertBefore(rail, container.firstChild);
        container.insertBefore(context, rail.nextSibling);
    }

    document.body.classList.add('portal-workspace-shell', 'app-shell-chrome');
    hideDashboardBackLinks();
    return true;
}

function syncPortalWorkspaceShell() {
    if (!shouldMountPortalWorkspace()) {
        unmountPortalWorkspaceShell();
        return;
    }
    const picked = pickTreeAndLocation(currentPageHref());
    const existing = document.getElementById('portalWorkspaceRail');
    if (existing && picked && existing.getAttribute('data-portal-view') !== picked.view) {
        unmountPortalWorkspaceShell();
    }
    mountPortalWorkspaceShell();
}

function wirePortalWorkspaceLifecycle() {
    if (lifecycleWired || typeof window === 'undefined') return;
    lifecycleWired = true;
    window.addEventListener('ms365-dashboard-persona-ready', syncPortalWorkspaceShell);
    window.addEventListener('ms365-auth-state-changed', syncPortalWorkspaceShell);
    window.addEventListener('ms365-it-shell-changed', syncPortalWorkspaceShell);
}

export function bootPortalWorkspaceShell() {
    try {
        wirePortalWorkspaceLifecycle();
        return mountPortalWorkspaceShell();
    } catch {
        return false;
    }
}
