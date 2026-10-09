/**
 * Einheitlicher Werkzeug-Subheader (unter dem App-Sticky-Header).
 *
 * Layout:
 * 1) app-tool-chrome__intro – Icon, Titel, Kurztext, Hilfe (zentriert)
 * 2) app-tool-chrome__nav – Toolbar; „← Dashboard“ immer als erster Button
 */
import { markToolSubheader } from './app-header-chrome.js';

/**
 * @param {HTMLAnchorElement} a
 */
export function isDashboardBackLink(a) {
    if (!a || !a.getAttribute) return false;
    const href = String(a.getAttribute('href') || '').replace(/\\/g, '/');
    if (!/(?:^|\/)index\.html(?:\?|#|$)/i.test(href)) return false;
    const label = String(a.textContent || '')
        .replace(/\s+/g, ' ')
        .trim();
    return /dashboard/i.test(label);
}

/**
 * @param {ParentNode} root
 * @returns {HTMLAnchorElement|null}
 */
function _findDashboardLink(root) {
    if (!root) return null;
    const nav = document.getElementById('ms365HeaderNav');
    if (nav) {
        const links = nav.querySelectorAll('a[href]');
        for (let i = 0; i < links.length; i++) {
            if (isDashboardBackLink(links[i])) return links[i];
        }
    }
    const wayfind = root.querySelector('.app-tool-chrome__wayfind');
    if (wayfind) {
        const w = wayfind.querySelector('a.app-tool-chrome__back, a[href]');
        if (w && isDashboardBackLink(w)) return w;
    }
    const toolbar = root.querySelector('.app-tool-chrome__nav, .toolbar');
    if (toolbar) {
        const links = toolbar.querySelectorAll('a[href]');
        for (let j = 0; j < links.length; j++) {
            if (isDashboardBackLink(links[j])) return links[j];
        }
    }
    return null;
}

/**
 * @param {HTMLAnchorElement} link
 */
function upgradeDashboardBackLink(link) {
    link.classList.remove('btn');
    link.classList.add('app-tool-chrome__back', 'app-tool-chrome__nav-link');
    if (link.querySelector('.app-tool-chrome__back-label')) return;
    const raw = String(link.textContent || '')
        .replace(/\s+/g, ' ')
        .trim();
    const label = /dashboard/i.test(raw) ? 'Dashboard' : raw || 'Dashboard';
    link.textContent = '';
    const ic = document.createElement('i');
    ic.className = 'bi bi-arrow-left';
    ic.setAttribute('aria-hidden', 'true');
    const span = document.createElement('span');
    span.className = 'app-tool-chrome__back-label';
    span.textContent = label;
    link.appendChild(ic);
    link.appendChild(span);
}

/**
 * @param {HTMLElement} sub
 */
function ensureNavBar(sub) {
    let nav = sub.querySelector('.app-tool-chrome__nav');
    const toolbar = sub.querySelector('.toolbar');
    if (toolbar) {
        if (!nav) {
            toolbar.classList.add('app-tool-chrome__nav');
            nav = toolbar;
        } else if (toolbar !== nav) {
            while (toolbar.firstChild) nav.appendChild(toolbar.firstChild);
            toolbar.remove();
        }
    }
    if (!nav) {
        nav = document.createElement('nav');
        nav.className = 'app-tool-chrome__nav toolbar';
        nav.setAttribute('aria-label', 'Werkzeug-Navigation');
        sub.appendChild(nav);
    }
    return nav;
}

/**
 * @param {HTMLElement} nav
 */
function flattenNavBar(nav) {
    if (!nav) return;
    const links = Array.from(nav.querySelectorAll('a[href]'));
    if (!links.length) return;

    let dash = null;
    const rest = [];
    links.forEach(function (a) {
        if (isDashboardBackLink(a)) dash = a;
        else rest.push(a);
    });

    const frag = document.createDocumentFragment();
    if (dash) {
        upgradeDashboardBackLink(dash);
        frag.appendChild(dash);
    }
    rest.forEach(function (a) {
        a.classList.add('app-tool-chrome__nav-link');
        frag.appendChild(a);
    });

    nav.textContent = '';
    nav.appendChild(frag);
}

/**
 * @param {HTMLElement} sub
 */
function placeDashboardInNav(sub) {
    const nav = ensureNavBar(sub);

    const wayfind = sub.querySelector('.app-tool-chrome__wayfind');
    if (wayfind) wayfind.remove();

    flattenNavBar(nav);

    const hasItems = nav.querySelector('a[href], button, input, select, textarea');
    nav.hidden = !hasItems;
}

/**
 * @param {HTMLElement} header
 */
export function normalizeToolPageChrome(header) {
    if (typeof document === 'undefined') return false;
    markToolSubheader();

    const sub =
        header ||
        document.querySelector('.container.page-card > .header.app-tool-subheader') ||
        document.querySelector('.container > .header.app-tool-subheader');
    if (!sub || sub.classList.contains('dash-compact-header')) return false;
    if (!document.getElementById('dashCompactHeader')) return false;

    sub.dataset.appToolChrome = '1';
    sub.classList.add('app-tool-chrome');

    const legacyNav = document.getElementById('ms365HeaderNav');
    if (legacyNav) {
        legacyNav.querySelectorAll('a[href]').forEach(function (a) {
            if (isDashboardBackLink(a)) return;
            ensureNavBar(sub).appendChild(a);
        });
        legacyNav.remove();
    }

    let intro = sub.querySelector('.app-tool-chrome__intro');
    if (!intro) {
        intro = document.createElement('div');
        intro.className = 'app-tool-chrome__intro';
        const navExisting = sub.querySelector('.app-tool-chrome__nav, .toolbar');
        if (navExisting) sub.insertBefore(intro, navExisting);
        else sub.appendChild(intro);
    }

    sub.querySelectorAll('.app-tool-chrome__wayfind').forEach(function (w) {
        w.remove();
    });

    const indicator = sub.querySelector('.header-tool-indicator');
    if (indicator && indicator.parentElement !== intro) intro.appendChild(indicator);

    const children = Array.from(sub.children);
    children.forEach(function (node) {
        if (!(node instanceof HTMLElement)) return;
        if (
            node === intro ||
            node.classList.contains('app-tool-chrome__nav') ||
            node.classList.contains('toolbar') ||
            node.tagName === 'H1'
        ) {
            return;
        }
        if (node.matches('p.header-help-row, p:not(.header-help-row)')) {
            if (node.parentElement !== intro) intro.appendChild(node);
        }
    });

    placeDashboardInNav(sub);

    const h1 = sub.querySelector(':scope > h1');
    if (h1) h1.hidden = true;

    return true;
}

export function bootToolPageChrome() {
    normalizeToolPageChrome(null);
}

if (typeof window !== 'undefined') {
    window.ms365NormalizeToolPageChrome = bootToolPageChrome;
}
