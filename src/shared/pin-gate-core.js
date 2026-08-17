/**
 * PIN-Zugang (Session) – reine Logik für Guard und Willkommensseite.
 */

export const SESSION_KEY = 'ms365-access-granted-v1';

/** @param {Storage} [storage] */
export function isAccessGranted(storage = sessionStorage) {
    return storage.getItem(SESSION_KEY) === '1';
}

/** @param {Storage} [storage] */
export function grantAccess(storage = sessionStorage) {
    storage.setItem(SESSION_KEY, '1');
}

/** @param {Storage} [storage] */
export function revokeAccess(storage = sessionStorage) {
    storage.removeItem(SESSION_KEY);
}

/** @param {unknown} pin */
export function normalizePin(pin) {
    return String(pin == null ? '' : pin).trim();
}

/**
 * @param {unknown} pin
 * @param {string[]} pins
 */
export function isValidPin(pin, pins) {
    const value = normalizePin(pin);
    if (!value || !Array.isArray(pins) || !pins.length) return false;
    const needle = value.toLowerCase();
    return pins.some(function (entry) {
        return normalizePin(entry).toLowerCase() === needle;
    });
}

/**
 * @param {{ enabled?: boolean, pins?: string[] }} [config]
 */
export function isPinGateEnabled(config) {
    if (!config || config.enabled === false) return false;
    return Array.isArray(config.pins) && config.pins.length > 0;
}

/** @param {string} pathname */
export function isWelcomePath(pathname) {
    return /\/welcome\.html$/i.test(pathname || '');
}

/** Hilfe und Datenschutz ohne PIN-Sperre. */
export function isHelpPath(pathname) {
    return /\/hilfe\.html$/i.test(pathname || '');
}

/** Seiten, die die PIN-Sperre nicht auslösen. */
export function isPinExemptPath(pathname) {
    return isWelcomePath(pathname) || isHelpPath(pathname);
}

/**
 * @param {string} scriptSrc URL von pin-gate.js (…/src/shared/pin-gate.js)
 */
export function resolveWelcomeUrl(scriptSrc) {
    if (!scriptSrc) return 'welcome.html';
    try {
        return new URL('../../welcome.html', scriptSrc).href;
    } catch {
        return 'welcome.html';
    }
}

/**
 * Verzeichnis der App (dort liegt welcome.html / index.html).
 * @param {string} [welcomeHref]
 */
export function appBaseHref(welcomeHref) {
    const href =
        welcomeHref ||
        (typeof location !== 'undefined' ? location.href : 'https://example.invalid/welcome.html');
    return new URL('./', href).href;
}

/**
 * Liefert eine gleiche-Origin-URL innerhalb der App (kein doppelter Unterordner).
 * @param {string | null | undefined} raw Query `return` (Pfad, relativ oder absolut)
 * @param {string} [welcomeHref] z. B. location.href der Willkommensseite
 */
export function resolveReturnUrl(raw, welcomeHref) {
    const welcome =
        welcomeHref ||
        (typeof location !== 'undefined' ? location.href : 'https://example.invalid/welcome.html');
    const fallback = new URL('index.html', welcome).href;
    const base = new URL('./', welcome);

    if (!raw) return fallback;

    let decoded = raw;
    try {
        decoded = decodeURIComponent(raw);
    } catch {
        return fallback;
    }
    if (!decoded || /welcome\.html/i.test(decoded)) return fallback;
    if (decoded.startsWith('//')) return fallback;

    let candidate;
    try {
        candidate = new URL(decoded, base);
    } catch {
        return fallback;
    }

    if (candidate.origin !== base.origin) return fallback;
    if (/\/welcome\.html$/i.test(candidate.pathname)) return fallback;

    const basePath = base.pathname.endsWith('/') ? base.pathname : base.pathname + '/';
    const path = candidate.pathname;
    const inApp = path === basePath.slice(0, -1) || path === basePath || path.startsWith(basePath);
    if (!inApp) return fallback;

    return candidate.href;
}

/**
 * @param {string | null | undefined} raw
 * @param {string} [fallback]
 * @param {string} [welcomeHref]
 */
export function safeReturnPath(raw, fallback = 'index.html', welcomeHref) {
    const welcome =
        welcomeHref ||
        (typeof location !== 'undefined' && /welcome\.html/i.test(location.pathname)
            ? location.href
            : new URL(fallback, 'https://example.invalid/').href.replace(/[^/]+$/, 'welcome.html'));
    const resolved = resolveReturnUrl(raw, welcome);
    try {
        const u = new URL(resolved);
        const base = new URL('./', welcome);
        let rel = u.pathname;
        if (rel.startsWith(base.pathname)) rel = rel.slice(base.pathname.length);
        rel = rel.replace(/^\//, '');
        return (rel || fallback) + u.search + u.hash;
    } catch {
        return fallback;
    }
}
