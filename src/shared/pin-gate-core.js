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
 * @param {string | null | undefined} raw
 * @param {string} [fallback]
 */
export function safeReturnPath(raw, fallback = 'index.html') {
    const fb = fallback || 'index.html';
    if (!raw) return fb;
    let decoded = raw;
    try {
        decoded = decodeURIComponent(raw);
    } catch {
        return fb;
    }
    if (!decoded || decoded.includes('welcome.html')) return fb;
    if (/^https?:\/\//i.test(decoded)) {
        try {
            const u = new URL(decoded);
            if (typeof location !== 'undefined' && u.origin !== location.origin) return fb;
            const path = u.pathname + u.search + u.hash;
            return path.replace(/^\//, '') || fb;
        } catch {
            return fb;
        }
    }
    if (decoded.startsWith('//')) return fb;
    return decoded.replace(/^\//, '') || fb;
}
