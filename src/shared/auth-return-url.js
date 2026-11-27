/**
 * Wohin ms365-schooltool.html nach der Microsoft-Anmeldung zurückgeht.
 *
 * Besitzer/Mitglieder schreiben über die IT-App. Dabei kann der Browser kurz
 * auf die Redirect-Seite wechseln. Die gemerkte Rückkehr-URL darf dann nicht
 * ein früher geöffnetes Werkzeug sein (z. B. Nextcloud → Teams).
 */

export const POST_LOGIN_KEY = 'ms365-post-login-url';

/**
 * @param {string} href
 * @returns {boolean}
 */
export function isSchooltoolHref(href) {
    try {
        const u = new URL(String(href || ''), 'https://ms365.schule');
        return /\/ms365-schooltool\.html$/i.test(u.pathname);
    } catch {
        return false;
    }
}

/**
 * @param {string} href
 * @returns {boolean}
 */
export function pageHasAuthResponse(href) {
    return /[#&?](code|error|id_token|access_token)=/i.test(String(href || ''));
}

/**
 * Gleiche Origin, keine Redirect-Seite. Sonst leer.
 * @param {string} href
 * @param {string} pageHref
 * @returns {string}
 */
export function sameOriginAppHref(href, pageHref) {
    try {
        const page = new URL(String(pageHref || ''), 'https://ms365.schule');
        const u = new URL(String(href || ''), page);
        if (u.origin !== page.origin) return '';
        if (!/^https?:$/i.test(u.protocol)) return '';
        if (/\/ms365-schooltool\.html$/i.test(u.pathname)) return '';
        return u.href;
    } catch {
        return '';
    }
}

/**
 * @param {string} href
 * @param {string} baseHref
 * @returns {string}
 */
function pathnameOf(href, baseHref) {
    try {
        return new URL(String(href || ''), String(baseHref || 'https://ms365.schule')).pathname;
    } catch {
        return '';
    }
}

/**
 * Seite, die diese Anmeldung ausgelöst hat, schlägt eine ältere Session-URL.
 * @param {{ pageHref?: string, storedHref?: string, requestOrigins?: string[] }} input
 * @returns {string}
 */
export function pickAuthReturnUrl(input) {
    const pageHref = String((input && input.pageHref) || '');
    const origins = Array.isArray(input && input.requestOrigins) ? input.requestOrigins : [];
    for (let i = 0; i < origins.length; i++) {
        const hit = sameOriginAppHref(origins[i], pageHref);
        if (hit) return hit;
    }
    return sameOriginAppHref(input && input.storedHref, pageHref);
}

/**
 * @param {Storage} storage
 * @returns {string[]}
 */
export function readRequestOrigins(storage) {
    const out = [];
    if (!storage) return out;
    try {
        for (let i = 0; i < storage.length; i++) {
            const key = storage.key(i);
            if (!key || !/\.request\.origin$/i.test(key)) continue;
            const val = storage.getItem(key);
            if (val) out.push(String(val));
        }
    } catch {
        /* ignore */
    }
    return out;
}

/**
 * Merkt die aktuelle Werkzeugseite als Rückkehrziel und überschreibt eine
 * MSAL-Herkunft, die noch auf ein anderes Werkzeug zeigt.
 * @param {Window & typeof globalThis} [win]
 * @returns {boolean}
 */
export function pinAuthReturnToThisPage(win) {
    const w = win || (typeof window !== 'undefined' ? window : null);
    if (!w || !w.location || !w.sessionStorage) return false;
    const href = String(w.location.href || '');
    if (!href || isSchooltoolHref(href) || pageHasAuthResponse(href)) return false;
    const storage = w.sessionStorage;
    try {
        storage.setItem(POST_LOGIN_KEY, href);
    } catch {
        return false;
    }
    let rewrote = false;
    try {
        const keys = [];
        for (let i = 0; i < storage.length; i++) {
            const key = storage.key(i);
            if (key) keys.push(key);
        }
        const herePath = pathnameOf(href, href);
        keys.forEach(function (key) {
            if (!/\.request\.origin$/i.test(key)) return;
            const prev = storage.getItem(key) || '';
            const prevPath = pathnameOf(prev, href);
            if (!prevPath || !herePath || prevPath === herePath) return;
            if (!sameOriginAppHref(prev, href)) return;
            storage.setItem(key, href);
            rewrote = true;
        });
        if (rewrote) {
            keys.forEach(function (key) {
                if (/\.interaction\.status$/i.test(key)) storage.removeItem(key);
            });
        }
    } catch {
        /* Herkunft ist nur ein Zusatz zur Session-URL */
    }
    return true;
}
