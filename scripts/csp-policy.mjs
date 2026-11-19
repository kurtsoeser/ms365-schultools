/**
 * Content-Security-Policy der App – eine Quelle für Dev-Server, Build und Prüfskript.
 *
 * Auslieferung als <meta http-equiv>, weil GitHub Pages keine eigenen HTTP-Header
 * zulässt. Das Meta-Tag steht als erstes Element im <head>, damit es für alle
 * Skripte der Seite gilt. (frame-ancestors und Report-Only gehen nur per Header.)
 *
 * Kern: script-src 'self' – keine Inline-Skripte, keine Inline-Handler, kein eval,
 * keine fremden Skript-Hosts. Neue Seiten müssen Skripte als Datei einbinden;
 * scripts/check-csp.mjs und tests/browser-trust-chain.test.mjs prüfen das.
 */
export const CSP_DIRECTIVES = {
    'default-src': ["'self'"],
    'script-src': ["'self'"],
    // Inline-Styles sind im Bestand überall (style="…", <style>); vergleichsweise geringes Risiko.
    'style-src': ["'self'", "'unsafe-inline'"],
    // https: wegen Schul-Logos per URL, Graph-/SharePoint-Vorschauen, QR-Codes.
    'img-src': ["'self'", 'data:', 'blob:', 'https:'],
    'font-src': ["'self'", 'data:'],
    // Bewusst breit (https:): Graph, Anmeldung, SharePoint und die Backend-URLs sind
    // je Schule konfigurierbar. Klartext-HTTP nur zum lokalen Nextcloud-Helfer.
    'connect-src': ["'self'", 'https:', 'data:', 'blob:', 'http://127.0.0.1:*', 'http://localhost:*'],
    // MSAL erneuert Tokens über ein verstecktes iframe zur Microsoft-Anmeldung.
    'frame-src': ["'self'", 'https://login.microsoftonline.com'],
    'worker-src': ["'self'"],
    'object-src': ["'none'"],
    'base-uri': ["'self'"],
    'form-action': ["'self'"]
};

/** Nur Vite-Dev-Server: Hot-Reload-WebSocket. */
const DEV_CONNECT_EXTRA = ['ws://localhost:*', 'ws://127.0.0.1:*'];

const CSP_META_RE = /\s*<meta\s+http-equiv=["']Content-Security-Policy["'][^>]*>/i;

/**
 * @param {{ dev?: boolean }} [opts]
 * @returns {string}
 */
export function buildCspContent(opts = {}) {
    return Object.entries(CSP_DIRECTIVES)
        .map(([name, values]) => {
            const list = name === 'connect-src' && opts.dev ? [...values, ...DEV_CONNECT_EXTRA] : values;
            return `${name} ${list.join(' ')}`;
        })
        .join('; ');
}

/** @param {{ dev?: boolean }} [opts] */
export function cspMetaTag(opts = {}) {
    return `<meta http-equiv="Content-Security-Policy" content="${buildCspContent(opts)}">`;
}

/**
 * Setzt das CSP-Meta-Tag direkt hinter <head> (ersetzt ein vorhandenes).
 * @param {string} html
 * @param {{ dev?: boolean }} [opts]
 */
export function injectCspMeta(html, opts = {}) {
    const withoutOld = html.replace(CSP_META_RE, '');
    const headOpen = withoutOld.match(/<head(\s[^>]*)?>/i);
    if (!headOpen) return html;
    const at = headOpen.index + headOpen[0].length;
    return `${withoutOld.slice(0, at)}\n    ${cspMetaTag(opts)}${withoutOld.slice(at)}`;
}

/** Vite-Plugin: Meta-Tag in jede HTML-Seite (Dev-Server und Build). */
export function cspMetaPlugin() {
    let dev = false;
    return {
        name: 'ms365-csp-meta',
        configResolved(config) {
            dev = config.command === 'serve';
        },
        transformIndexHtml: {
            order: 'post',
            handler(html) {
                return injectCspMeta(html, { dev });
            }
        }
    };
}
