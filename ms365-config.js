/**
 * Tragen Sie unten Ihre Anwendungs-ID (Client) aus der Entra-App-Registrierung ein.
 * Ausführliche Schritte: siehe ms365-config.example.js (Kommentarblock oben).
 */
window.MS365_MSAL_CONFIG = {
    clientId: 'e1d877c3-004c-4040-8c3b-81a59e0c7050',
    authority: 'https://login.microsoftonline.com/organizations',
    redirectUri: (function () {
        if (typeof window === 'undefined') return '';
        try {
            const origin = window.location.origin;
            const host = (window.location.hostname || '').toLowerCase();
            const isLocal =
                host === 'localhost' ||
                host === '127.0.0.1' ||
                host === '::1' ||
                host.endsWith('.localhost');

            function basePathForThisHost() {
                // Ziel: bei GitHub Pages Project Pages (…/repo/…) automatisch den Repo-Pfad mitnehmen.
                // Beispiele:
                // - /ms365-schultools/tools/schulstruktur-sync.html  -> /ms365-schultools
                // - /ms365-schultools/index.html                   -> /ms365-schultools
                // - /tools/archiv/arge.html                        -> (root)
                const p = String(window.location.pathname || '/');
                const noQuery = p.split('?')[0].split('#')[0];
                // Wenn wir in /tools/… sind, ist alles davor die "Basis"
                const iTools = noQuery.toLowerCase().indexOf('/tools/');
                if (iTools !== -1) {
                    const base = noQuery.slice(0, iTools);
                    return base.endsWith('/') ? base.slice(0, -1) : base;
                }
                // Sonst: Ordner der aktuellen Datei; bei /index.html oder /ms365-schooltool.html ist das bereits die Basis
                const lastSlash = noQuery.lastIndexOf('/');
                if (lastSlash <= 0) return '';
                const base = noQuery.slice(0, lastSlash);
                return base.endsWith('/') ? base.slice(0, -1) : base;
            }

            // Auch auf localhost den Repo-/Projektpfad mitnehmen (z. B. vite preview --base /MS365schule/).
            const base = basePathForThisHost();
            // Immer stabile Redirect-Seite verwenden (keine Tool-Unterseite),
            // damit Entra nur 1 Redirect-URI pro Umgebung braucht.
            return origin + (base ? base : '') + '/ms365-schooltool.html';
        } catch {
            return window.location.href.split('#')[0];
        }
    })()
};

/**
 * Kursteams Azure-Backend – Basis-URL und API-Scope (öffentlich).
 * Anmeldung über das Benutzer-Token, kein Function Key im Browser.
 */
window.MS365_KURSTEAMS_API = {
    baseUrl: 'https://func-ms365-kursteams-dev-cmatbeawgqf8daaq.westeurope-01.azurewebsites.net/api/kursteams',
    scope: 'api://c7e6f467-e6f3-4221-a9ee-574b35120029/Kursteams.Create'
};

/**
 * License-API – baseUrl nach Deploy setzen (oder in ms365-config.local.js).
 * scope: delegierte Berechtigung der License-Backend-App (License.Access).
 * Endpunkt: GET {baseUrl}/me  mit Authorization: Bearer <Access-Token>
 */
window.MS365_LICENSE_API = {
    baseUrl: 'https://func-ms365-license-dev.azurewebsites.net/api/license',
    scope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/License.Access',
    functionKey: ''
};

(function loadMs365LocalConfig() {
    if (typeof XMLHttpRequest === 'undefined' || typeof document === 'undefined') return;
    try {
        const scripts = document.getElementsByTagName('script');
        let localUrl = '';
        for (let i = scripts.length - 1; i >= 0; i--) {
            const src = scripts[i].src || '';
            if (/ms365-config\.js(\?|$)/i.test(src)) {
                localUrl = src.replace(/ms365-config\.js(\?.*)?$/i, 'ms365-config.local.js$1');
                break;
            }
        }
        if (!localUrl) return;
        const xhr = new XMLHttpRequest();
        xhr.open('GET', localUrl, false);
        xhr.send(null);
        if (xhr.status !== 200 || !String(xhr.responseText || '').trim()) return;
        // Lokale Override-Datei (gitignored) – optional, nur auf Ihrer Maschine / im Deployment
        // eslint-disable-next-line no-new-func
        new Function(xhr.responseText)();
        const local = window.MS365_CONFIG_LOCAL;
        if (!local) return;
        if (window.MS365_KURSTEAMS_API) {
            const k = local.MS365_KURSTEAMS_API;
            if (k && k.baseUrl) {
                window.MS365_KURSTEAMS_API.baseUrl = String(k.baseUrl).trim();
            }
            if (k && k.scope) {
                window.MS365_KURSTEAMS_API.scope = String(k.scope).trim();
            }
        }
        if (!window.MS365_LICENSE_API) {
            window.MS365_LICENSE_API = {
                baseUrl: '',
                scope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/License.Access',
                functionKey: ''
            };
        }
        const lic = local.MS365_LICENSE_API;
        if (lic && lic.functionKey) {
            window.MS365_LICENSE_API.functionKey = String(lic.functionKey).trim();
        }
        if (lic && lic.baseUrl) {
            window.MS365_LICENSE_API.baseUrl = String(lic.baseUrl).trim();
        }
        if (lic && lic.scope) {
            window.MS365_LICENSE_API.scope = String(lic.scope).trim();
        }
    } catch {
        /* lokale Overrides optional */
    }
})();

(function () {
    if (typeof document === 'undefined') return;

    function resolveSharedScriptPath() {
        // Ziel: funktioniert in /tools/*.html, /tools/archiv/*.html und im Repo-Subpfad (GitHub Pages).
        try {
            const noQuery = String(window.location.pathname || '/').split('?')[0].split('#')[0];
            const lower = noQuery.toLowerCase();
            const idx = lower.indexOf('/tools/');
            if (idx === -1) return 'src/shared/msal-auth-ui.js';
            const afterTools = noQuery.slice(idx + '/tools/'.length);
            const depth = Math.max(0, afterTools.split('/').length - 1);
            return '../'.repeat(depth + 1) + 'src/shared/msal-auth-ui.js';
        } catch {
            return 'src/shared/msal-auth-ui.js';
        }
    }

    function ensureGlobalAuthUi() {
        // Auth-Widget auf allen Seiten einbinden (einmalig).
        try {
            if (document.getElementById('ms365GlobalAuthUiScript')) return;
            if (typeof window.ms365AuthAcquireToken === 'function') return;
            const already = document.querySelector('script[src*="msal-auth-ui"]');
            if (already) return;
            const s = document.createElement('script');
            s.id = 'ms365GlobalAuthUiScript';
            s.type = 'module';
            s.src = resolveSharedScriptPath();
            document.head.appendChild(s);
        } catch {
            // ignore
        }
    }

    function ensureFooterContainer() {
        let footer = document.getElementById('ms365FixedFooter');
        if (footer) return footer;
        footer = document.createElement('div');
        footer.id = 'ms365FixedFooter';
        footer.className = 'app-fixed-footer';

        const left = document.createElement('div');
        left.id = 'ms365FixedFooterLeft';
        left.className = 'app-fixed-footer__left';

        const right = document.createElement('div');
        right.id = 'ms365FixedFooterRight';
        right.className = 'app-fixed-footer__right';

        footer.appendChild(left);
        footer.appendChild(right);
        document.body.appendChild(footer);
        return footer;
    }

    function moveFooterItemsIntoFooter() {
        const footer = ensureFooterContainer();
        const left = footer.querySelector('#ms365FixedFooterLeft');
        const right = footer.querySelector('#ms365FixedFooterRight');
        if (!left || !right) return;

        const siteCredit = document.querySelector('.site-credit-row');
        const helpRow = document.querySelector('.header-help-row');
        const stamp = document.getElementById('ms365AppPublishedStamp');

        if (siteCredit && siteCredit.parentElement !== left) left.appendChild(siteCredit);
        if (stamp && stamp.parentElement !== right) right.appendChild(stamp);
        if (helpRow && helpRow.parentElement !== right) right.appendChild(helpRow);
    }

    function landingPageHref() {
        try {
            const pathOnly = String(window.location.pathname || '/')
                .split('?')[0]
                .split('#')[0];
            const parts = pathOnly.split('/').filter(Boolean);
            if (parts.length && /\.html?$/i.test(parts[parts.length - 1])) {
                parts.pop();
            }
            const lower = parts.map((p) => p.toLowerCase());
            const toolsIdx = lower.indexOf('tools');
            if (toolsIdx === -1) return 'landing/';
            const ups = parts.length - toolsIdx;
            return '../'.repeat(Math.max(1, ups)) + 'landing/';
        } catch {
            return 'landing/';
        }
    }

    function injectSiteCredit() {
        let p = document.getElementById('siteCreditKurtrocks') || document.querySelector('.site-credit-row');
        if (!p) {
            p = document.createElement('p');
            p.className = 'site-credit-row';
            document.body.appendChild(p);
        }
        p.id = 'siteCreditKurtrocks';

        let landing = p.querySelector('.site-landing-link');
        if (!landing) {
            landing = document.createElement('a');
            landing.className = 'site-credit-link site-landing-link';
            const icon = document.createElement('i');
            icon.className = 'bi bi-globe2';
            icon.setAttribute('aria-hidden', 'true');
            landing.appendChild(icon);
            landing.appendChild(document.createTextNode('Website'));
            const kur = p.querySelector('.site-credit-link:not(.site-landing-link)');
            if (kur) p.insertBefore(landing, kur);
            else p.appendChild(landing);
        }
        landing.href = landingPageHref();
        landing.title = 'Marketing-Website / Landing Page';
        landing.setAttribute('aria-label', 'Website – Landing Page – MS365-Schul-Tools');

        let impressum = p.querySelector('.site-impressum-link');
        if (!impressum) {
            impressum = document.createElement('a');
            impressum.className = 'site-credit-link site-impressum-link';
            const iconImp = document.createElement('i');
            iconImp.className = 'bi bi-file-earmark-text';
            iconImp.setAttribute('aria-hidden', 'true');
            impressum.appendChild(iconImp);
            impressum.appendChild(document.createTextNode('Impressum'));
            const afterLanding = p.querySelector('.site-landing-link');
            if (afterLanding && afterLanding.nextSibling) {
                p.insertBefore(impressum, afterLanding.nextSibling);
            } else if (afterLanding) {
                afterLanding.after(impressum);
            } else {
                p.appendChild(impressum);
            }
        }
        impressum.href = landingPageHref().replace(/\/?$/, '/') + 'impressum.html';
        impressum.title = 'Impressum';
        impressum.setAttribute('aria-label', 'Impressum');

        let privacy = p.querySelector('.site-privacy-link');
        if (!privacy) {
            privacy = document.createElement('a');
            privacy.className = 'site-credit-link site-privacy-link';
            const iconPriv = document.createElement('i');
            iconPriv.className = 'bi bi-shield-lock';
            iconPriv.setAttribute('aria-hidden', 'true');
            privacy.appendChild(iconPriv);
            privacy.appendChild(document.createTextNode('Datenschutz'));
            if (impressum.nextSibling) p.insertBefore(privacy, impressum.nextSibling);
            else impressum.after(privacy);
        }
        privacy.href = landingPageHref().replace(/\/?$/, '/') + 'datenschutz.html';
        privacy.title = 'Datenschutzerklärung';
        privacy.setAttribute('aria-label', 'Datenschutzerklärung');

        let a = p.querySelector('.site-credit-link:not(.site-landing-link):not(.site-impressum-link):not(.site-privacy-link)');
        if (!a) {
            a = document.createElement('a');
            a.className = 'site-credit-link';
            a.href = 'https://www.kurtrocks.com/';
            a.target = '_blank';
            a.rel = 'noopener noreferrer';
            const icon = document.createElement('i');
            icon.className = 'bi bi-info-circle';
            icon.setAttribute('aria-hidden', 'true');
            a.appendChild(icon);
            a.appendChild(document.createTextNode('kurtrocks.com'));
            p.appendChild(a);
        }
        a.title = 'Ein Projekt von Kurt Söser';
        a.setAttribute('aria-label', 'kurtrocks.com - Ein Projekt von Kurt Söser');

        moveFooterItemsIntoFooter();
        try {
            if (window.ms365Theme && typeof window.ms365Theme.mount === 'function') {
                window.ms365Theme.mount();
            }
        } catch {
            /* ignore */
        }
    }
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            ensureGlobalAuthUi();
            injectSiteCredit();
            moveFooterItemsIntoFooter();
        });
    } else {
        ensureGlobalAuthUi();
        injectSiteCredit();
        moveFooterItemsIntoFooter();
    }
})();
