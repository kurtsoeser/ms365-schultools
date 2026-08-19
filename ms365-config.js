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
 * Kursteams Azure-Backend – functionKey aus Azure Portal eintragen (App-Schlüssel → default).
 */
window.MS365_KURSTEAMS_API = {
    baseUrl: 'https://func-ms365-kursteams-dev-cmatbeawgqf8daaq.westeurope-01.azurewebsites.net/api/kursteams',
    functionKey: ''
};

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

    function injectSiteCredit() {
        let p = document.getElementById('siteCreditKurtrocks') || document.querySelector('.site-credit-row');
        if (!p) {
            p = document.createElement('p');
            p.className = 'site-credit-row';
            document.body.appendChild(p);
        }
        p.id = 'siteCreditKurtrocks';

        let a = p.querySelector('.site-credit-link');
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
