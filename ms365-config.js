/**
 * Tragen Sie unten Ihre Anwendungs-ID (Client) aus der Entra-App-Registrierung ein.
 * Ausführliche Schritte: siehe ms365-config.example.js (Kommentarblock oben).
 */
/** Freistellungs-Planer: Team-Site-Pfade (ohne IT-Bibliothek) – Mandant kurtrocks / Demo */
window.MS365_FREISTELLUNG_PLANER = {
    sitePaths: ['/sites/MS365-Schultools']
};

window.MS365_MSAL_CONFIG = {
    clientId: 'e1d877c3-004c-4040-8c3b-81a59e0c7050',
    /**
     * Zweite App „ms365-schultools-it“: nur Schreib- und hochprivilegierte Aufrufe.
     * Öffentliche Client-ID, mehrere Mandanten, keine Zuweisung erforderlich.
     * Überschreibbar über meta name="ms365-graph-it-client-id", wenn dieses Feld leer ist.
     */
    itClientId: 'a1f27f94-78c2-4503-9603-192c86924842',
    authority: 'https://login.microsoftonline.com/organizations',
    redirectUri: (function () {
        if (typeof window === 'undefined') return '';
        try {
            const origin = window.location.origin;
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
 * Kursteams Azure-Backend – Scope öffentlich; baseUrl nur über Deploy-Secret
 * oder ms365-config.local.js (kein -dev in Produktion, a9).
 * Anmeldung über Benutzer-Token, kein Function Key im Browser.
 */
window.MS365_KURSTEAMS_API = {
    baseUrl: '',
    scope: 'api://c7e6f467-e6f3-4221-a9ee-574b35120029/Kursteams.Create'
};

/**
 * License-API – baseUrl per Deploy (LICENSE_API_BASE_URL) oder lokal.
 * scope: License.Access. Endpunkt: GET {baseUrl}/me mit Bearer-Token.
 * Kein Function Key in dist/ (a9).
 */
window.MS365_LICENSE_API = {
    baseUrl: '',
    scope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/License.Access',
    functionKey: '',
    /** true = kein /admin/me (nur Schulbetrieb, kein Betreiber-Menü) */
    skipOperatorCheck: false
};

/**
 * Schulsprecherwahl-API – Scope öffentlich; baseUrl per Deploy/Lokal.
 */
window.MS365_ELECTION_API = {
    baseUrl: '',
    scope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/Election.Access',
    fallbackScope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/License.Access',
    functionKey: ''
};

(function loadMs365LocalConfig() {
    if (typeof document === 'undefined') return;

    /**
     * Übernimmt window.MS365_CONFIG_LOCAL (aus ms365-config.local.js) in die Konfiguration oben.
     */
    function applyLocalConfig() {
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
        if (lic && lic.skipOperatorCheck === true) {
            window.MS365_LICENSE_API.skipOperatorCheck = true;
        }
        if (!window.MS365_ELECTION_API) {
            window.MS365_ELECTION_API = {
                baseUrl: '',
                scope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/Election.Access',
                fallbackScope: 'api://12e0cfe2-8337-4b35-93e8-542faf658eb3/License.Access',
                functionKey: ''
            };
        }
        const elec = local.MS365_ELECTION_API;
        if (elec && elec.baseUrl) {
            window.MS365_ELECTION_API.baseUrl = String(elec.baseUrl).trim();
        }
        if (elec && elec.scope) {
            window.MS365_ELECTION_API.scope = String(elec.scope).trim();
        }
        if (elec && elec.fallbackScope) {
            window.MS365_ELECTION_API.fallbackScope = String(elec.fallbackScope).trim();
        }
        if (elec && elec.functionKey) {
            window.MS365_ELECTION_API.functionKey = String(elec.functionKey).trim();
        }
    }

    function applyLocalConfigSafely() {
        try {
            applyLocalConfig();
        } catch {
            /* lokale Overrides optional */
        }
    }

    function findOwnScript() {
        if (document.currentScript) return document.currentScript;
        const scripts = document.getElementsByTagName('script');
        for (let i = scripts.length - 1; i >= 0; i--) {
            if (/ms365-config\.js(\?|$)/i.test(scripts[i].src || '')) return scripts[i];
        }
        return null;
    }

    try {
        // Lokale Override-Datei (gitignored) – optional, nur auf Ihrer Maschine / im Deployment.
        // Sie wird als normales Skript derselben Origin geladen. Früher: XHR + new Function –
        // das blockiert eine Content-Security-Policy ohne 'unsafe-eval'.
        const self = findOwnScript();
        const src = (self && self.src) || '';
        if (!/ms365-config\.js(\?|$)/i.test(src)) return;
        const localUrl = src.replace(/ms365-config\.js(\?.*)?$/i, 'ms365-config.local.js$1');
        const marker = 'data-ms365-config-local';
        const parserBlocking =
            document.readyState === 'loading' && !self.async && !self.defer && self.type !== 'module';
        if (parserBlocking) {
            // Wie bisher synchron: die Datei läuft direkt nach diesem Skript und vor allen folgenden.
            const attrUrl = localUrl.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
            document.write('<script src="' + attrUrl + '" ' + marker + '="1"><\/script>');
            const written = document.querySelector('script[' + marker + ']');
            if (written) written.addEventListener('load', applyLocalConfigSafely);
        } else {
            const s = document.createElement('script');
            s.src = localUrl;
            s.async = false;
            s.setAttribute(marker, '1');
            s.addEventListener('load', applyLocalConfigSafely);
            document.head.appendChild(s);
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

    const KURTROCKS_HREF = 'https://www.kurtrocks.education/';

    function moveFooterItemsIntoFooter() {
        const footer = ensureFooterContainer();
        const right = footer.querySelector('#ms365FixedFooterRight');
        if (!right) return;

        const help = document.getElementById('siteCreditKurtrocks');
        if (help && help.parentElement !== right) right.appendChild(help);

        const menu = document.getElementById('appFooterHelpMenu');
        const stamp = document.getElementById('ms365AppPublishedStamp');
        if (stamp && menu && stamp.parentElement !== menu) menu.appendChild(stamp);
    }

    function stripLooseThemeToggles() {
        document.querySelectorAll('[data-ms365-theme-toggle]').forEach(function (el) {
            if (el.closest('#ms365AuthDropdown, .ms365-auth-menu')) return;
            el.remove();
        });
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

    function ensureFooterHelpStyles() {
        if (document.getElementById('ms365FooterHelpStyles')) return;
        const style = document.createElement('style');
        style.id = 'ms365FooterHelpStyles';
        style.textContent =
            '.app-fixed-footer{position:fixed;left:0;bottom:0;width:0;height:0;margin:0;padding:0;overflow:visible;pointer-events:none;z-index:80;}' +
            '.app-footer-help{position:fixed;right:12px;bottom:12px;z-index:90;pointer-events:auto;}' +
            'body.ts-admin-page .app-footer-help{right:14px;bottom:14px;}' +
            'html[data-ms365-it-embed] #ms365FixedFooter,html[data-ms365-it-embed] #siteCreditKurtrocks,html[data-ms365-it-embed] .app-footer-help{display:none!important;}' +
            '.app-footer-help__btn{display:inline-flex;align-items:center;justify-content:center;width:46px;height:46px;margin:0;padding:0;border-radius:12px;border:1px solid var(--border,rgba(50,50,93,.16));background:var(--card,#fff);color:var(--heading,#1e2440);box-shadow:0 8px 24px rgba(12,18,34,.1);cursor:pointer;font:inherit;line-height:1;}' +
            '.app-footer-help__btn .bi{margin:0;font-size:1.2rem;color:inherit;}' +
            '.app-footer-help__btn:hover,.app-footer-help.is-open .app-footer-help__btn{background:var(--soft,#f4f6fb);}' +
            '.app-footer-help__menu{position:absolute;right:0;bottom:calc(100% + 8px);min-width:240px;max-width:min(320px,calc(100vw - 24px));padding:6px;border-radius:10px;background:var(--card,#fff);color:var(--heading,#1e2440);border:1px solid var(--border,rgba(50,50,93,.12));box-shadow:0 16px 40px rgba(0,0,0,.18);display:flex;flex-direction:column;gap:1px;text-align:left;}' +
            '.app-footer-help__menu[hidden]{display:none !important;}' +
            '.app-footer-help__menu .app-footer-help__item,.app-footer-help__menu .header-help-link,.app-footer-help__menu .site-credit-link,.app-footer-help__menu .site-feedback-link{display:flex;align-items:center;gap:8px;width:100%;margin:0;padding:8px 10px;border:0;border-radius:8px;background:transparent;box-shadow:none;color:var(--heading,#1e2440);text-decoration:none;font-weight:600;font-size:14px;line-height:1.3;cursor:pointer;opacity:1;}' +
            '.app-footer-help__menu .header-help-link__label{display:inline;}' +
            '.app-footer-help__menu .app-footer-help__item:hover,.app-footer-help__menu .header-help-link:hover,.app-footer-help__menu .site-credit-link:hover,.app-footer-help__menu .site-feedback-link:hover{background:var(--soft,#f4f6fb);}' +
            '.app-footer-help__menu .bi{margin:0;font-size:1.05rem;color:var(--heading,#1e2440);opacity:.8;}' +
            '.app-footer-help__sep{height:1px;margin:4px 6px;background:var(--border,rgba(50,50,93,.12));}' +
            '.app-footer-help__menu .app-published-stamp{margin:4px 10px 6px;padding:0;text-align:left;color:var(--muted,#5c6578);opacity:.9;font-size:.75rem;font-weight:500;white-space:normal;}' +
            '.app-fixed-footer [data-ms365-theme-toggle],.app-footer-help [data-ms365-theme-toggle],.site-credit-row [data-ms365-theme-toggle]{display:none !important;}' +
            'html[data-theme="dark"] .app-footer-help__btn{background:var(--card,#1c1f33);color:var(--soft,#e8eaf4);border-color:rgba(255,255,255,.18);}' +
            'html[data-theme="dark"] .app-footer-help__menu .app-footer-help__item,html[data-theme="dark"] .app-footer-help__menu .header-help-link,html[data-theme="dark"] .app-footer-help__menu .site-credit-link,html[data-theme="dark"] .app-footer-help__menu .site-feedback-link{background:transparent;color:var(--soft,#e8eaf4);border:0;box-shadow:none;}' +
            'html[data-theme="dark"] .app-footer-help__menu .app-footer-help__item:hover,html[data-theme="dark"] .app-footer-help__menu .header-help-link:hover,html[data-theme="dark"] .app-footer-help__menu .site-credit-link:hover,html[data-theme="dark"] .app-footer-help__menu .site-feedback-link:hover{background:var(--surface-muted,#2a2e45);}' +
            '@media (max-width:640px){.app-footer-help{right:10px;bottom:10px;}}';
        document.head.appendChild(style);
    }

    function makeFooterHelpItem(className, icon, label) {
        const a = document.createElement('a');
        a.className = 'app-footer-help__item ' + className;
        a.setAttribute('role', 'menuitem');
        const i = document.createElement('i');
        i.className = 'bi ' + icon;
        i.setAttribute('aria-hidden', 'true');
        const span = document.createElement('span');
        span.textContent = label;
        a.appendChild(i);
        a.appendChild(span);
        return a;
    }

    function ensureFooterHelpItem(menu, selector, className, icon, label) {
        let el = menu.querySelector(selector);
        if (!el) {
            el = makeFooterHelpItem(className, icon, label);
            const sep = menu.querySelector('.app-footer-help__sep');
            if (sep) menu.insertBefore(el, sep);
            else menu.appendChild(el);
        } else {
            el.classList.add('app-footer-help__item');
            el.setAttribute('role', 'menuitem');
        }
        return el;
    }

    var SAVE_DROP_OPEN =
        'details.ts-register-save-drop[open], details.tool-state-bar__drop[open]';
    var SAVE_DROP_ITEM =
        'details.ts-register-save-drop .ts-register-save-link, details.tool-state-bar__drop .tool-state-bar__link';

    function closeSaveDropMenus(except) {
        document.querySelectorAll(SAVE_DROP_OPEN).forEach(function (el) {
            if (except && el === except) return;
            el.open = false;
        });
    }

    function bindSaveDropClickAway() {
        if (document.documentElement.dataset.saveDropAwayBound === '1') return;
        document.documentElement.dataset.saveDropAwayBound = '1';
        document.addEventListener('click', function (ev) {
            const target = ev.target;
            document.querySelectorAll(SAVE_DROP_OPEN).forEach(function (el) {
                if (target instanceof Node && el.contains(target)) return;
                el.open = false;
            });
        });
        document.addEventListener('keydown', function (ev) {
            if (ev.key !== 'Escape') return;
            closeSaveDropMenus();
        });
        document.addEventListener(
            'toggle',
            function (ev) {
                const el = ev.target;
                if (!(el instanceof HTMLDetailsElement) || !el.open) return;
                if (!el.classList.contains('ts-register-save-drop') && !el.classList.contains('tool-state-bar__drop')) {
                    return;
                }
                closeSaveDropMenus(el);
                if (window.ms365FooterHelp && typeof window.ms365FooterHelp.close === 'function') {
                    window.ms365FooterHelp.close();
                }
            },
            true
        );
        document.addEventListener('click', function (ev) {
            const t = ev.target;
            if (!(t instanceof Element)) return;
            const item = t.closest(SAVE_DROP_ITEM);
            if (!item) return;
            const fold = item.closest('details');
            if (fold) {
                setTimeout(function () {
                    fold.open = false;
                }, 0);
            }
        });
    }

    function bindFooterHelpMenu(root, btn, menu) {
        if (!btn || !menu || btn.dataset.bound === '1') {
            if (btn && btn.dataset.bound === '1' && !window.ms365FooterHelp) {
                window.ms365FooterHelp = {
                    open: function () {
                        closeSaveDropMenus();
                        menu.hidden = false;
                        btn.setAttribute('aria-expanded', 'true');
                        root.classList.add('is-open');
                    },
                    close: function () {
                        menu.hidden = true;
                        btn.setAttribute('aria-expanded', 'false');
                        root.classList.remove('is-open');
                    }
                };
            }
            return;
        }
        btn.dataset.bound = '1';
        function setOpen(open) {
            if (open) closeSaveDropMenus();
            menu.hidden = !open;
            btn.setAttribute('aria-expanded', open ? 'true' : 'false');
            root.classList.toggle('is-open', open);
        }
        btn.addEventListener('click', function (e) {
            e.preventDefault();
            e.stopPropagation();
            setOpen(menu.hidden);
        });
        document.addEventListener('click', function (e) {
            if (!root.contains(e.target)) setOpen(false);
        });
        document.addEventListener('keydown', function (e) {
            if (e.key === 'Escape') setOpen(false);
        });
        menu.addEventListener('click', function () {
            setOpen(false);
        });
        window.ms365FooterHelp = {
            open: function () {
                setOpen(true);
            },
            close: function () {
                setOpen(false);
            },
            toggle: function () {
                setOpen(menu.hidden);
            }
        };
    }

    function normalizeAbsorbedHelpLink(link) {
        if (!link) return;
        const href = String(link.getAttribute('href') || '');
        const isHilfe = /hilfe\.html/i.test(href);
        const isNews = /neuigkeiten\.html/i.test(href);
        let icon = link.querySelector('i.bi');
        if (!icon) {
            icon = document.createElement('i');
            icon.setAttribute('aria-hidden', 'true');
            link.insertBefore(icon, link.firstChild);
        }
        let label = link.querySelector('.header-help-link__label, span:not(.sr-only)');
        if (!label) {
            const textNode = Array.prototype.find.call(link.childNodes, function (n) {
                return n.nodeType === Node.TEXT_NODE && String(n.textContent || '').trim();
            });
            if (textNode) {
                label = document.createElement('span');
                label.className = 'header-help-link__label';
                label.textContent = String(textNode.textContent || '').trim();
                textNode.replaceWith(label);
            }
        }
        if (isHilfe) {
            icon.className = 'bi bi-journal-text';
            if (label) label.textContent = 'Hilfeseite';
            link.title = 'Hilfe, Datenschutz & Windows-Hinweise';
            link.setAttribute('aria-label', 'Hilfeseite');
        } else if (isNews) {
            icon.className = 'bi bi-megaphone';
            if (label) label.textContent = 'Neuigkeiten';
        }
    }

    function absorbPageHelpLinks(menu) {
        const landing = menu.querySelector('.site-landing-link');
        document.querySelectorAll('a.header-help-link').forEach(function (link) {
            if (menu.contains(link)) {
                normalizeAbsorbedHelpLink(link);
                return;
            }
            link.classList.add('app-footer-help__item');
            link.setAttribute('role', 'menuitem');
            normalizeAbsorbedHelpLink(link);
            if (landing) menu.insertBefore(link, landing);
            else {
                const sep = menu.querySelector('.app-footer-help__sep');
                if (sep) menu.insertBefore(link, sep);
                else menu.appendChild(link);
            }
        });
        document.querySelectorAll('.header-help-row').forEach(function (row) {
            if (!row.querySelector('a, button')) row.remove();
        });
    }

    function injectSiteCredit() {
        ensureFooterHelpStyles();
        let root = document.getElementById('siteCreditKurtrocks') || document.querySelector('.site-credit-row');
        if (!root) {
            root = document.createElement('div');
            document.body.appendChild(root);
        }
        root.id = 'siteCreditKurtrocks';
        root.className = 'app-footer-help';
        if (root.tagName === 'P') {
            const div = document.createElement('div');
            div.id = 'siteCreditKurtrocks';
            div.className = 'app-footer-help';
            while (root.firstChild) div.appendChild(root.firstChild);
            root.replaceWith(div);
            root = div;
        }

        let btn = root.querySelector('#appFooterHelpBtn');
        let menu = root.querySelector('#appFooterHelpMenu');
        if (!btn || !menu) {
            const leftover = root.querySelector('.site-feedback-link');
            root.replaceChildren();
            btn = document.createElement('button');
            btn.type = 'button';
            btn.id = 'appFooterHelpBtn';
            btn.className = 'app-footer-help__btn';
            btn.setAttribute('aria-haspopup', 'menu');
            btn.setAttribute('aria-expanded', 'false');
            btn.setAttribute('aria-controls', 'appFooterHelpMenu');
            btn.title = 'Hilfe & Infos';
            btn.innerHTML =
                '<i class="bi bi-question-circle" aria-hidden="true"></i><span class="sr-only">Hilfe & Infos</span>';
            menu = document.createElement('div');
            menu.id = 'appFooterHelpMenu';
            menu.className = 'app-footer-help__menu';
            menu.setAttribute('role', 'menu');
            menu.hidden = true;
            const sep = document.createElement('div');
            sep.className = 'app-footer-help__sep';
            sep.setAttribute('aria-hidden', 'true');
            menu.appendChild(sep);
            root.appendChild(btn);
            root.appendChild(menu);
            if (leftover) menu.insertBefore(leftover, sep);
        }

        const legalBase = landingPageHref().replace(/\/?$/, '/');
        const landing = ensureFooterHelpItem(menu, '.site-landing-link', 'site-landing-link', 'bi-globe2', 'Website');
        landing.href = landingPageHref();
        landing.title = 'Website von MS365-Schul-Tools';
        landing.setAttribute('aria-label', 'Website');

        const impressum = ensureFooterHelpItem(
            menu,
            '.site-impressum-link',
            'site-impressum-link',
            'bi-file-earmark-text',
            'Impressum'
        );
        impressum.href = legalBase + 'impressum.html';
        impressum.title = 'Impressum';

        const privacy = ensureFooterHelpItem(
            menu,
            '.site-privacy-link',
            'site-privacy-link',
            'bi-shield-lock',
            'Datenschutz'
        );
        privacy.href = legalBase + 'datenschutz.html';
        privacy.title = 'Datenschutzerklärung';

        const terms = ensureFooterHelpItem(
            menu,
            '.site-terms-link',
            'site-terms-link',
            'bi-journal-text',
            'Nutzungsbedingungen'
        );
        terms.href = legalBase + 'nutzungsbedingungen.html';
        terms.title = 'Nutzungsbedingungen';

        let credit = menu.querySelector(
            '.site-credit-link:not(.site-landing-link):not(.site-impressum-link):not(.site-privacy-link):not(.site-terms-link):not(.site-feedback-link)'
        );
        if (!credit) {
            credit = makeFooterHelpItem('site-credit-link', 'bi-info-circle', 'kurtrocks.education');
            const sep = menu.querySelector('.app-footer-help__sep');
            if (sep && sep.nextSibling) menu.insertBefore(credit, sep.nextSibling);
            else if (sep) sep.after(credit);
            else menu.appendChild(credit);
        } else {
            credit.classList.add('app-footer-help__item', 'site-credit-link');
            const label = credit.querySelector('span');
            if (label) label.textContent = 'kurtrocks.education';
            else {
                const text = Array.prototype.find.call(credit.childNodes, function (n) {
                    return n.nodeType === Node.TEXT_NODE && String(n.textContent || '').trim();
                });
                if (text) text.textContent = 'kurtrocks.education';
            }
        }
        credit.href = KURTROCKS_HREF;
        credit.target = '_blank';
        credit.rel = 'noopener noreferrer';
        credit.title = 'Ein Projekt von Kurt Söser';
        credit.setAttribute('aria-label', 'kurtrocks.education – Ein Projekt von Kurt Söser');

        absorbPageHelpLinks(menu);
        bindFooterHelpMenu(root, btn, menu);
        bindSaveDropClickAway();
        moveFooterItemsIntoFooter();
        stripLooseThemeToggles();
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
            bindSaveDropClickAway();
            moveFooterItemsIntoFooter();
        });
    } else {
        ensureGlobalAuthUi();
        injectSiteCredit();
        bindSaveDropClickAway();
        moveFooterItemsIntoFooter();
    }
    [400, 1200, 3600].forEach(function (ms) {
        setTimeout(function () {
            stripLooseThemeToggles();
            moveFooterItemsIntoFooter();
            bindSaveDropClickAway();
        }, ms);
    });
})();
