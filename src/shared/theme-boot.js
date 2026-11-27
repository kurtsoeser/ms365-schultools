/*
 * Theme und Markenfarbe setzen, bevor die Seite gezeichnet wird (kein Aufblitzen).
 * Klassisches Skript, im <head> bewusst ohne defer/async eingebunden.
 * Früher stand dieser Block inline in jeder Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 *
 * Marke: nur Teal (Klassisch/Wein und weitere Varianten abgekündigt).
 * Erlaubte Marken: dieselbe Liste wie in theme-toggle.js und msal-auth-ui.js.
 */
(function () {
    var BRAND_IDS = ['teal'];
    var DEPRECATED_BRANDS = ['classic', 'wine', 'azure', 'forest', 'amber', 'custom'];
    var BRAND_KEY = 'ms365-brand-v1';
    var CUSTOM_KEY = 'ms365-brand-custom-v1';
    var HEADER_LUM_MAX = 0.55;
    var HEADER_INK = '#0c1222';

    function parseHex(value) {
        if (typeof value !== 'string') return '';
        var s = value.trim();
        var shortHex = /^#([0-9a-fA-F]{3})$/.exec(s);
        if (shortHex) {
            var h = shortHex[1];
            return ('#' + h.charAt(0) + h.charAt(0) + h.charAt(1) + h.charAt(1) + h.charAt(2) + h.charAt(2)).toLowerCase();
        }
        var fullHex = /^#([0-9a-fA-F]{6})$/.exec(s);
        if (fullHex) return ('#' + fullHex[1]).toLowerCase();
        return '';
    }

    function hexToRgb(hex) {
        var n = parseInt(hex.slice(1), 16);
        return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
    }

    function rgbToHex(r, g, b) {
        function channel(v) {
            var s = Math.max(0, Math.min(255, Math.round(v))).toString(16);
            return s.length === 1 ? '0' + s : s;
        }
        return '#' + channel(r) + channel(g) + channel(b);
    }

    function channelLin(c) {
        var x = c / 255;
        return x <= 0.04045 ? x / 12.92 : Math.pow((x + 0.055) / 1.055, 2.4);
    }

    function luminance(hex) {
        var rgb = hexToRgb(hex);
        return 0.2126 * channelLin(rgb.r) + 0.7152 * channelLin(rgb.g) + 0.0722 * channelLin(rgb.b);
    }

    function mixHex(a, b, t) {
        var A = hexToRgb(a);
        var B = hexToRgb(b);
        return rgbToHex(A.r + (B.r - A.r) * t, A.g + (B.g - A.g) * t, A.b + (B.b - A.b) * t);
    }

    function darkenForHeader(hex) {
        var color = parseHex(hex) || hex;
        var guard = 0;
        while (luminance(color) > HEADER_LUM_MAX && guard < 16) {
            color = mixHex(color, HEADER_INK, 0.16);
            guard += 1;
        }
        return color;
    }

    function readCustom() {
        try {
            var raw = localStorage.getItem(CUSTOM_KEY);
            if (!raw) return null;
            var data = JSON.parse(raw);
            var brand1 = parseHex(data && data.brand1);
            var brand2 = parseHex(data && data.brand2);
            if (!brand1 || !brand2) return null;
            return { brand1: brand1, brand2: brand2 };
        } catch (e) {
            return null;
        }
    }

    function clearCustomProps(root) {
        ['--brand1', '--brand2', '--brand-header-1', '--brand-header-2'].forEach(function (name) {
            try {
                root.style.removeProperty(name);
            } catch (e) {
                /* ignore */
            }
        });
    }

    function applyCustomProps(root, seeds) {
        root.style.setProperty('--brand1', seeds.brand1);
        root.style.setProperty('--brand2', seeds.brand2);
        root.style.setProperty('--brand-header-1', darkenForHeader(seeds.brand1));
        root.style.setProperty('--brand-header-2', darkenForHeader(seeds.brand2));
    }

    function syncCustom(root, brand) {
        if (brand === 'custom') {
            var seeds = readCustom();
            if (!seeds) {
                clearCustomProps(root);
                return 'teal';
            }
            applyCustomProps(root, seeds);
            return 'custom';
        }
        clearCustomProps(root);
        return brand;
    }

    function applyStored(root) {
        var theme = 'light';
        try {
            var storedTheme = localStorage.getItem('ms365-theme-v1');
            if (storedTheme === 'dark' || storedTheme === 'light') theme = storedTheme;
        } catch (e) {
            /* ignore */
        }
        root.setAttribute('data-theme', theme);
        try {
            root.style.colorScheme = theme;
        } catch (e2) {
            /* ignore */
        }

        var brand = 'teal';
        try {
            var storedBrand = localStorage.getItem(BRAND_KEY);
            if (BRAND_IDS.indexOf(storedBrand) !== -1) {
                brand = storedBrand;
            } else if (DEPRECATED_BRANDS.indexOf(storedBrand) !== -1) {
                brand = 'teal';
                try {
                    localStorage.setItem(BRAND_KEY, brand);
                } catch (eMigrate) {
                    /* ignore */
                }
            }
        } catch (e3) {
            /* ignore */
        }
        // Custom-Farben abgekündigt: immer Teal-Tokens aus CSS
        clearCustomProps(root);
        brand = 'teal';
        try {
            localStorage.setItem(BRAND_KEY, brand);
        } catch (e4) {
            /* ignore */
        }
        root.setAttribute('data-brand', brand);
    }

    window.ms365BrandBoot = {
        ids: BRAND_IDS,
        customKey: CUSTOM_KEY,
        headerLumMax: HEADER_LUM_MAX,
        parseHex: parseHex,
        readCustom: readCustom,
        luminance: luminance,
        darkenForHeader: darkenForHeader,
        clearCustomProps: clearCustomProps,
        syncCustom: syncCustom,
        applyStored: applyStored
    };

    /**
     * IT-Arbeitsfläche: data-dash-it-shell schon vor dem Paint setzen,
     * sonst blitzt bei jedem Tool-Wechsel kurz Portal-/Alt-Header auf.
     */
    function readItShellHintValue() {
        var preferred = '';
        try {
            preferred = String(localStorage.getItem('ms365-dashboard-it-preview-v1') || '')
                .trim()
                .toLowerCase();
        } catch (e0) {
            preferred = '';
        }
        if (preferred === 'full' || preferred === 'it') return 'full';
        // Immer alle Keys scannen: Tenant-Key kann full sein, Unscoped noch portal/leer.
        try {
            for (var i = 0; i < localStorage.length; i++) {
                var k = localStorage.key(i);
                if (!k || k.indexOf('ms365-dashboard-it-preview-v1') !== 0) continue;
                var v = String(localStorage.getItem(k) || '')
                    .trim()
                    .toLowerCase();
                if (v === 'full' || v === 'it') return 'full';
            }
        } catch (e1) {
            /* ignore */
        }
        if (preferred === 'schueler') return 'schueler';
        return preferred || '';
    }

    function clearItShellSkeleton() {
        try {
            var sk = document.getElementById('ms365ItShellSkeleton');
            if (sk && sk.parentNode) sk.parentNode.removeChild(sk);
            document.documentElement.removeAttribute('data-ms365-it-fouc');
        } catch (eClear) {
            /* ignore */
        }
    }

    function injectItShellSkeleton() {
        function place() {
            try {
                if (!document.body || document.getElementById('ms365ItShellSkeleton')) return;
                if (!document.documentElement.hasAttribute('data-ms365-it-fouc')) return;
                if (document.body.classList.contains('it-workspace-shell')) return;
                var el = document.createElement('div');
                el.id = 'ms365ItShellSkeleton';
                el.className = 'it-shell-skel';
                el.setAttribute('role', 'status');
                el.setAttribute('aria-busy', 'true');
                el.setAttribute('aria-live', 'polite');
                el.setAttribute('aria-label', 'IT-Arbeitsfläche wird geladen');
                el.innerHTML =
                    '<div class="it-shell-skel__top" aria-hidden="true">' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--brand"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--search"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--chip"></span>' +
                    '</div>' +
                    '<div class="it-shell-skel__body" aria-hidden="true">' +
                    '<aside class="it-shell-skel__rail">' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--nav"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--nav"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--nav is-wide"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--nav"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--nav is-active"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--sub"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--sub"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--sub"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--nav"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--nav"></span>' +
                    '</aside>' +
                    '<div class="it-shell-skel__main">' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--title"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--lead"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--row"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--row"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--row"></span>' +
                    '<span class="it-shell-skel__bar it-shell-skel__bar--row is-short"></span>' +
                    '</div>' +
                    '</div>';
                document.body.appendChild(el);
            } catch (ePlace) {
                /* ignore */
            }
        }
        if (document.body) {
            place();
            return;
        }
        if (typeof MutationObserver !== 'undefined') {
            var obs = new MutationObserver(function () {
                if (document.body) {
                    obs.disconnect();
                    place();
                }
            });
            try {
                obs.observe(document.documentElement, { childList: true });
            } catch (eObs) {
                document.addEventListener('DOMContentLoaded', place);
            }
        } else {
            document.addEventListener('DOMContentLoaded', place);
        }
    }

    function injectEarlyItShellAssets(inTools) {
        try {
            if (document.getElementById('ms365-it-fouc-css')) return;
            var prefix = inTools ? '../' : '';
            var style = document.createElement('style');
            style.id = 'ms365-it-fouc-css';
            // Legacy unsichtbar + Skeleton-Vorschau der IT-Shell (wie in Planern).
            style.textContent =
                'html[data-ms365-it-fouc] body:not(.it-workspace-shell) > .container,' +
                'html[data-ms365-it-fouc] body:not(.it-workspace-shell) > .container.page-card{' +
                'visibility:hidden!important;}' +
                'html[data-ms365-it-fouc] body.it-workspace-shell #itWorkspaceShell{visibility:visible!important;}' +
                'html[data-ms365-it-fouc] body.it-workspace-shell #ms365ItShellSkeleton{display:none!important;}' +
                'html[data-ms365-it-fouc] #ms365LicenseGateOverlay,' +
                'html[data-ms365-it-fouc] .ms365-license-gate,' +
                'html[data-ms365-it-fouc] #portalWorkspaceRail,' +
                'html[data-ms365-it-embed] #ms365LicenseGateOverlay,' +
                'html[data-ms365-it-embed] .ms365-license-gate,' +
                'html[data-ms365-it-embed] #portalWorkspaceRail,' +
                'html[data-ms365-it-embed] #itWorkspaceShell,' +
                'html[data-ms365-it-embed] #dashCompactHeader,' +
                'html[data-ms365-it-embed] #ms365AuthWidget,' +
                'html[data-ms365-it-embed] #ms365ItShellSkeleton{display:none!important;}' +
                'html[data-ms365-it-embed] body > .container > .header:not(.dash-compact-header){' +
                'background:transparent!important;color:var(--heading,#0c1222)!important;box-shadow:none!important;}' +
                'html[data-ms365-it-embed] .dash-tasks-split__nav,' +
                'html[data-ms365-it-embed] .dash-tasks-split__resize{display:none!important;}' +
                'html[data-ms365-it-embed] .dash-tasks-split{grid-template-columns:minmax(0,1fr)!important;}' +
                'html[data-ms365-it-embed] .dash-tasks-split__detail{grid-column:1!important;}' +
                'html[data-ms365-it-embed] .ts-admin-rail{display:none!important;}' +
                'html[data-ms365-it-embed] .ts-register-shell.ts-admin-shell{grid-template-columns:minmax(0,1fr)!important;}' +
                'html[data-ms365-it-fouc] #ms365AuthWidget .btn,' +
                'html[data-ms365-it-fouc] #ms365AuthWidget button{visibility:hidden!important;}' +
                '.it-shell-skel{position:fixed;inset:0;z-index:45;display:flex;flex-direction:column;' +
                'background:var(--paper,#f6f7f7);pointer-events:none;}' +
                'html[data-theme=dark] .it-shell-skel{background:#181818;}' +
                '.it-shell-skel__top{display:flex;align-items:center;gap:14px;height:56px;padding:0 18px;' +
                'border-bottom:1px solid var(--border,rgba(12,18,34,.08));background:var(--card,#fff);flex:0 0 auto;}' +
                'html[data-theme=dark] .it-shell-skel__top{background:#1f1f1f;border-color:rgba(255,255,255,.1);}' +
                '.it-shell-skel__body{flex:1;min-height:0;display:grid;grid-template-columns:256px minmax(0,1fr);}' +
                '.it-shell-skel__rail{padding:16px 12px;border-right:1px solid var(--border,rgba(12,18,34,.08));' +
                'background:var(--rail-bg,#f6f7f7);display:flex;flex-direction:column;gap:10px;}' +
                'html[data-theme=dark] .it-shell-skel__rail{background:#1b1b1b;border-color:rgba(255,255,255,.1);}' +
                '.it-shell-skel__main{padding:28px 36px;display:flex;flex-direction:column;gap:14px;}' +
                '.it-shell-skel__bar{display:block;border-radius:8px;height:12px;' +
                'background:linear-gradient(90deg,color-mix(in srgb,var(--heading,#0c1222) 8%,transparent) 20%,' +
                'color-mix(in srgb,var(--heading,#0c1222) 3%,transparent) 40%,' +
                'color-mix(in srgb,var(--heading,#0c1222) 8%,transparent) 60%);' +
                'background-size:200% 100%;animation:it-shell-skel-shimmer 1.25s linear infinite;}' +
                '.it-shell-skel__bar--brand{width:132px;height:16px;}' +
                '.it-shell-skel__bar--search{flex:1;max-width:280px;height:32px;border-radius:10px;}' +
                '.it-shell-skel__bar--chip{width:72px;height:28px;margin-left:auto;border-radius:10px;}' +
                '.it-shell-skel__bar--nav{width:78%;height:28px;border-radius:8px;}' +
                '.it-shell-skel__bar--nav.is-wide{width:88%;}' +
                '.it-shell-skel__bar--nav.is-active{width:84%;' +
                'background:color-mix(in srgb,var(--brand1,#0d9f8a) 18%,transparent);animation:none;}' +
                '.it-shell-skel__bar--sub{width:62%;height:22px;margin-left:14px;border-radius:6px;}' +
                '.it-shell-skel__bar--title{width:min(280px,55%);height:28px;margin-top:8px;}' +
                '.it-shell-skel__bar--lead{width:min(420px,70%);height:14px;}' +
                '.it-shell-skel__bar--row{width:100%;max-width:720px;height:52px;border-radius:10px;}' +
                '.it-shell-skel__bar--row.is-short{width:72%;max-width:520px;}' +
                '@keyframes it-shell-skel-shimmer{to{background-position:-200% 0;}}' +
                '@media (max-width:900px){.it-shell-skel__body{grid-template-columns:1fr;}' +
                '.it-shell-skel__rail{display:none;}}';
            document.head.appendChild(style);

            function ensureLink(id, href) {
                if (document.getElementById(id)) return;
                var link = document.createElement('link');
                link.id = id;
                link.rel = 'stylesheet';
                link.href = prefix + href;
                document.head.appendChild(link);
            }
            ensureLink('ms365-it-workspace-shell-css', 'src/shared/it-workspace-shell.css?v=10');
            ensureLink('ms365-dashboard-compact-header-css-early', 'src/shared/dashboard-compact-header.css?v=16');
        } catch (eCss) {
            /* ignore */
        }
    }

    function applyItShellHint(root) {
        try {
            var path = String((typeof location !== 'undefined' && location.pathname) || '').replace(/\\/g, '/');
            var inTools = /\/tools\//i.test(path);
            var isIndex = /(?:^|\/)index\.html$/i.test(path) || /\/$/.test(path);
            if (!inTools && !isIndex && !/(tenant|dashboard-werkzeug-zugriff)\.html$/i.test(path)) {
                return;
            }
            var embed = false;
            try {
                embed = window.self !== window.top;
            } catch (eEmb) {
                embed = true;
            }
            // iframe-Inhalt in der IT-Shell: ruhig, ohne zweite Rail/Portal/Lizenz/Auth-Flash.
            if (embed) {
                root.setAttribute('data-ms365-it-embed', '');
                if (readItShellHintValue() === 'full') {
                    root.setAttribute('data-dash-it-shell', '');
                    root.setAttribute('data-frontend-planner-it', '');
                }
                injectEarlyItShellAssets(inTools);
                return;
            }
            var shell = readItShellHintValue();
            if (shell !== 'full') return;
            // Shell-Wahl full reicht für FOUC auf Tool-Seiten – nicht nur Session-Flag
            // (Flag wurde früher bei pending-directory gelöscht → Lehrer-Flash).
            try {
                sessionStorage.setItem('ms365-it-workspace-v1', '1');
            } catch (eFlag) {
                /* ignore */
            }
            root.setAttribute('data-dash-it-shell', '');
            root.setAttribute('data-frontend-planner-it', '');
            root.setAttribute('data-ms365-it-fouc', '');
            injectEarlyItShellAssets(inTools);
            injectItShellSkeleton();
        } catch (eHint) {
            /* ignore */
        }
    }

    window.ms365ClearItShellSkeleton = clearItShellSkeleton;

    try {
        applyStored(document.documentElement);
        applyItShellHint(document.documentElement);
        // Sicherheitsnetz: Skeleton/Hide nie dauerhaft, falls Shell-Mount scheitert.
        if (document.documentElement.hasAttribute('data-ms365-it-fouc')) {
            setTimeout(function () {
                try {
                    if (
                        document.documentElement.hasAttribute('data-ms365-it-fouc') &&
                        !document.body.classList.contains('it-workspace-shell')
                    ) {
                        clearItShellSkeleton();
                    }
                } catch (eClear) {
                    /* ignore */
                }
            }, 2500);
        }
    } catch (e) {
        /* ignore */
    }
})();
