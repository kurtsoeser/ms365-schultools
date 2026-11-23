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

    try {
        applyStored(document.documentElement);
    } catch (e) {
        /* ignore */
    }
})();
