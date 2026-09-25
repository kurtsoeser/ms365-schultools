/**
 * Light/Dark + Brand-Theme (teal | classic) für die gesamte App.
 * Standard: Hellmodus + Brand teal (Landing).
 * Speichert: ms365-theme-v1, ms365-brand-v1.
 * Hell/Dunkel-Schalter: Fußzeile. Brand: Konto-Menü oben rechts.
 */
(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-theme-v1';
    const BRAND_KEY = 'ms365-brand-v1';
    const THEMES = { light: 'light', dark: 'dark' };
    const BRANDS = { teal: 'teal', classic: 'classic' };

    function preferredTheme() {
        return THEMES.light;
    }

    function preferredBrand() {
        return BRANDS.teal;
    }

    function readStored() {
        try {
            const v = localStorage.getItem(STORAGE_KEY);
            if (v === THEMES.dark || v === THEMES.light) return v;
        } catch {
            /* ignore */
        }
        return null;
    }

    function readStoredBrand() {
        try {
            const v = localStorage.getItem(BRAND_KEY);
            if (v === BRANDS.classic || v === BRANDS.teal) return v;
        } catch {
            /* ignore */
        }
        return null;
    }

    function currentTheme() {
        const attr = document.documentElement.getAttribute('data-theme');
        if (attr === THEMES.dark || attr === THEMES.light) return attr;
        return preferredTheme();
    }

    function currentBrand() {
        const attr = document.documentElement.getAttribute('data-brand');
        if (attr === BRANDS.classic || attr === BRANDS.teal) return attr;
        return preferredBrand();
    }

    function applyTheme(theme) {
        const next = theme === THEMES.dark ? THEMES.dark : THEMES.light;
        document.documentElement.setAttribute('data-theme', next);
        try {
            document.documentElement.style.colorScheme = next;
        } catch {
            /* ignore */
        }
        try {
            localStorage.setItem(STORAGE_KEY, next);
        } catch {
            /* ignore */
        }
        syncButtons();
        syncBrandButtons();
        try {
            window.dispatchEvent(new CustomEvent('ms365-theme-change', { detail: { theme: next } }));
        } catch {
            /* ignore */
        }
        return next;
    }

    function syncBrandLogos() {
        const brand = currentBrand();
        const classic = brand === BRANDS.classic;
        document.querySelectorAll('.app-brand-logo').forEach(function (wrap) {
            const teal = wrap.querySelector('.app-brand-logo__mark--teal');
            const classicImg = wrap.querySelector('.app-brand-logo__mark--classic');
            if (teal) {
                teal.setAttribute('aria-hidden', classic ? 'true' : 'false');
                if (!classic) teal.setAttribute('alt', 'Logo MS365-Schul-Tools');
                else teal.setAttribute('alt', '');
            }
            if (classicImg) {
                classicImg.setAttribute('aria-hidden', classic ? 'false' : 'true');
                if (classic) classicImg.setAttribute('alt', 'Logo MS365-Schul-Tools');
                else classicImg.setAttribute('alt', '');
            }
        });
    }

    function applyBrand(brand) {
        const next = brand === BRANDS.classic ? BRANDS.classic : BRANDS.teal;
        document.documentElement.setAttribute('data-brand', next);
        try {
            localStorage.setItem(BRAND_KEY, next);
        } catch {
            /* ignore */
        }
        syncBrandButtons();
        syncBrandLogos();
        try {
            window.dispatchEvent(new CustomEvent('ms365-brand-change', { detail: { brand: next } }));
        } catch {
            /* ignore */
        }
        return next;
    }

    function toggleTheme() {
        return applyTheme(currentTheme() === THEMES.dark ? THEMES.light : THEMES.dark);
    }

    function labelFor(theme) {
        return theme === THEMES.dark ? 'Hell' : 'Dunkel';
    }

    function iconFor(theme) {
        return theme === THEMES.dark
            ? '<i class="bi bi-sun" aria-hidden="true"></i>'
            : '<i class="bi bi-moon-stars" aria-hidden="true"></i>';
    }

    function syncButtons() {
        const theme = currentTheme();
        document.querySelectorAll('[data-ms365-theme-toggle]').forEach(function (btn) {
            btn.setAttribute('aria-label', 'Darstellung umschalten: ' + labelFor(theme) + 'modus');
            btn.setAttribute('title', labelFor(theme) + 'modus');
            btn.innerHTML = iconFor(theme) + '<span>' + labelFor(theme) + '</span>';
        });
    }

    function syncBrandButtons() {
        const brand = currentBrand();
        document.querySelectorAll('[data-ms365-brand]').forEach(function (btn) {
            const b = btn.getAttribute('data-ms365-brand');
            const on = b === brand;
            btn.setAttribute('aria-checked', on ? 'true' : 'false');
            btn.classList.toggle('is-active', on);
        });
    }

    function ensureCreditRow() {
        let row = document.getElementById('siteCreditKurtrocks');
        if (row) return row;
        row = document.querySelector('.site-credit-row');
        return row || null;
    }

    function removeHeaderToggles() {
        document.querySelectorAll('.header [data-ms365-theme-toggle], .ms365-header-tools [data-ms365-theme-toggle]').forEach(function (el) {
            el.remove();
        });
    }

    function mountToggle() {
        removeHeaderToggles();
        const row = ensureCreditRow();
        if (!row) return null;

        let btn = row.querySelector('[data-ms365-theme-toggle]');
        if (!btn) {
            btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'ms365-theme-toggle';
            btn.setAttribute('data-ms365-theme-toggle', '1');
            btn.addEventListener('click', function (e) {
                e.preventDefault();
                toggleTheme();
            });
            const link = row.querySelector('.site-credit-link');
            if (link) row.insertBefore(btn, link);
            else row.appendChild(btn);
        }
        syncButtons();
        syncBrandButtons();
        return btn;
    }

    function mountWhenReady(attemptsLeft) {
        if (mountToggle()) return;
        if (attemptsLeft <= 0) return;
        setTimeout(function () {
            mountWhenReady(attemptsLeft - 1);
        }, 80);
    }

    function init() {
        const stored = readStored();
        applyTheme(stored || preferredTheme());
        applyBrand(readStoredBrand() || preferredBrand());
        mountWhenReady(40);
    }

    window.ms365Theme = {
        get: currentTheme,
        set: applyTheme,
        toggle: toggleTheme,
        mount: mountToggle,
        getBrand: currentBrand,
        setBrand: applyBrand,
        brands: BRANDS
    };

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();
