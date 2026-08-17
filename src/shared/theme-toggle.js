/**
 * Light/Dark-Theme für die gesamte App.
 * Speichert die Wahl in localStorage (ms365-theme-v1).
 * Schalter: unten links, vor dem Link kurtrocks.com.
 */
(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-theme-v1';
    const THEMES = { light: 'light', dark: 'dark' };

    function preferredTheme() {
        try {
            if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
                return THEMES.dark;
            }
        } catch {
            /* ignore */
        }
        return THEMES.light;
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

    function currentTheme() {
        const attr = document.documentElement.getAttribute('data-theme');
        if (attr === THEMES.dark || attr === THEMES.light) return attr;
        return preferredTheme();
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
        try {
            window.dispatchEvent(new CustomEvent('ms365-theme-change', { detail: { theme: next } }));
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
        mountWhenReady(40);
    }

    window.ms365Theme = {
        get: currentTheme,
        set: applyTheme,
        toggle: toggleTheme,
        mount: mountToggle
    };

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();
