(function () {
    'use strict';

    var path = String(location.pathname || '').replace(/\\/g, '/');
    if (/\/welcome\.html(?:\?|#|$)/i.test(path)) return;
    if (/\/ms365-schooltool\.html(?:\?|#|$)/i.test(path)) return;

    function readContext() {
        var year = '';
        var domain = '';
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                var c = window.ms365AppDataV2.getContainer();
                year = String((c && c.years && c.years.current) || '').trim();
                domain = String((c && c.core && c.core.domain) || '').trim();
            }
        } catch {
            /* ignore */
        }
        if (!domain && typeof window.ms365TenantSettingsLoad === 'function') {
            try {
                var s = window.ms365TenantSettingsLoad();
                domain = String((s && s.domain) || '').trim();
            } catch {
                /* ignore */
            }
        }
        if (!domain && typeof window.ms365GetSchoolDomainNoAt === 'function') {
            try {
                domain = String(window.ms365GetSchoolDomainNoAt() || '').trim();
            } catch {
                /* ignore */
            }
        }
        try {
            var inp = document.getElementById('schoolEmailDomain');
            if (inp && String(inp.value || '').trim()) domain = String(inp.value).trim();
        } catch {
            /* ignore */
        }
        try {
            var yearSel = document.getElementById('schoolYearSelect');
            if (yearSel && String(yearSel.value || '').trim()) year = String(yearSel.value).trim();
        } catch {
            /* ignore */
        }
        if (!year || !domain) {
            try {
                var raw = localStorage.getItem('ms365-schooltool-data-v2');
                var o = raw ? JSON.parse(raw) : null;
                if (o && typeof o === 'object') {
                    if (!year) year = String((o.years && o.years.current) || '').trim();
                    if (!domain) domain = String((o.core && o.core.domain) || '').trim();
                }
            } catch {
                /* ignore */
            }
        }
        if (!domain) {
            try {
                var tRaw = localStorage.getItem('ms365-tenant-settings-v1');
                var t = tRaw ? JSON.parse(tRaw) : null;
                if (t && t.domain) domain = String(t.domain).trim();
            } catch {
                /* ignore */
            }
        }
        return { year: year, domain: domain };
    }

    function isDashboardLink(a) {
        var href = String(a.getAttribute('href') || '').replace(/\\/g, '/');
        if (!/(?:^|\/)index\.html(?:\?|#|$)/i.test(href)) return false;
        var label = String(a.textContent || '')
            .replace(/\s+/g, ' ')
            .trim();
        return /dashboard/i.test(label);
    }

    function placeDashboardNav(header) {
        var nav = document.getElementById('ms365HeaderNav');
        if (nav && nav.querySelector('a[href]')) return;
        var toolbar = header.querySelector('.toolbar');
        if (!toolbar) return;
        var found = null;
        var links = toolbar.querySelectorAll('a[href]');
        for (var i = 0; i < links.length; i++) {
            if (!isDashboardLink(links[i])) continue;
            found = links[i];
            break;
        }
        if (!found) return;
        if (!nav) {
            nav = document.createElement('nav');
            nav.id = 'ms365HeaderNav';
            nav.className = 'ms365-header-nav';
            nav.setAttribute('aria-label', 'Zurück zum Dashboard');
            header.insertBefore(nav, header.firstChild);
        }
        nav.appendChild(found);
    }

    function hideEmptyToolbar(header) {
        var toolbar = header.querySelector('.toolbar');
        if (!toolbar) return;
        var keep = toolbar.querySelector('a[href], button, input, select, textarea');
        toolbar.hidden = !keep;
    }

    function fillAccountContext(ctx) {
        var yearEl = document.getElementById('ms365AuthCtxYear');
        var domainEl = document.getElementById('ms365AuthCtxDomain');
        if (yearEl) yearEl.textContent = ctx.year || '–';
        if (domainEl) domainEl.textContent = ctx.domain || '–';
        return !!(yearEl || domainEl);
    }

    var headerReadySent = false;

    function render() {
        var header = document.querySelector('.header');
        if (!header) return false;

        placeDashboardNav(header);
        hideEmptyToolbar(header);
        fillAccountContext(readContext());
        if (!headerReadySent) {
            headerReadySent = true;
            try {
                window.dispatchEvent(new CustomEvent('ms365-menu-header-ready'));
            } catch {
                /* ignore */
            }
        }
        return true;
    }

    function boot() {
        var tries = 0;
        var t = setInterval(function () {
            tries += 1;
            if (render() || tries > 40) clearInterval(t);
        }, 50);
        document.addEventListener('input', function (e) {
            var el = e.target;
            if (!el || !el.id) return;
            if (el.id === 'schoolEmailDomain' || el.id === 'schoolYearSelect') render();
        });
        document.addEventListener('change', function (e) {
            var el = e.target;
            if (!el || !el.id) return;
            if (el.id === 'schoolEmailDomain' || el.id === 'schoolYearSelect') render();
        });
        window.addEventListener('ms365-tenant-settings-changed', function () {
            render();
        });
        window.addEventListener('ms365-auth-widget-ready', function () {
            fillAccountContext(readContext());
        });
        window.ms365RefreshContextBar = render;
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
    } else {
        boot();
    }
})();
