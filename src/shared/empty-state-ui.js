(function () {
    'use strict';

    function hasMeaningfulTenantData(settings) {
        if (window.ms365DemoMode && typeof window.ms365DemoMode.hasMeaningfulTenantData === 'function') {
            return window.ms365DemoMode.hasMeaningfulTenantData(settings);
        }
        if (!settings || typeof settings !== 'object') return false;
        var domain = String(settings.domain || '').trim();
        var schoolName = String(settings.schoolName || '').trim();
        var subjects = Array.isArray(settings.subjects) ? settings.subjects.length : 0;
        var teachers = Array.isArray(settings.teachers) ? settings.teachers.length : 0;
        var students = Array.isArray(settings.students) ? settings.students.length : 0;
        var classes = Array.isArray(settings.classes) ? settings.classes.length : 0;
        return !!(domain || schoolName || subjects || teachers || students || classes);
    }

    function loadSettings() {
        try {
            if (typeof window.ms365TenantSettingsLoad === 'function') {
                return window.ms365TenantSettingsLoad();
            }
        } catch {
            /* ignore */
        }
        return null;
    }

    function isDemoMode() {
        return !!(window.ms365DemoMode && window.ms365DemoMode.isActive());
    }

    function setupFinished() {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                var setup = window.ms365AppDataV2.getSetup();
                return !!(setup && setup.finishedAt);
            }
        } catch {
            /* ignore */
        }
        return false;
    }

    function resolveHref(path) {
        try {
            var p = String(location.pathname || '');
            if (/\/tools\//i.test(p)) return '../' + path.replace(/^\//, '');
        } catch {
            /* ignore */
        }
        return path;
    }

    function createBanner(message, options) {
        var opts = options || {};
        var banner = document.createElement('div');
        banner.className = 'ms365-empty-state-banner';
        banner.setAttribute('role', 'status');

        var p = document.createElement('p');
        if (opts.html) p.innerHTML = message;
        else p.textContent = message;
        banner.appendChild(p);

        var actions = document.createElement('div');
        actions.className = 'ms365-empty-state-banner__actions';

        (opts.actions || []).forEach(function (action) {
            if (action.tag === 'button') {
                var btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn' + (action.ghost ? ' btn-ghost' : '');
                btn.innerHTML = action.label;
                if (action.id) btn.id = action.id;
                if (typeof action.onClick === 'function') {
                    btn.addEventListener('click', action.onClick);
                }
                actions.appendChild(btn);
            } else {
                var a = document.createElement('a');
                a.className = 'btn' + (action.ghost ? ' btn-ghost' : '');
                a.href = resolveHref(action.href || '#');
                a.innerHTML = action.label;
                actions.appendChild(a);
            }
        });

        banner.appendChild(actions);
        return banner;
    }

    function defaultActions() {
        return [
            { href: 'einrichtung.html', label: '<i class="bi bi-rocket-takeoff"></i>Einrichtung starten' },
            { href: 'tenant.html', label: '<i class="bi bi-gear"></i>Stammdaten', ghost: true },
            {
                href: 'index.html#start-demo',
                label: '<i class="bi bi-play-circle"></i>Demo ausprobieren',
                ghost: true
            }
        ];
    }

    function mountEmptyStateTargets(root) {
        var scope = root || document;
        var targets = scope.querySelectorAll('[data-ms365-empty-state]');
        targets.forEach(function (target) {
            var settings = loadSettings();
            var hasData = hasMeaningfulTenantData(settings);
            var demo = isDemoMode();

            if (hasData || demo) {
                target.hidden = true;
                target.replaceChildren();
                return;
            }

            var customMsg = target.getAttribute('data-ms365-empty-message');
            var message =
                customMsg ||
                'Noch keine Stammdaten. Dieses Werkzeug braucht Listen aus der Einrichtung oder den Stammdaten – oder starten Sie die Demo auf dem Dashboard.';

            target.hidden = false;
            target.replaceChildren();
            target.appendChild(
                createBanner(message, {
                    actions: defaultActions()
                })
            );
        });
    }

    function refreshTenantEmptyBanner() {
        var mount = document.getElementById('tenantEmptyStateMount');
        if (!mount) return;

        var settings = loadSettings();
        var hasData = hasMeaningfulTenantData(settings);
        var demo = isDemoMode();

        if (hasData) {
            mount.hidden = true;
            mount.replaceChildren();
            return;
        }

        mount.hidden = false;
        mount.replaceChildren();
        mount.appendChild(
            createBanner(
                'Noch keine Schuldaten erfasst. Starten Sie die geführte Einrichtung – oder pflegen Sie Domain und Listen direkt hier. Alternativ: Demo auf dem Dashboard.',
                {
                    actions: [
                        { href: 'einrichtung.html', label: '<i class="bi bi-rocket-takeoff"></i>Einrichtung starten' },
                        {
                            tag: 'button',
                            id: 'tenantBtnStartDemo',
                            label: '<i class="bi bi-play-circle"></i>Demo laden',
                            ghost: true,
                            onClick: function () {
                                if (window.ms365DemoMode && window.ms365DemoMode.activate()) {
                                    window.location.reload();
                                } else {
                                    window.location.href = 'index.html#start-demo';
                                }
                            }
                        }
                    ]
                }
            )
        );
    }

    function shouldShowHygieneTeaser() {
        if (isDemoMode()) return true;
        return setupFinished() && hasMeaningfulTenantData(loadSettings());
    }

    window.ms365EmptyStateUi = {
        hasMeaningfulTenantData: hasMeaningfulTenantData,
        setupFinished: setupFinished,
        shouldShowHygieneTeaser: shouldShowHygieneTeaser,
        mountEmptyStateTargets: mountEmptyStateTargets,
        refreshTenantEmptyBanner: refreshTenantEmptyBanner
    };

    function boot() {
        mountEmptyStateTargets(document);
        refreshTenantEmptyBanner();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
    } else {
        boot();
    }

    window.addEventListener('ms365-tenant-settings-changed', boot);
    window.addEventListener('ms365-demo-mode-changed', boot);
})();
