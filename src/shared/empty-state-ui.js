(function () {
    'use strict';

    function hasMeaningfulTenantData(settings) {
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
            { href: 'tenant.html', label: '<i class="bi bi-journal-bookmark"></i>Stammdaten öffnen' },
            {
                href: 'tools/webuntis-stammdaten-import.html?from=tenant',
                label: '<i class="bi bi-box-arrow-in-down"></i>Daten importieren',
                ghost: true
            },
            {
                tag: 'button',
                id: 'emptyStateImportBackup',
                label: '<i class="bi bi-upload"></i>Backup importieren',
                ghost: true,
                onClick: function () {
                    var input = document.getElementById('browserBackupImportFile');
                    if (input) {
                        input.click();
                        return;
                    }
                    window.location.href = resolveHref('index.html#dashboard-local');
                }
            }
        ];
    }

    function mountEmptyStateTargets(root) {
        var scope = root || document;
        var targets = scope.querySelectorAll('[data-ms365-empty-state]');
        targets.forEach(function (target) {
            var settings = loadSettings();
            var hasData = hasMeaningfulTenantData(settings);

            if (hasData) {
                target.hidden = true;
                target.replaceChildren();
                return;
            }

            var customMsg = target.getAttribute('data-ms365-empty-message');
            var message =
                customMsg ||
                'Noch keine Stammdaten in <strong>diesem Browser</strong>. Gruppen in Microsoft&nbsp;365 sind davon unabhängig – hier fehlen nur die lokalen Listen. Im <strong>Stammdaten</strong> starten, Daten importieren oder ein vorhandenes <strong>Browser-Backup</strong> laden.';

            target.hidden = false;
            target.replaceChildren();
            target.appendChild(
                createBanner(message, {
                    html: /<[a-z][\s\S]*>/i.test(message),
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

        if (hasData) {
            mount.hidden = true;
            mount.replaceChildren();
            return;
        }

        mount.hidden = false;
        mount.replaceChildren();
        mount.appendChild(
            createBanner(
                'Noch keine Schuldaten in diesem Browser. Die Microsoft-365-Gruppen existieren unabhängig davon – hier fehlen nur die lokalen Stammdaten. Stammdaten öffnen, Daten importieren oder ein Browser-Backup (JSON) laden.',
                {
                    actions: [
                        { href: 'tenant.html', label: '<i class="bi bi-journal-bookmark"></i>Stammdaten öffnen' },
                        {
                            href: 'tools/webuntis-stammdaten-import.html?from=tenant',
                            label: '<i class="bi bi-box-arrow-in-down"></i>Daten importieren',
                            ghost: true
                        },
                        {
                            tag: 'button',
                            id: 'tenantBtnImportBackup',
                            label: '<i class="bi bi-upload"></i>Backup importieren',
                            ghost: true,
                            onClick: function () {
                                var input =
                                    document.getElementById('browserBackupImportFile') ||
                                    document.querySelector('[data-ms365-backup="import-file"]');
                                if (input) input.click();
                            }
                        }
                    ]
                }
            )
        );
    }

    function shouldShowHygieneTeaser() {
        return hasMeaningfulTenantData(loadSettings());
    }

    window.ms365EmptyStateUi = {
        hasMeaningfulTenantData: hasMeaningfulTenantData,
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
})();
