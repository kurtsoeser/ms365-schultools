/**
 * Admin-Setup: Status-Klarheit, Spalten-Check, API-Diagnose.
 */
(function () {
    'use strict';

    function $(id) {
        return document.getElementById(id);
    }

    function toast(m) {
        if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
        else window.alert(m);
    }

    function apiBase() {
        var cfg = window.MS365_LICENSE_API || {};
        return String(cfg.baseUrl || '')
            .trim()
            .replace(/\/+$/, '');
    }

    function setCheck(el, state, title, detail) {
        if (!el) return;
        el.className = 'admin-setup-check admin-setup-check--' + state;
        var icon = el.querySelector('.admin-setup-check__icon');
        var t = el.querySelector('.admin-setup-check__title');
        var d = el.querySelector('.admin-setup-check__detail');
        if (icon) {
            icon.className =
                'bi admin-setup-check__icon ' +
                (state === 'ok'
                    ? 'bi-check-circle-fill'
                    : state === 'warn'
                      ? 'bi-exclamation-circle-fill'
                      : state === 'off'
                        ? 'bi-dash-circle'
                        : 'bi-hourglass-split');
        }
        if (t) t.textContent = title;
        if (d) d.textContent = detail || '';
    }

    function refreshApiConfigUi() {
        var base = apiBase();
        var el = $('licApiBaseUrl');
        if (el) el.textContent = base || 'nicht gesetzt';
        var chip = $('licApiConfigChip');
        if (chip) {
            chip.textContent = base ? 'API konfiguriert' : 'API fehlt';
            chip.className =
                'admin-app__pill' + (base ? ' admin-app__pill--ok' : ' admin-app__pill--warn');
        }
        setCheck(
            $('licCheckApi'),
            base ? 'ok' : 'warn',
            base ? 'License-API erreichbar konfiguriert' : 'License-API noch einrichten',
            base
                ? 'Base-URL ist gesetzt (GitHub Secret / ms365-config.local.js).'
                : 'Ohne baseUrl gibt es kein Lizenz-Gate und keine Freischaltung über die API.'
        );
    }

    function isAuthInteractionError(e) {
        var msg = String((e && e.message) || e || '');
        var code = String((e && (e.errorCode || e.code)) || '');
        return /monitor_window_timeout|interaction_required|login_required|consent_required|popup_window|iframe|silent.?token|nicht angemeldet/i.test(
            msg + ' ' + code
        );
    }

    function friendlySchemaAuthDetail() {
        return 'Anmeldung für die Prüfung nötig. Bitte „Status prüfen“ klicken und ggf. das Anmeldefenster erlauben.';
    }

    /**
     * Schema über License-API (Server-Graph) – ohne Sites.Manage.All im Browser/iframe.
     * @param {{ interactive?: boolean }} [opts]
     * @returns {Promise<boolean>} true wenn erledigt
     */
    async function refreshSchemaViaLicenseApi(opts) {
        var interactive = !!(opts && opts.interactive);
        var api = window.ms365LicenseApi;
        if (!apiBase() || !api || typeof api.adminEnsureSchema !== 'function') return false;
        if (typeof api.acquireLicenseToken !== 'function') return false;

        var token = await api.acquireLicenseToken(
            interactive ? { popup: true } : { silentOnly: true }
        );
        var result = await api.adminEnsureSchema(token);
        var stamm = result && result.stammdaten ? result.stammdaten : {};
        var created = Array.isArray(stamm.created) ? stamm.created : [];
        var existing = Array.isArray(stamm.existing) ? stamm.existing : [];
        var detailParts = [];
        if (created.length) detailParts.push('neu: ' + created.join(', '));
        if (existing.length) {
            detailParts.push(
                existing.length + ' Stammdaten-Spalte(n) vorhanden' + (created.length ? '' : '')
            );
        }
        if (result && result.toolSiteUrlColumn && result.toolSiteUrlColumn.created) {
            detailParts.push('ToolSiteUrl angelegt');
        }
        setCheck(
            $('licCheckSchema'),
            'ok',
            created.length ? 'Schema aktualisiert' : 'Liste & Spalten vollständig',
            detailParts.length
                ? detailParts.join(' · ')
                : 'License-API hat das Schema geprüft (inkl. Rollen-E-Mail-Spalten).'
        );
        return true;
    }

    /**
     * @param {{ interactive?: boolean }} [opts]
     */
    async function refreshSchemaStatus(opts) {
        var interactive = !!(opts && opts.interactive);
        setCheck($('licCheckSchema'), 'pending', 'SharePoint-Schema wird geprüft …', '');

        // 1) Bevorzugt License-API (kein iframe / Sites.Manage.All)
        try {
            var viaApi = await refreshSchemaViaLicenseApi({ interactive: interactive });
            if (viaApi) return;
        } catch (apiErr) {
            if (!interactive && isAuthInteractionError(apiErr)) {
                setCheck(
                    $('licCheckSchema'),
                    'warn',
                    'Schema noch nicht geprüft',
                    friendlySchemaAuthDetail()
                );
                return;
            }
            if (interactive && isAuthInteractionError(apiErr)) {
                setCheck(
                    $('licCheckSchema'),
                    'warn',
                    'Anmeldung für Schema-Prüfung nötig',
                    friendlySchemaAuthDetail()
                );
                return;
            }
            // API-Fehler: Fallback auf Browser-Graph unten
            console.warn('[admin-setup] License-API Schema:', apiErr);
        }

        var setup = window.ms365LicenseListSetup;
        if (!setup || typeof setup.checkSchemaStatus !== 'function') {
            setCheck(
                $('licCheckSchema'),
                interactive ? 'warn' : 'pending',
                'SharePoint-Schema',
                interactive
                    ? 'License-API Schema nicht verfügbar – bitte API prüfen oder „Liste / Spalten aktualisieren“.'
                    : 'Bitte mit MS365 anmelden, dann „Status prüfen“.'
            );
            return;
        }

        try {
            var st = await setup.checkSchemaStatus();
            var open = $('licOpenLink');
            if (open && st.webUrl) {
                open.href = st.webUrl;
                open.hidden = false;
                var openLabel = $('licOpenLinkLabel');
                if (openLabel) openLabel.textContent = 'Liste öffnen';
            }
            if (!st.listExists) {
                setCheck(
                    $('licCheckSchema'),
                    'warn',
                    'Liste fehlt noch',
                    '„Liste / Spalten aktualisieren“ einmal ausführen.'
                );
                return;
            }
            if (st.missing && st.missing.length) {
                setCheck(
                    $('licCheckSchema'),
                    'warn',
                    'Spalten unvollständig',
                    'Fehlt: ' + st.missing.join(', ') + ' → bitte aktualisieren.'
                );
                return;
            }
            setCheck(
                $('licCheckSchema'),
                'ok',
                'Liste & Spalten vollständig',
                'Alle benötigten Felder sind vorhanden (inkl. AdditionalDomains).'
            );
        } catch (e) {
            if (isAuthInteractionError(e)) {
                setCheck(
                    $('licCheckSchema'),
                    'warn',
                    interactive ? 'Anmeldung für Schema-Prüfung nötig' : 'Schema noch nicht geprüft',
                    friendlySchemaAuthDetail()
                );
                return;
            }
            setCheck(
                $('licCheckSchema'),
                'warn',
                'Schema-Prüfung fehlgeschlagen',
                (e && e.message) || String(e)
            );
        }
    }

    function acquireToken() {
        if (window.ms365LicenseApi && typeof window.ms365LicenseApi.acquireLicenseToken === 'function') {
            return window.ms365LicenseApi.acquireLicenseToken({ popup: true });
        }
        return Promise.reject(new Error('Bitte zuerst mit MS365 anmelden.'));
    }

    function wireApiTests() {
        var out = $('licApiResult');
        var btnMe = $('licBtnApiMe');
        var btnHealth = $('licBtnApiHealth');

        if (btnHealth) {
            btnHealth.addEventListener('click', function () {
                var base = apiBase();
                if (!out) return;
                if (!base) {
                    out.textContent = 'MS365_LICENSE_API.baseUrl ist nicht gesetzt.';
                    return;
                }
                out.textContent = 'GET /health …';
                var healthUrl = base.replace(/\/+$/, '') + '/health';
                // baseUrl endet auf /api/license → /api/license/health
                fetch(healthUrl)
                    .then(function (res) {
                        return res.text().then(function (t) {
                            var body = t;
                            try {
                                body = JSON.stringify(JSON.parse(t), null, 2);
                            } catch {
                                /* raw */
                            }
                            out.textContent = 'HTTP ' + res.status + '\n' + body;
                        });
                    })
                    .catch(function (e) {
                        out.textContent = 'Fehler: ' + ((e && e.message) || e);
                    });
            });
        }

        if (btnMe) {
            btnMe.addEventListener('click', function () {
                if (!out) return;
                out.textContent = 'Lade Token …';
                acquireToken()
                    .then(function (token) {
                        out.textContent = 'Rufe GET /license/me auf …';
                        return window.ms365LicenseApi.fetchLicenseMe(token);
                    })
                    .then(function (data) {
                        out.textContent = JSON.stringify(data, null, 2);
                    })
                    .catch(function (e) {
                        out.textContent = 'Fehler: ' + ((e && e.message) || e);
                        if (e && e.payload) {
                            out.textContent += '\n' + JSON.stringify(e.payload, null, 2);
                        }
                    });
            });
        }
    }

    function wireSchemaActions() {
        var btnRefresh = $('licBtnSchemaStatus');
        if (btnRefresh) {
            btnRefresh.addEventListener('click', function () {
                refreshSchemaStatus({ interactive: true }).then(function () { toast('Schema-Status aktualisiert.'); });
            });
        }

        document.addEventListener('ms365-license-list-setup-done', function () { refreshSchemaStatus({ interactive: false }); });
    }

    function goLicensesTab() {
        if (typeof window.ms365AdminSetTab === 'function') {
            window.ms365AdminSetTab('licenses');
        } else {
            location.hash = 'licenses';
        }
    }

    function init() {
        if (!$('adminPanelSetup')) return;

        var cfg = window.MS365_LICENSE_BACKEND || {};
        var siteInput = $('licSiteUrl');
        var listInput = $('licListName');
        if (siteInput && !siteInput.value) siteInput.value = cfg.siteWebUrl || '';
        if (listInput && !listInput.value) listInput.value = cfg.listDisplayName || '';

        setCheck(
            $('licCheckLicenses'),
            'off',
            'Freischalten unter Lizenzen',
            'Hier nur Schema und API – Schulen anlegen im Tab Lizenzen.'
        );
        refreshApiConfigUi();

        var goBtn = $('licBtnGoLicenses');
        if (goBtn) goBtn.addEventListener('click', goLicensesTab);

        setCheck(
            $('licCheckSchema'),
            'off',
            'SharePoint-Schema',
            'Noch nicht geprüft.'
        );

        wireApiTests();
        wireSchemaActions();

        setTimeout(function () {
            if (typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn()) {
                // Still – kein Popup/iframe-Zwang beim Tab-Öffnen
                refreshSchemaStatus({ interactive: false });
            }
        }, 600);
        window.addEventListener('ms365-auth-state-changed', function () {
            refreshApiConfigUi();
            if (typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn()) {
                refreshSchemaStatus({ interactive: false });
            } else {
                setCheck(
                    $('licCheckSchema'),
                    'off',
                    'SharePoint-Schema',
                    'Noch nicht geprüft.'
                );
            }
        });
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();
