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

    async function refreshSchemaStatus() {
        var setup = window.ms365LicenseListSetup;
        if (!setup || typeof setup.checkSchemaStatus !== 'function') {
            setCheck(
                $('licCheckSchema'),
                'pending',
                'SharePoint-Schema',
                'Bitte mit MS365 anmelden, dann „Status prüfen“.'
            );
            return;
        }
        setCheck($('licCheckSchema'), 'pending', 'SharePoint-Schema wird geprüft …', '');
        try {
            var st = await setup.checkSchemaStatus();
            var open = $('licOpenLink');
            if (open && st.webUrl) {
                open.href = st.webUrl;
                open.hidden = false;
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
            setCheck(
                $('licCheckSchema'),
                'warn',
                'Schema-Prüfung fehlgeschlagen',
                (e && e.message) || String(e)
            );
        }
    }

    function acquireToken() {
        var scopes = ['https://graph.microsoft.com/User.Read'];
        if (typeof window.ms365AuthAcquireIdTokenPopup === 'function') {
            return window.ms365AuthAcquireIdTokenPopup(scopes);
        }
        if (typeof window.ms365AuthAcquireIdToken === 'function') {
            return window.ms365AuthAcquireIdToken(scopes);
        }
        if (typeof window.ms365AuthAcquireTokenPopup === 'function') {
            return window.ms365AuthAcquireTokenPopup(scopes);
        }
        if (typeof window.ms365AuthAcquireToken === 'function') {
            return window.ms365AuthAcquireToken(scopes);
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
                refreshSchemaStatus().then(function () {
                    toast('Schema-Status aktualisiert.');
                });
            });
        }

        document.addEventListener('ms365-license-list-setup-done', function () {
            refreshSchemaStatus();
        });
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
            'Schulen freischalten – nicht hier',
            'Erledigt im Tab „Lizenzen“. Manuelle SharePoint-Zeilen brauchst du im Alltag nicht mehr.'
        );
        refreshApiConfigUi();

        var goBtn = $('licBtnGoLicenses');
        if (goBtn) goBtn.addEventListener('click', goLicensesTab);

        wireApiTests();
        wireSchemaActions();

        // Schema erst prüfen, wenn User ggf. schon angemeldet ist
        setTimeout(function () {
            refreshSchemaStatus();
        }, 600);
        window.addEventListener('ms365-auth-state-changed', function () {
            refreshSchemaStatus();
            refreshApiConfigUi();
        });
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();
