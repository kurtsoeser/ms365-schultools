/**
 * Phase 3 – Lizenz-Gate nach MSAL-Login (nur UX-Overlay).
 * Ohne MS365_LICENSE_API.baseUrl: kein Gate (lokale Dev ohne Backend).
 * App-Only-Aktionen (Kursteams-Backend) prüfen die Lizenz serverseitig.
 */
(function () {
    'use strict';

    var core = window.ms365LicenseGateCore;
    if (!core) return;
    if (core.isExemptPath(location.pathname)) return;

    var OVERLAY_ID = 'ms365LicenseGateOverlay';
    var checking = false;
    var _lastMode = '';

    function $(id) {
        return document.getElementById(id);
    }

    function ensureOverlay() {
        var el = $(OVERLAY_ID);
        if (el) return el;
        el = document.createElement('div');
        el.id = OVERLAY_ID;
        el.className = 'ms365-license-gate';
        el.setAttribute('role', 'dialog');
        el.setAttribute('aria-modal', 'true');
        el.hidden = true;
        el.innerHTML =
            '<div class="ms365-license-gate__card">' +
            '  <div class="ms365-license-gate__icon" aria-hidden="true"><i class="bi bi-key"></i></div>' +
            '  <h2 class="ms365-license-gate__title" id="ms365LicenseGateTitle">Lizenzprüfung</h2>' +
            '  <p class="ms365-license-gate__text" id="ms365LicenseGateText"></p>' +
            '  <div class="ms365-license-gate__meta" id="ms365LicenseGateMeta" hidden></div>' +
            '  <div class="ms365-license-gate__actions" id="ms365LicenseGateActions"></div>' +
            '</div>';
        document.body.appendChild(el);
        return el;
    }

    function hideOverlay() {
        var el = $(OVERLAY_ID);
        if (el) el.hidden = true;
        document.documentElement.removeAttribute('data-ms365-license');
        _lastMode = 'ok';
    }

    function setActions(buttons) {
        var wrap = $('ms365LicenseGateActions');
        if (!wrap) return;
        wrap.replaceChildren();
        (buttons || []).forEach(function (b) {
            if (b && b.href) {
                var a = document.createElement('a');
                a.className = b.primary ? 'btn' : 'btn alt';
                a.href = String(b.href);
                a.target = b.target || '_blank';
                a.rel = b.rel || 'noopener noreferrer';
                a.innerHTML = b.html || b.label || '';
                wrap.appendChild(a);
                return;
            }
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = b.primary ? 'btn' : 'btn alt';
            btn.innerHTML = b.html || b.label;
            if (typeof b.onClick === 'function') btn.addEventListener('click', b.onClick);
            wrap.appendChild(btn);
        });
    }

    function showOverlay(opts) {
        var el = ensureOverlay();
        var title = $('ms365LicenseGateTitle');
        var text = $('ms365LicenseGateText');
        var meta = $('ms365LicenseGateMeta');
        if (title) title.textContent = opts.title || 'Lizenzprüfung';
        if (text) text.textContent = opts.text || '';
        if (meta) {
            if (opts.metaHtml) {
                meta.innerHTML = opts.metaHtml;
                meta.hidden = false;
                var copyBtn = document.getElementById('ms365LicenseGateCopyTenant');
                if (copyBtn) {
                    copyBtn.addEventListener('click', function () {
                        var code = document.getElementById('ms365LicenseGateTenantId');
                        var tid = code ? String(code.textContent || '').trim() : '';
                        if (!tid) return;
                        var done = function () {
                            if (typeof window.ms365ToastOrAlert === 'function') {
                                window.ms365ToastOrAlert('Tenant-ID kopiert');
                            } else {
                                window.alert('Tenant-ID kopiert');
                            }
                        };
                        if (navigator.clipboard && navigator.clipboard.writeText) {
                            navigator.clipboard.writeText(tid).then(done).catch(function () {
                                window.prompt('Tenant-ID kopieren:', tid);
                            });
                        } else {
                            window.prompt('Tenant-ID kopieren:', tid);
                        }
                    });
                }
            } else {
                meta.innerHTML = '';
                meta.hidden = true;
            }
        }
        setActions(opts.actions || []);
        el.hidden = false;
        document.documentElement.setAttribute('data-ms365-license', opts.mode || 'blocked');
        _lastMode = opts.mode || 'blocked';
    }

    function metaFromResult(result) {
        var lic = (result && result.license) || {};
        var parts = [];
        if (lic.schoolName) parts.push('<div><strong>Schule:</strong> ' + escapeHtml(lic.schoolName) + '</div>');
        if (lic.status) parts.push('<div><strong>Status:</strong> ' + escapeHtml(lic.status) + '</div>');
        if (lic.validUntil) parts.push('<div><strong>Gültig bis:</strong> ' + escapeHtml(lic.validUntil) + '</div>');
        var domains = Array.isArray(lic.domains)
            ? lic.domains
            : lic.primaryDomain
              ? [lic.primaryDomain]
              : [];
        if (domains.length) {
            parts.push(
                '<div><strong>Domains:</strong> ' +
                    escapeHtml(domains.join(', ')) +
                    '</div>'
            );
        }
        if (result && result.tenantId) {
            parts.push(
                '<div class="ms365-license-gate__tenant">' +
                    '<strong>Tenant-ID:</strong> <code id="ms365LicenseGateTenantId">' +
                    escapeHtml(result.tenantId) +
                    '</code> ' +
                    '<button type="button" class="btn alt ms365-license-gate__copy" id="ms365LicenseGateCopyTenant" title="Tenant-ID kopieren">' +
                    '<i class="bi bi-clipboard"></i> Kopieren</button>' +
                    '</div>' +
                    '<div class="ms365-license-gate__hint">Bitte diese Tenant-ID an den Betreiber weitergeben, damit die Schule freigeschaltet werden kann.</div>'
            );
        }
        if (lic.contactEmail) {
            parts.push(
                '<div><strong>Kontakt:</strong> <a href="mailto:' +
                    escapeAttr(lic.contactEmail) +
                    '">' +
                    escapeHtml(lic.contactEmail) +
                    '</a></div>'
            );
        }
        return parts.length ? parts.join('') : '';
    }

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function escapeAttr(s) {
        return escapeHtml(s).replace(/'/g, '&#39;');
    }

    function showNeedLogin() {
        showOverlay({
            mode: 'login',
            title: 'Anmeldung erforderlich',
            text: 'Melden Sie sich mit dem Microsoft-365-Konto Ihrer Schule an, um die Werkzeuge zu nutzen.',
            actions: [
                {
                    primary: true,
                    html: '<i class="bi bi-box-arrow-in-right"></i> Mit Microsoft anmelden',
                    onClick: function () {
                        if (typeof window.ms365AuthLogin === 'function') {
                            window.ms365AuthLogin().catch(function () {});
                        }
                    }
                }
            ]
        });
    }

    function showDenied(result) {
        var reason = result && result.reason ? String(result.reason) : '';
        var deniedText =
            (result && result.message) ||
            'Ihre Schule ist für MS365-Schul-Tools noch nicht freigeschaltet. Bitte informieren Sie Ihren IT-Admin oder stellen Sie hier eine Anfrage.';
        if (reason === 'not_registered') {
            deniedText =
                'Ihre Schule ist noch nicht freigeschaltet. Bitte informieren Sie Ihren IT-Admin – oder stellen Sie eine Anfrage über die Homepage. Die Tenant-ID unten hilft bei der Zuordnung.';
        }
        showOverlay({
            mode: 'denied',
            title: 'Schule noch nicht freigeschaltet',
            text: deniedText,
            metaHtml: metaFromResult(result),
            actions: [
                {
                    primary: true,
                    html: '<i class="bi bi-envelope"></i> Anfrage stellen',
                    href: 'https://ms365.schule/#kontakt',
                    target: '_blank',
                    rel: 'noopener noreferrer'
                },
                {
                    html: '<i class="bi bi-arrow-clockwise"></i> Erneut prüfen',
                    onClick: function () {
                        core.clearCache();
                        runCheck({ force: true });
                    }
                },
                {
                    html: '<i class="bi bi-people"></i> Anderes Konto',
                    onClick: function () {
                        core.clearCache();
                        if (typeof window.ms365AuthSwitchAccount === 'function') {
                            window.ms365AuthSwitchAccount().catch(function () {});
                        } else if (typeof window.ms365AuthLogin === 'function') {
                            window.ms365AuthLogin().catch(function () {});
                        }
                    }
                }
            ]
        });
    }

    function showChecking() {
        showOverlay({
            mode: 'checking',
            title: 'Lizenz wird geprüft …',
            text: 'Einen Moment bitte.',
            actions: []
        });
    }

    function showError(message) {
        showOverlay({
            mode: 'error',
            title: 'Lizenzprüfung fehlgeschlagen',
            text: message || 'Die License-API ist nicht erreichbar.',
            actions: [
                {
                    primary: true,
                    html: '<i class="bi bi-arrow-clockwise"></i> Erneut versuchen',
                    onClick: function () {
                        runCheck({ force: true });
                    }
                },
                {
                    html: '<i class="bi bi-box-arrow-in-right"></i> Neu anmelden',
                    onClick: function () {
                        if (typeof window.ms365AuthLogin === 'function') {
                            window.ms365AuthLogin().catch(function () {});
                        }
                    }
                }
            ]
        });
    }

    function acquireTokenForLicense() {
        if (window.ms365LicenseApi && typeof window.ms365LicenseApi.acquireLicenseToken === 'function') {
            return window.ms365LicenseApi.acquireLicenseToken();
        }
        return Promise.reject(new Error('MSAL / License-API-Client nicht verfügbar.'));
    }

    function showFeatureDenied(result, toolId) {
        showOverlay({
            mode: 'feature-denied',
            title: 'Modul nicht freigeschaltet',
            text:
                'Dieses Werkzeug („' +
                String(toolId || '') +
                '“) ist für Ihre Schule derzeit nicht in der Lizenz enthalten. Bitte die Schul-IT bzw. den Betreiber kontaktieren.',
            metaHtml: metaFromResult(result),
            actions: [
                {
                    primary: true,
                    html: '<i class="bi bi-house"></i> Zum Dashboard',
                    onClick: function () {
                        try {
                            var p = String(location.pathname || '').replace(/\\/g, '/');
                            var i = p.toLowerCase().indexOf('/tools/');
                            location.href = i >= 0 ? p.slice(0, i) + '/index.html' : 'index.html';
                        } catch (e) {
                            location.href = 'index.html';
                        }
                    }
                },
                {
                    html: '<i class="bi bi-arrow-clockwise"></i> Erneut prüfen',
                    onClick: function () {
                        core.clearCache();
                        runCheck({ force: true });
                    }
                }
            ]
        });
    }

    function enforceToolFeature(result) {
        var toolId = core.toolIdFromPathname(location.pathname);
        if (!toolId) return true;
        if (core.hasFeature(toolId, result)) return true;
        showFeatureDenied(result || {}, toolId);
        return false;
    }

    function applyAllowed(result) {
        if (!enforceToolFeature(result)) {
            try {
                window.dispatchEvent(
                    new CustomEvent('ms365-license-changed', {
                        detail: {
                            allowed: true,
                            featureAllowed: false,
                            result: result || null
                        }
                    })
                );
            } catch (e) {
                /* ignore */
            }
            return;
        }
        hideOverlay();
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-license-changed', {
                    detail: {
                        allowed: true,
                        featureAllowed: true,
                        result: result || null
                    }
                })
            );
        } catch (e) {
            /* ignore */
        }
    }

    function applyDenied(result) {
        showDenied(result || {});
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-license-changed', {
                    detail: { allowed: false, featureAllowed: false, result: result || null }
                })
            );
        } catch (e) {
            /* ignore */
        }
    }

    async function waitForAuthReadyBeforeNeedLogin() {
        var waitMs = function (ms) {
            return new Promise(function (resolve) {
                setTimeout(resolve, ms);
            });
        };
        try {
            if (typeof window.ms365AuthEnsureInitialized === 'function') {
                await Promise.race([Promise.resolve(window.ms365AuthEnsureInitialized()), waitMs(4000)]);
            }
        } catch (e) {
            /* ignore */
        }
        if (typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn()) {
            return;
        }
        try {
            if (
                typeof window.ms365AuthEnsureBootSettled === 'function' &&
                !(typeof window.ms365AuthIsBootSettled === 'function' && window.ms365AuthIsBootSettled())
            ) {
                await Promise.race([Promise.resolve(window.ms365AuthEnsureBootSettled()), waitMs(4000)]);
            }
        } catch (e) {
            /* ignore */
        }
    }

    async function runCheck(opts) {
        if (!core.isLicenseApiConfigured()) {
            hideOverlay();
            return;
        }
        if (checking) return;
        checking = true;
        try {
            var loggedIn =
                typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
            if (!loggedIn) {
                await waitForAuthReadyBeforeNeedLogin();
                loggedIn =
                    typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
            }
            if (!loggedIn) {
                showNeedLogin();
                return;
            }

            var accountKey = core.accountKeyFromAuth();
            if (!(opts && opts.force)) {
                var cached = core.readCache();
                if (cached && core.cacheMatchesAccount(cached, accountKey)) {
                    if (cached.allowed) applyAllowed(cached);
                    else applyDenied(cached);
                    return;
                }
            }

            showChecking();
            var token = await acquireTokenForLicense();
            if (!window.ms365LicenseApi || typeof window.ms365LicenseApi.fetchLicenseMe !== 'function') {
                showError('License-API-Client fehlt.');
                return;
            }
            var result = await window.ms365LicenseApi.fetchLicenseMe(token);
            core.writeCache(result, accountKey);
            if (result && result.allowed) applyAllowed(result);
            else applyDenied(result);
        } catch (e) {
            showError((e && e.message) || String(e));
        } finally {
            checking = false;
        }
    }

    function scheduleCheck(force) {
        setTimeout(function () {
            runCheck({ force: !!force });
        }, 50);
    }

    function onAuthChanged() {
        if (!core.isLicenseApiConfigured()) return;
        var loggedIn =
            typeof window.ms365AuthIsLoggedIn === 'function' && window.ms365AuthIsLoggedIn();
        if (!loggedIn) {
            if (
                typeof window.ms365AuthIsBootSettled === 'function' &&
                !window.ms365AuthIsBootSettled()
            ) {
                return;
            }
            core.clearCache();
            showNeedLogin();
            return;
        }
        scheduleCheck(false);
    }

    function boot() {
        if (!core.isLicenseApiConfigured()) {
            // Ohne baseUrl kein Gate (z. B. bis Azure deployt ist).
            return;
        }
        ensureOverlay();

        window.addEventListener('ms365-auth-state-changed', onAuthChanged);
        window.addEventListener('ms365-auth-boot-settled', function () {
            scheduleCheck(false);
        });

        if (typeof window.ms365AuthEnsureInitialized === 'function') {
            Promise.resolve(window.ms365AuthEnsureInitialized())
                .then(function () {
                    scheduleCheck(false);
                })
                .catch(function () {
                    scheduleCheck(false);
                });
        } else {
            scheduleCheck(false);
        }
    }

    window.ms365LicenseGateRecheck = function () {
        core.clearCache();
        return runCheck({ force: true });
    };
    window.ms365LicenseGateClear = function () {
        core.clearCache();
    };
    window.ms365LicenseHasFeature = function (featureId) {
        return core.hasFeature(featureId);
    };
    window.ms365LicenseGetEntitlements = function () {
        var cached = core.readCache();
        return (cached && cached.entitlements) || null;
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
    } else {
        boot();
    }
})();
