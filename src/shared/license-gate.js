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
    var lastMode = '';

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
        lastMode = 'ok';
    }

    function setActions(buttons) {
        var wrap = $('ms365LicenseGateActions');
        if (!wrap) return;
        wrap.replaceChildren();
        (buttons || []).forEach(function (b) {
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = b.primary ? 'btn' : 'btn alt';
            btn.innerHTML = b.html || b.label;
            btn.addEventListener('click', b.onClick);
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
            } else {
                meta.innerHTML = '';
                meta.hidden = true;
            }
        }
        setActions(opts.actions || []);
        el.hidden = false;
        document.documentElement.setAttribute('data-ms365-license', opts.mode || 'blocked');
        lastMode = opts.mode || 'blocked';
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
            parts.push('<div><strong>Tenant-ID:</strong> <code>' + escapeHtml(result.tenantId) + '</code></div>');
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
        showOverlay({
            mode: 'denied',
            title: 'Kein Zugang',
            text:
                (result && result.message) ||
                'Dieser Mandant ist nicht für die MS365-Schultools freigeschaltet.',
            metaHtml: metaFromResult(result),
            actions: [
                {
                    primary: true,
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

    function applyAllowed(result) {
        hideOverlay();
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-license-changed', {
                    detail: { allowed: true, result: result || null }
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
                    detail: { allowed: false, result: result || null }
                })
            );
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
        window.addEventListener('ms365-auth-widget-ready', function () {
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

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
    } else {
        boot();
    }
})();
