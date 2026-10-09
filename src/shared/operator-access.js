/**
 * Betreiber-Zugang: serverseitig über License-API (/admin/me).
 * Keine Betreiber-UPNs und keine Admin-PINs in der öffentlichen Config.
 */
(function (global) {
    'use strict';

    var CACHE_KEY = 'ms365-admin-operator-cache-v1';
    var USER_SESSION_KEY = 'ms365-access-granted-v1';
    var TTL_MS = 30 * 60 * 1000;

    /** @type {{ oid: string, upn: string, checkedAt: number, operator?: boolean|null, pending?: boolean } | null} */
    var memoryCache = null;

    function normalizeUpn(v) {
        return String(v == null ? '' : v)
            .trim()
            .toLowerCase();
    }

    function currentAccountInfo() {
        try {
            if (typeof global.ms365AuthGetAccountInfo === 'function') {
                var info = global.ms365AuthGetAccountInfo();
                if (info) {
                    return {
                        oid: String(info.oid || '')
                            .trim()
                            .toLowerCase(),
                        upn: normalizeUpn(info.upn || info.username || '')
                    };
                }
            }
        } catch (e) {
            /* ignore */
        }
        return { oid: '', upn: normalizeUpn(currentAccountUpn()) };
    }

    function currentAccountUpn() {
        try {
            if (typeof global.ms365AuthGetUserPrincipalName === 'function') {
                return normalizeUpn(global.ms365AuthGetUserPrincipalName());
            }
        } catch (e) {
            /* ignore */
        }
        return '';
    }

    function readCache() {
        if (memoryCache) return memoryCache;
        try {
            var raw = sessionStorage.getItem(CACHE_KEY);
            if (!raw) return null;
            var data = JSON.parse(raw);
            if (!data || typeof data !== 'object') return null;
            memoryCache = data;
            return data;
        } catch (e) {
            return null;
        }
    }

    function writeCache(entry) {
        memoryCache = entry;
        try {
            if (entry) sessionStorage.setItem(CACHE_KEY, JSON.stringify(entry));
            else sessionStorage.removeItem(CACHE_KEY);
        } catch (e) {
            /* ignore */
        }
    }

    function clearOperatorCache() {
        writeCache(null);
        try {
            sessionStorage.removeItem('ms365-admin-access-granted-v1');
            sessionStorage.removeItem(USER_SESSION_KEY);
        } catch (e) {
            /* ignore */
        }
    }

    function cacheMatchesAccount(cache, account) {
        if (!cache || !cache.checkedAt) return false;
        var oid = account && account.oid ? account.oid : '';
        var upn = account && account.upn ? account.upn : '';
        if (oid && String(cache.oid || '').toLowerCase() === oid) return true;
        if (upn && normalizeUpn(cache.upn) === upn) return true;
        return false;
    }

    function cacheFreshForAccount(account) {
        var cache = readCache();
        if (!cacheMatchesAccount(cache, account)) return false;
        return Date.now() - Number(cache.checkedAt) <= TTL_MS;
    }

    /**
     * Sync-Hinweis für Menü: nur wenn Cache zum aktuellen Konto passt.
     */
    function isCurrentUserOperator() {
        var account = currentAccountInfo();
        if (!cacheFreshForAccount(account)) return false;
        var cache = readCache();
        return !!(cache && cache.operator === true);
    }

    /**
     * Admin-Menüpunkt: sichtbar bei bestätigtem Betreiber ODER solange der
     * stille Check noch nicht klar „kein Betreiber“ geliefert hat (z. B. License.Access fehlt).
     */
    function shouldShowAdminMenuLink() {
        var account = currentAccountInfo();
        if (!account.oid && !account.upn) return false;
        var cfg = global.MS365_LICENSE_API || {};
        if (cfg.skipOperatorCheck === true) return false;
        var cache = readCache();
        if (!cacheMatchesAccount(cache, account)) return false;
        if (cache.operator === true) return true;
        if (cache.operator === false) return false;
        /* pending / unklar nach Silent-Fehler */
        return cache.pending === true || cache.operator == null;
    }

    /**
     * Betreiber über License-API prüfen (LICENSE_OPERATOR_* in Azure).
     * @param {{ force?: boolean }} [opts]
     * @returns {Promise<boolean>}
     */
    async function refreshOperatorStatus(opts) {
        var force = !!(opts && opts.force);
        var account = currentAccountInfo();
        if (!account.oid && !account.upn) {
            clearOperatorCache();
            return false;
        }
        var cfg = global.MS365_LICENSE_API || {};
        if (cfg.skipOperatorCheck === true) {
            clearOperatorCache();
            return false;
        }

        if (!force && cacheFreshForAccount(account)) {
            var cached = readCache();
            if (cached && cached.operator === true) return true;
            if (cached && cached.operator === false) return false;
            /* pending: erneut versuchen, sobald Token da sein könnte */
        }

        var api = global.ms365LicenseApi;
        if (!api || typeof api.acquireLicenseToken !== 'function' || typeof api.fetchAdminMe !== 'function') {
            if (!force) {
                writeCache({
                    oid: account.oid,
                    upn: account.upn,
                    checkedAt: Date.now(),
                    operator: null,
                    pending: true
                });
            }
            return false;
        }
        if (!String(cfg.baseUrl || '').trim()) {
            clearOperatorCache();
            return false;
        }

        try {
            /* Menü: silentOnly. Admin-Seite / Klick (force): Popup bei Bedarf. */
            var token = await api.acquireLicenseToken(force ? { popup: true } : { silentOnly: true });
            var me = await api.fetchAdminMe(token);
            var oid = String((me && me.user && me.user.oid) || account.oid || '')
                .trim()
                .toLowerCase();
            var upn = normalizeUpn((me && me.user && me.user.upn) || account.upn);
            var isOp = !!(me && me.operator === true);
            writeCache({ oid: oid, upn: upn, checkedAt: Date.now(), operator: isOp, pending: false });
            if (!isOp) {
                try {
                    sessionStorage.removeItem(USER_SESSION_KEY);
                } catch (e) {
                    /* ignore */
                }
                return false;
            }
            try {
                sessionStorage.setItem(USER_SESSION_KEY, '1');
            } catch (e) {
                /* ignore */
            }
            return true;
        } catch (e) {
            /* Positiven Cache behalten */
            if (!force && cacheFreshForAccount(account)) {
                var c = readCache();
                if (c && c.operator === true) return true;
            }
            /*
             * Silent-Fehler ≠ „kein Betreiber“. Cache nicht löschen, sonst verschwindet
             * der Admin-Menüpunkt (License.Access oft erst nach License-Gate).
             */
            if (!force) {
                var prev = readCache();
                if (prev && cacheMatchesAccount(prev, account) && prev.operator === true) {
                    return true;
                }
                writeCache({
                    oid: account.oid,
                    upn: account.upn,
                    checkedAt: Date.now(),
                    operator: null,
                    pending: true
                });
                return false;
            }
            clearOperatorCache();
            return false;
        }
    }

    function resolveAppRootHref(file) {
        var name = String(file || 'admin.html').replace(/^[./]+/, '');
        try {
            var p = String(global.location.pathname || '/').split('?')[0].split('#')[0];
            var lower = p.toLowerCase();
            var idx = lower.indexOf('/tools/');
            if (idx !== -1) {
                return p.slice(0, idx) + '/' + name;
            }
            var slash = p.lastIndexOf('/');
            var dir = slash >= 0 ? p.slice(0, slash + 1) : '/';
            return dir + name;
        } catch (e) {
            return name;
        }
    }

    async function openAdminArea() {
        try {
            var ok = await refreshOperatorStatus({ force: true });
            if (!ok) {
                if (typeof global.ms365ShowToast === 'function') {
                    global.ms365ShowToast(
                        'Admin-Bereich nur für das Betreiber-Konto (License-API / LICENSE_OPERATOR_*).',
                        { kind: 'warning', title: 'Kein Betreiber-Zugang' }
                    );
                } else if (typeof global.ms365ToastOrAlert === 'function') {
                    global.ms365ToastOrAlert('Admin-Bereich nur für das Betreiber-Konto.');
                } else {
                    global.alert('Admin-Bereich nur für das Betreiber-Konto.');
                }
                notifyAuthWidget();
                return;
            }
        } catch (e) {
            /* trotzdem versuchen zu öffnen – Boot prüft erneut */
        }
        global.location.href = resolveAppRootHref('admin.html');
    }

    global.ms365OperatorAccess = {
        isCurrentUserOperator: isCurrentUserOperator,
        shouldShowAdminMenuLink: shouldShowAdminMenuLink,
        refreshOperatorStatus: refreshOperatorStatus,
        clearOperatorCache: clearOperatorCache,
        resolveAppRootHref: resolveAppRootHref,
        openAdminArea: openAdminArea,
        currentAccountUpn: currentAccountUpn,
        /** @deprecated nur noch Alias – Admin-Session ist der API-Cache */
        grantAdminSessionIfOperator: function () {
            return isCurrentUserOperator();
        },
        grantAdminSession: function () {
            /* no-op: Session entsteht nur nach refreshOperatorStatus */
        },
        isOperatorUpn: function () {
            return isCurrentUserOperator();
        }
    };

    /* Nach deferred inject: Auth-Menü neu bewerten (Admin-Link) */
    function notifyAuthWidget() {
        try {
            if (typeof global.ms365AuthRefreshWidget === 'function') {
                global.ms365AuthRefreshWidget({ silent: true });
            }
        } catch (e) {
            /* ignore */
        }
    }
    if (typeof document !== 'undefined') {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', function () {
                setTimeout(notifyAuthWidget, 0);
            });
        } else {
            setTimeout(notifyAuthWidget, 0);
        }
        try {
            /* License-Gate holt License.Access oft erst nach dem ersten Menü-Check */
            global.addEventListener('ms365-license-changed', function () {
                /* Auch bei allowed:false erneut prüfen – Betreiber-Tenant braucht keine Schul-Lizenz,
                 * aber License.Access kann erst jetzt im Token-Cache liegen. */
                refreshOperatorStatus({ force: false }).then(function () {
                    notifyAuthWidget();
                });
            });
        } catch (e) {
            /* ignore */
        }
    }
})(typeof window !== 'undefined' ? window : globalThis);
