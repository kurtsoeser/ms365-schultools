/**
 * Betreiber-Zugang: serverseitig über License-API (/admin/me).
 * Keine Betreiber-UPNs und keine Admin-PINs in der öffentlichen Config.
 */
(function (global) {
    'use strict';

    var CACHE_KEY = 'ms365-admin-operator-cache-v1';
    var USER_SESSION_KEY = 'ms365-access-granted-v1';
    var TTL_MS = 30 * 60 * 1000;

    /** @type {{ oid: string, upn: string, checkedAt: number } | null} */
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

    function cacheValidForAccount(account) {
        var cache = readCache();
        if (!cache || !cache.checkedAt) return false;
        if (Date.now() - Number(cache.checkedAt) > TTL_MS) return false;
        var oid = account && account.oid ? account.oid : '';
        if (!oid || String(cache.oid || '').toLowerCase() !== oid) return false;
        return true;
    }

    /**
     * Sync-Hinweis für Menü: nur wenn Cache zum aktuellen Konto passt.
     */
    function isCurrentUserOperator() {
        return cacheValidForAccount(currentAccountInfo());
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
        if (!force && cacheValidForAccount(account)) return true;

        var api = global.ms365LicenseApi;
        if (!api || typeof api.acquireLicenseToken !== 'function' || typeof api.fetchAdminMe !== 'function') {
            clearOperatorCache();
            return false;
        }
        var cfg = global.MS365_LICENSE_API || {};
        if (!String(cfg.baseUrl || '').trim()) {
            clearOperatorCache();
            return false;
        }

        try {
            var token = await api.acquireLicenseToken();
            var me = await api.fetchAdminMe(token);
            if (!me || me.operator !== true) {
                clearOperatorCache();
                return false;
            }
            var oid = String((me.user && me.user.oid) || account.oid || '')
                .trim()
                .toLowerCase();
            var upn = normalizeUpn((me.user && me.user.upn) || account.upn);
            writeCache({ oid: oid, upn: upn, checkedAt: Date.now() });
            try {
                sessionStorage.setItem(USER_SESSION_KEY, '1');
            } catch (e) {
                /* ignore */
            }
            return true;
        } catch (e) {
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

    function openAdminArea() {
        global.location.href = resolveAppRootHref('admin.html');
    }

    global.ms365OperatorAccess = {
        isCurrentUserOperator: isCurrentUserOperator,
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
})(typeof window !== 'undefined' ? window : globalThis);
