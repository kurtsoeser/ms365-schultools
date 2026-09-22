/**
 * Phase 3 – Lizenz-Gate Kernlogik (ohne DOM).
 */
(function (global) {
    'use strict';

    var CACHE_KEY = 'ms365-license-me-v1';
    var TTL_MS = 30 * 60 * 1000;

    function isExemptPath(pathname) {
        var p = String(pathname || '');
        return (
            /\/welcome\.html(?:\?|#|$)/i.test(p) ||
            /\/hilfe\.html(?:\?|#|$)/i.test(p) ||
            /\/ms365-schooltool\.html(?:\?|#|$)/i.test(p) ||
            /\/admin\.html(?:\?|#|$)/i.test(p) ||
            /\/tools\/license-backend-setup\.html(?:\?|#|$)/i.test(p)
        );
    }

    function isLicenseApiConfigured() {
        var cfg = global.MS365_LICENSE_API || {};
        return !!(cfg && String(cfg.baseUrl || '').trim());
    }

    function readCache(storage) {
        var store = storage || (typeof sessionStorage !== 'undefined' ? sessionStorage : null);
        if (!store) return null;
        try {
            var raw = store.getItem(CACHE_KEY);
            if (!raw) return null;
            var data = JSON.parse(raw);
            if (!data || typeof data !== 'object') return null;
            var at = Number(data.checkedAt || 0);
            if (!at || Date.now() - at > TTL_MS) return null;
            return data;
        } catch (e) {
            return null;
        }
    }

    function writeCache(result, accountKey, storage) {
        var store = storage || (typeof sessionStorage !== 'undefined' ? sessionStorage : null);
        if (!store) return;
        try {
            store.setItem(
                CACHE_KEY,
                JSON.stringify({
                    allowed: !!(result && result.allowed),
                    reason: result && result.reason ? String(result.reason) : '',
                    message: result && result.message ? String(result.message) : '',
                    tenantId: result && result.tenantId ? String(result.tenantId) : '',
                    license: (result && result.license) || null,
                    user: (result && result.user) || null,
                    accountKey: String(accountKey || ''),
                    checkedAt: Date.now()
                })
            );
        } catch (e) {
            /* ignore */
        }
    }

    function clearCache(storage) {
        var store = storage || (typeof sessionStorage !== 'undefined' ? sessionStorage : null);
        if (!store) return;
        try {
            store.removeItem(CACHE_KEY);
        } catch (e) {
            /* ignore */
        }
    }

    function accountKeyFromAuth() {
        try {
            if (typeof global.ms365AuthGetAccountInfo === 'function') {
                var info = global.ms365AuthGetAccountInfo();
                if (info) {
                    return String(info.oid || info.upn || info.username || info.tenantId || '').trim();
                }
            }
        } catch (e) {
            /* ignore */
        }
        try {
            if (typeof global.ms365AuthGetAccountLabel === 'function') {
                return String(global.ms365AuthGetAccountLabel() || '').trim();
            }
        } catch (e2) {
            /* ignore */
        }
        return '';
    }

    function cacheMatchesAccount(cache, accountKey) {
        if (!cache) return false;
        var a = String(accountKey || '').trim().toLowerCase();
        var b = String(cache.accountKey || '').trim().toLowerCase();
        if (!a || !b) return false;
        return a === b;
    }

    global.ms365LicenseGateCore = {
        CACHE_KEY: CACHE_KEY,
        TTL_MS: TTL_MS,
        isExemptPath: isExemptPath,
        isLicenseApiConfigured: isLicenseApiConfigured,
        readCache: readCache,
        writeCache: writeCache,
        clearCache: clearCache,
        accountKeyFromAuth: accountKeyFromAuth,
        cacheMatchesAccount: cacheMatchesAccount
    };
})(typeof window !== 'undefined' ? window : globalThis);
