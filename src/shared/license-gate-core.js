/**
 * Lizenz-Gate Kernlogik (ohne DOM).
 *
 * Overlay = UX. App-Only-Aktionen prüfen die Lizenz serverseitig.
 * Entitlements steuern Modul-Sichtbarkeit (Dashboard + Tool-Seiten).
 */
(function (global) {
    'use strict';

    var CACHE_KEY = 'ms365-license-me-v2';
    /** Kurz halten – Cache ist nur Komfort, nicht Vertrauensgrenze. */
    var TTL_MS = 10 * 60 * 1000;

    /** HTML-Dateiname (ohne .html) → data-tool-id, wenn abweichend */
    var PATH_ALIASES = {
        jahrgangsgruppen: 'jahrgang',
        'schueler-sammelgruppe': 'slg-schueler',
        'lehrer-sammelgruppe': 'slg-lehrer'
    };

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

    /**
     * Tool-ID aus URL, z. B. /tools/kursteams.html → kursteams
     * @param {string} [pathname]
     * @returns {string|null}
     */
    function toolIdFromPathname(pathname) {
        var p = String(pathname || '')
            .replace(/\\/g, '/')
            .split('?')[0]
            .split('#')[0];
        var m = p.match(/\/tools\/([^/]+?)\.html$/i);
        if (!m) return null;
        var base = String(m[1] || '')
            .trim()
            .toLowerCase();
        if (!base) return null;
        if (PATH_ALIASES[base]) return PATH_ALIASES[base];
        return base;
    }

    /**
     * @param {unknown} entitlements
     * @returns {{ packages: string[], features: string[], denied: string[] }|null}
     */
    function normalizeEntitlements(entitlements) {
        if (!entitlements || typeof entitlements !== 'object') return null;
        var features = Array.isArray(entitlements.features)
            ? entitlements.features.map(function (x) {
                  return String(x || '').trim();
              }).filter(Boolean)
            : null;
        if (!features) return null;
        return {
            packages: Array.isArray(entitlements.packages)
                ? entitlements.packages.map(function (x) {
                      return String(x || '').trim();
                  }).filter(Boolean)
                : [],
            features: features,
            denied: Array.isArray(entitlements.denied)
                ? entitlements.denied.map(function (x) {
                      return String(x || '').trim();
                  }).filter(Boolean)
                : []
        };
    }

    /**
     * Modul freigeschaltet?
     * Ohne License-API / ohne Cache / ohne Entitlements → erlaubt (Dev + Legacy).
     * Unbekannte Tool-IDs (nicht in features und nicht in denied) → erlaubt.
     * @param {string} featureId
     * @param {object|null} [cacheOrResult]
     */
    function hasFeature(featureId, cacheOrResult) {
        var id = String(featureId || '').trim();
        if (!id) return true;
        if (!isLicenseApiConfigured()) return true;
        var data = cacheOrResult != null ? cacheOrResult : readCache();
        /* Noch keine Prüfung → nicht vorschnell ausblenden */
        if (!data) return true;
        if (data.allowed !== true) return false;
        var ent = normalizeEntitlements(data.entitlements);
        if (!ent) return true;
        if (ent.features.indexOf(id) >= 0) return true;
        if (ent.denied.indexOf(id) >= 0) return false;
        /* Tool nicht im Katalog → nicht sperren */
        return true;
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
            if (typeof data.allowed !== 'boolean') return null;
            var at = Number(data.checkedAt || 0);
            if (!at || Date.now() - at > TTL_MS) return null;
            if (!String(data.accountKey || '').trim()) return null;
            if (!String(data.tenantId || '').trim()) return null;
            return data;
        } catch (e) {
            return null;
        }
    }

    function writeCache(result, accountKey, storage) {
        var store = storage || (typeof sessionStorage !== 'undefined' ? sessionStorage : null);
        if (!store) return;
        try {
            var tid = result && result.tenantId ? String(result.tenantId).trim() : '';
            var key = String(accountKey || '').trim();
            if (!tid || !key) return;
            store.setItem(
                CACHE_KEY,
                JSON.stringify({
                    allowed: !!(result && result.allowed),
                    reason: result && result.reason ? String(result.reason) : '',
                    message: result && result.message ? String(result.message) : '',
                    tenantId: tid,
                    license: (result && result.license) || null,
                    entitlements: normalizeEntitlements(result && result.entitlements),
                    user: (result && result.user) || null,
                    isDesignatedSchoolIt: !!(result && result.isDesignatedSchoolIt),
                    accountKey: key,
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
                    var oid = String(info.oid || '').trim();
                    var tid = String(info.tenantId || '').trim();
                    if (oid && tid) return oid + '|' + tid;
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
        PATH_ALIASES: PATH_ALIASES,
        isExemptPath: isExemptPath,
        isLicenseApiConfigured: isLicenseApiConfigured,
        toolIdFromPathname: toolIdFromPathname,
        normalizeEntitlements: normalizeEntitlements,
        hasFeature: hasFeature,
        readCache: readCache,
        writeCache: writeCache,
        clearCache: clearCache,
        accountKeyFromAuth: accountKeyFromAuth,
        cacheMatchesAccount: cacheMatchesAccount
    };
})(typeof window !== 'undefined' ? window : globalThis);
