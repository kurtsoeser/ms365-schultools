/**
 * Betreiber-Zugang (Admin) über MS365-UPN – ergänzt den Master-PIN.
 */
(function (global) {
    'use strict';

    var ADMIN_SESSION_KEY = 'ms365-admin-access-granted-v1';
    var USER_SESSION_KEY = 'ms365-access-granted-v1';

    function normalizeUpn(v) {
        return String(v == null ? '' : v)
            .trim()
            .toLowerCase();
    }

    function operatorUpnsFromConfig() {
        var cfg = global.MS365_ACCESS_CONFIG || {};
        var list = [];
        if (Array.isArray(cfg.operatorUpns)) {
            cfg.operatorUpns.forEach(function (u) {
                var n = normalizeUpn(u);
                if (n) list.push(n);
            });
        }
        if (typeof cfg.operatorUpn === 'string' && cfg.operatorUpn.trim()) {
            list.push(normalizeUpn(cfg.operatorUpn));
        }
        return list;
    }

    function isOperatorUpn(upn) {
        var needle = normalizeUpn(upn);
        if (!needle) return false;
        return operatorUpnsFromConfig().indexOf(needle) !== -1;
    }

    function currentAccountUpn() {
        try {
            if (typeof global.ms365AuthGetAccountInfo === 'function') {
                var info = global.ms365AuthGetAccountInfo();
                if (info && info.upn) return normalizeUpn(info.upn);
                if (info && info.username) return normalizeUpn(info.username);
            }
        } catch (e) {
            /* ignore */
        }
        try {
            if (typeof global.ms365AuthGetUserPrincipalName === 'function') {
                return normalizeUpn(global.ms365AuthGetUserPrincipalName());
            }
        } catch (e2) {
            /* ignore */
        }
        return '';
    }

    function isCurrentUserOperator() {
        return isOperatorUpn(currentAccountUpn());
    }

    function grantAdminSession() {
        try {
            sessionStorage.setItem(ADMIN_SESSION_KEY, '1');
            sessionStorage.setItem(USER_SESSION_KEY, '1');
        } catch (e) {
            /* ignore */
        }
    }

    function grantAdminSessionIfOperator() {
        if (!isCurrentUserOperator()) return false;
        grantAdminSession();
        return true;
    }

    /**
     * admin.html / license-setup aus beliebigem Pfad.
     * @param {string} file
     */
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
        if (!grantAdminSessionIfOperator()) {
            global.alert('Admin nur für hinterlegte Betreiber-Konten.');
            return;
        }
        global.location.href = resolveAppRootHref('admin.html');
    }

    global.ms365OperatorAccess = {
        isOperatorUpn: isOperatorUpn,
        isCurrentUserOperator: isCurrentUserOperator,
        grantAdminSessionIfOperator: grantAdminSessionIfOperator,
        grantAdminSession: grantAdminSession,
        resolveAppRootHref: resolveAppRootHref,
        openAdminArea: openAdminArea,
        currentAccountUpn: currentAccountUpn
    };
})(typeof window !== 'undefined' ? window : globalThis);
