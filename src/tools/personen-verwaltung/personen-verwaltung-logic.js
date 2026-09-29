/**
 * Helfer für Personen-Verwaltung (Analyse 02 Phase B).
 */
import { pv } from './personen-verwaltung-state.js';

export function graphErrorFriendly(e) {
    const raw = String(e && e.message ? e.message : e);
    const idx = raw.indexOf('{');
    if (idx !== -1) {
        try {
            const obj = JSON.parse(raw.slice(idx));
            const inner = obj.error || obj;
            if (inner && inner.message) return String(inner.message);
        } catch {
            // ignore
        }
    }
    return raw;
}
export function norm(s) {
    return String(s || '').trim().toLowerCase();
}

export function compareStrings(a, b) {
    return String(a || '').localeCompare(String(b || ''), 'de', { sensitivity: 'base' });
}

export function readInputTrim(el) {
    if (!el) return '';
    return String(el.value || '').trim();
}

export function readSortFromSelect() {
    const sel = document.getElementById('pvSortKey');
    const raw = sel && sel.value ? String(sel.value) : 'displayName:asc';
    const parts = raw.split(':');
    const key = parts[0] || 'displayName';
    const dir = parts[1] === 'desc' ? 'desc' : 'asc';
    return { key: key, dir: dir };
}

export function formatPhones(u) {
    const m = u && u.mobilePhone ? String(u.mobilePhone).trim() : '';
    const bp = u && Array.isArray(u.businessPhones) ? u.businessPhones.filter(Boolean).join(', ') : '';
    if (m && bp) return m + ' · ' + bp;
    return m || bp || '';
}

export function formatDate(iso) {
    if (!iso) return '–';
    try {
        const d = new Date(iso);
        if (isNaN(d.getTime())) return String(iso);
        return d.toLocaleString(undefined, {
            dateStyle: 'medium',
            timeStyle: 'short'
        });
    } catch {
        return String(iso);
    }
}

export function groupTypeLabel(g) {
    if (!g || typeof g !== 'object') return '–';
    const types = g.groupTypes;
    if (Array.isArray(types) && types.indexOf('Unified') !== -1) return 'Microsoft 365 (Unified)';
    if (g.securityEnabled && !g.mailEnabled) return 'Sicherheitsgruppe';
    if (g.mailEnabled && !g.securityEnabled) return 'Verteilerliste';
    if (g.securityEnabled && g.mailEnabled) return 'Mail-aktivierte Sicherheitsgruppe';
    return 'Gruppe';
}

export function userTypeLabel(ut) {
    const t = String(ut || '').toLowerCase();
    if (t === 'guest') return 'Gast';
    if (t === 'member') return 'Mitglied';
    return ut ? String(ut) : '–';
}
export function sanitizeMailNickname(raw, fallbackFromUpn) {
    let s = String(raw || '').trim();
    if (!s && fallbackFromUpn) {
        const at = fallbackFromUpn.indexOf('@');
        s = at > 0 ? fallbackFromUpn.slice(0, at) : fallbackFromUpn;
    }
    s = s.split('@')[0].replace(/[^a-zA-Z0-9._-]/g, '');
    return s;
}
export function isGuid(s) {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(s || '').trim());
}

export function isDuplicateMemberError(e) {
    const m = String((e && e.message) || e || '');
    return (
        m.indexOf('added object references already exist') !== -1 ||
        m.indexOf('One or more added object references already exist') !== -1 ||
        m.indexOf('already exist') !== -1
    );
}
export function assignedSkuIdsOfUser(u) {
    const list = u && Array.isArray(u.assignedLicenses) ? u.assignedLicenses : [];
    return list
        .map(function (l) {
            return String((l && l.skuId) || '').toLowerCase();
        })
        .filter(Boolean);
}

export function skuLookupFromSubscribed() {
    const map = new Map();
    (pv.subscribedSkus || []).forEach(function (s) {
        const id = String((s && s.skuId) || '').toLowerCase();
        if (!id) return;
        map.set(id, { skuId: id, skuPartNumber: String((s && s.skuPartNumber) || '') });
    });
    return map;
}
export function Lic() {
    return window.ms365GraphLicenses || null;
}

export function userLicenseSummary(u) {
    const api = Lic();
    if (!api || typeof api.summarizeUserLicenses !== 'function') return null;
    return api.summarizeUserLicenses(u);
}

export function loadAdUserFlags() {
    try {
        const raw = localStorage.getItem(pv.AD_FLAGS_KEY);
        if (!raw) return {};
        const obj = JSON.parse(raw);
        if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return {};
        return obj;
    } catch {
        return {};
    }
}

export function saveAdUserFlags(map) {
    try {
        localStorage.setItem(pv.AD_FLAGS_KEY, JSON.stringify(map && typeof map === 'object' ? map : {}));
    } catch {
        /* ignore */
    }
}

export function applyAdFlagsToUsers(users) {
    const flags = loadAdUserFlags();
    return (Array.isArray(users) ? users : []).map(function (u) {
        if (!u || !u.id) return u;
        const f = flags[String(u.id)] || null;
        const next = Object.assign({}, u);
        next.onPremisesSyncEnabled = u.onPremisesSyncEnabled === true;
        next.adFlagged = !!(f && f.flagged);
        next.adFlagNote = f && f.note ? String(f.note) : '';
        next.adFlaggedAt = f && f.flaggedAt ? String(f.flaggedAt) : '';
        return next;
    });
}
