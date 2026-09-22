'use strict';

/**
 * Domains/URLs aus Listenzeile normalisieren (eine pro Zeile, Komma oder Semikolon).
 * @param {unknown} primary
 * @param {unknown} additional
 * @returns {string[]}
 */
function parseDomainList(primary, additional) {
    const chunks = [primary, additional]
        .map((v) => String(v == null ? '' : v))
        .join('\n');
    const raw = chunks
        .split(/[\n,;]+/)
        .map((s) => s.trim())
        .filter(Boolean);

    const out = [];
    const seen = new Set();
    for (let i = 0; i < raw.length; i++) {
        let entry = raw[i];
        // URL → Hostname, sonst Domain so lassen
        try {
            if (/^https?:\/\//i.test(entry)) {
                entry = new URL(entry).hostname;
            } else if (entry.indexOf('/') >= 0 && entry.indexOf(' ') < 0) {
                entry = entry.replace(/^\/+/, '').split('/')[0];
            }
        } catch {
            /* raw behalten */
        }
        entry = entry.replace(/^\.+/, '').replace(/\.+$/, '').toLowerCase();
        if (!entry || seen.has(entry)) continue;
        seen.add(entry);
        out.push(entry);
    }
    return out;
}

/**
 * Pure Auswertung: Listenzeile → erlaubt ja/nein.
 * @param {{
 *   tenantId: string,
 *   fields: Record<string, unknown> | null,
 *   allowedStatuses: string[],
 *   now?: Date
 * }} input
 */
function evaluateLicense(input) {
    const tenantId = String(input.tenantId || '').trim().toLowerCase();
    const allowedStatuses = (input.allowedStatuses || []).map((s) => String(s).toLowerCase());
    const now = input.now instanceof Date ? input.now : new Date();
    const fields = input.fields && typeof input.fields === 'object' ? input.fields : null;

    const empty = {
        allowed: false,
        reason: 'missing_tenant',
        message: 'Keine Tenant-ID im Token.',
        schoolName: null,
        status: null,
        validUntil: null,
        primaryDomain: null,
        domains: [],
        contactEmail: null
    };

    if (!tenantId) {
        return empty;
    }

    if (!fields) {
        return {
            ...empty,
            reason: 'not_registered',
            message: 'Dieser Mandant ist nicht freigeschaltet.'
        };
    }

    const schoolName = String(fields.Title || '').trim() || null;
    const status = String(fields.Status || '').trim().toLowerCase() || null;
    const primaryDomain = String(fields.PrimaryDomain || '').trim() || null;
    const domains = parseDomainList(fields.PrimaryDomain, fields.AdditionalDomains);
    const contactEmail = String(fields.ContactEmail || '').trim() || null;
    const validUntilRaw = fields.ValidUntil;
    let validUntil = null;
    if (validUntilRaw != null && String(validUntilRaw).trim() !== '') {
        const d = new Date(validUntilRaw);
        if (!Number.isNaN(d.getTime())) {
            validUntil = d.toISOString().slice(0, 10);
        }
    }

    const base = {
        schoolName,
        status,
        validUntil,
        primaryDomain,
        domains,
        contactEmail
    };

    if (status === 'blocked') {
        return {
            allowed: false,
            reason: 'blocked',
            message: 'Der Zugang für diesen Mandanten ist gesperrt.',
            ...base
        };
    }

    if (status === 'expired') {
        return {
            allowed: false,
            reason: 'expired',
            message: 'Die Lizenz für diesen Mandanten ist abgelaufen.',
            ...base
        };
    }

    if (!status || !allowedStatuses.includes(status)) {
        return {
            allowed: false,
            reason: 'status_not_allowed',
            message: 'Der Lizenzstatus erlaubt keinen Zugang (' + (status || 'leer') + ').',
            ...base
        };
    }

    if (validUntil) {
        const end = new Date(validUntil + 'T23:59:59.999Z');
        if (now.getTime() > end.getTime()) {
            return {
                allowed: false,
                reason: 'expired',
                message: 'Die Lizenz ist seit ' + validUntil + ' abgelaufen.',
                ...base
            };
        }
    }

    return {
        allowed: true,
        reason: 'ok',
        message: 'Zugang freigeschaltet.',
        ...base
    };
}

/**
 * @param {unknown} items Graph list items with fields
 * @param {string} tenantId
 */
function findFieldsForTenant(items, tenantId) {
    const want = String(tenantId || '').trim().toLowerCase();
    const rows = Array.isArray(items) ? items : [];
    for (let i = 0; i < rows.length; i++) {
        const fields = rows[i] && rows[i].fields ? rows[i].fields : null;
        if (!fields) continue;
        const tid = String(fields.TenantId || '').trim().toLowerCase();
        if (tid && tid === want) return fields;
    }
    return null;
}

module.exports = { evaluateLicense, findFieldsForTenant, parseDomainList };
