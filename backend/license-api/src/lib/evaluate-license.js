'use strict';

const {
    parseEntitlementsRaw,
    resolveFeatureEntitlements
} = require('./feature-catalog');
const { pickSchoolStammdatenFromFields } = require('./school-profile-sync');

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
 * IT-Kontakt(e) aus Listenfeld (eine pro Zeile, Komma oder Semikolon).
 * @param {unknown} primary
 * @param {unknown} [additional]
 * @returns {string[]}
 */
function parseContactEmails(primary, additional) {
    const chunks = [primary, additional]
        .map((v) => String(v == null ? '' : v))
        .join('\n');
    const raw = chunks
        .split(/[\n,;]+/)
        .map((s) => s.trim().toLowerCase())
        .filter(Boolean);
    const out = [];
    const seen = new Set();
    for (let i = 0; i < raw.length; i++) {
        const e = raw[i];
        if (!e.includes('@') || seen.has(e)) continue;
        seen.add(e);
        out.push(e);
    }
    return out;
}

/**
 * @param {string|null|undefined} upn
 * @param {string[]} contactEmails
 */
function upnMatchesDesignatedItContact(upn, contactEmails) {
    const u = String(upn || '')
        .trim()
        .toLowerCase();
    if (!u || !u.includes('@')) return false;
    const list = Array.isArray(contactEmails) ? contactEmails : [];
    for (let i = 0; i < list.length; i++) {
        if (String(list[i] || '').trim().toLowerCase() === u) return true;
    }
    return false;
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

    const emptyEntitlements = {
        packages: [],
        toolsInclude: [],
        toolsExclude: [],
        features: [],
        denied: [],
        featureCount: 0,
        legacyFullAccess: false
    };

    const empty = {
        allowed: false,
        reason: 'missing_tenant',
        message: 'Keine Tenant-ID im Token.',
        schoolName: null,
        status: null,
        validUntil: null,
        primaryDomain: null,
        domains: [],
        contactEmail: null,
        contactEmails: [],
        toolSiteUrl: null,
        schoolStammdaten: {},
        entitlements: emptyEntitlements
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
    const contactEmails = parseContactEmails(fields.ContactEmail, fields.ItContacts);
    const contactEmail = contactEmails.length ? contactEmails[0] : null;
    const validUntilRaw = fields.ValidUntil;
    let validUntil = null;
    if (validUntilRaw != null && String(validUntilRaw).trim() !== '') {
        const d = new Date(validUntilRaw);
        if (!Number.isNaN(d.getTime())) {
            validUntil = d.toISOString().slice(0, 10);
        }
    }

    const entitlementDoc = parseEntitlementsRaw(
        fields.LicenseEntitlements != null ? fields.LicenseEntitlements : fields.licenseEntitlements
    );
    const entitlements = resolveFeatureEntitlements(entitlementDoc, { legacyFullAccess: true });

    // SharePoint kann ToolSiteUrl0 o. Ä. erzeugen – nicht nur exakt ToolSiteUrl
    let toolSiteUrl = '';
    const toolKeys = Object.keys(fields || {});
    for (let ti = 0; ti < toolKeys.length; ti++) {
        const tk = toolKeys[ti];
        if (!/tool.?site.?url/i.test(tk) && !/tool_x002d_site/i.test(tk)) continue;
        const raw = fields[tk];
        if (raw && typeof raw === 'object' && (raw.Url || raw.url)) {
            toolSiteUrl = String(raw.Url || raw.url || '').trim();
        } else {
            toolSiteUrl = String(raw == null ? '' : raw).trim();
        }
        if (toolSiteUrl) break;
    }
    if (!toolSiteUrl) {
        toolSiteUrl = String(fields.ToolSiteUrl || fields.toolSiteUrl || '').trim();
    }
    toolSiteUrl = toolSiteUrl || null;

    const schoolStammdaten = pickSchoolStammdatenFromFields(fields);

    const base = {
        schoolName,
        status,
        validUntil,
        primaryDomain,
        domains,
        contactEmail,
        contactEmails,
        toolSiteUrl,
        schoolStammdaten,
        entitlements
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
    const hit = findLicenseItemForTenant(items, tenantId);
    return hit ? hit.fields : null;
}

/**
 * Listenzeile inkl. Item-ID (für PATCH toolSiteUrl durch Schul-IT).
 * @param {unknown} items
 * @param {string} tenantId
 * @returns {{ id: string, fields: Record<string, unknown> } | null}
 */
function findLicenseItemForTenant(items, tenantId) {
    const want = String(tenantId || '').trim().toLowerCase();
    const rows = Array.isArray(items) ? items : [];
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const fields = row && row.fields ? row.fields : null;
        if (!fields) continue;
        const tid = String(fields.TenantId || '').trim().toLowerCase();
        if (tid && tid === want) {
            return { id: String(row.id || ''), fields };
        }
    }
    return null;
}

/**
 * Darf die Schul-Person die Tool-Site-URL an den Betreiber melden?
 * Designierter IT-Kontakt; oder (Pilot) niemand hinterlegt → jeder mit gültiger Lizenz.
 * @param {{ allowed?: boolean, contactEmails?: string[], upn?: string }} input
 */
function canReportToolSiteUrl(input) {
    const o = input || {};
    if (!o.allowed) return { ok: false, reason: 'license_denied' };
    const contacts = Array.isArray(o.contactEmails) ? o.contactEmails : [];
    if (!contacts.length) return { ok: true, reason: 'no_contacts_bootstrap' };
    if (upnMatchesDesignatedItContact(o.upn, contacts)) {
        return { ok: true, reason: 'designated_it' };
    }
    return { ok: false, reason: 'not_designated_it' };
}

/**
 * Nur https SharePoint-/Sites-URLs als Tool-Site akzeptieren.
 * @param {unknown} raw
 * @returns {{ ok: true, url: string } | { ok: false, reason: string }}
 */
function normalizeReportedToolSiteUrl(raw) {
    const s = String(raw || '')
        .trim()
        .replace(/\/+$/, '');
    if (!s) return { ok: false, reason: 'empty' };
    let u;
    try {
        u = new URL(s);
    } catch {
        return { ok: false, reason: 'invalid_url' };
    }
    if (!/^https:$/i.test(u.protocol)) return { ok: false, reason: 'not_https' };
    const host = String(u.hostname || '').toLowerCase();
    const path = String(u.pathname || '');
    const looksSharePoint =
        host.endsWith('.sharepoint.com') ||
        host === 'sharepoint.com' ||
        /\/sites\//i.test(path) ||
        /\/teams\//i.test(path);
    if (!looksSharePoint) return { ok: false, reason: 'not_sharepoint' };
    return { ok: true, url: (u.origin + u.pathname).replace(/\/+$/, '') };
}

module.exports = {
    evaluateLicense,
    findFieldsForTenant,
    findLicenseItemForTenant,
    parseDomainList,
    parseContactEmails,
    upnMatchesDesignatedItContact,
    canReportToolSiteUrl,
    normalizeReportedToolSiteUrl
};
