'use strict';

/**
 * Schulprofil ↔ Lizenzliste: sichere Stammdaten für GET/PATCH /license/me.
 * Spiegel der erlaubten Felder in src/shared/school-profile-license-sync-logic.js
 */

const { SCHOOL_STAMMDATEN_FIELDS } = require('./school-stammdaten-catalog');

/** Felder, die Schul-IT melden/ändern darf (kein Workshop/Meta). */
const SCHOOL_PROFILE_SYNC_FIELD_NAMES = [
    'Standort',
    'Strasse',
    'PLZ',
    'Ort',
    'Direktion',
    'DirektionEmail',
    'Schuladministration',
    'SchuladministrationEmail',
    'ItM365Admin',
    'ItM365AdminEmail',
    'WeitereKontakte',
    'Telefon',
    'Schulnummer'
];

const SYNC_SET = new Set(SCHOOL_PROFILE_SYNC_FIELD_NAMES.map((n) => n.toLowerCase()));

/** @param {unknown} raw */
function norm(raw) {
    return String(raw == null ? '' : raw).trim();
}

/**
 * @param {Record<string, unknown>|null|undefined} fields
 * @returns {Record<string, string>}
 */
function pickSchoolStammdatenFromFields(fields) {
    const f = fields && typeof fields === 'object' ? fields : {};
    /** @type {Record<string, string>} */
    const out = {};
    for (let i = 0; i < SCHOOL_PROFILE_SYNC_FIELD_NAMES.length; i++) {
        const name = SCHOOL_PROFILE_SYNC_FIELD_NAMES[i];
        const v = norm(f[name]);
        if (v) out[name] = v;
    }
    return out;
}

/**
 * Body aus PATCH /license/me → Update-Payload für updateLicense (partial).
 * Erlaubt: schoolName, schoolStammdaten[sync fields]. Kein Status/Tenant/Entitlements/Domain.
 * @param {Record<string, unknown>|null|undefined} body
 * @returns {{ ok: true, patch: Record<string, unknown> } | { ok: false, reason: string }}
 */
function normalizeSchoolProfileMePatch(body) {
    const src = body && typeof body === 'object' ? body : {};
    /** @type {Record<string, unknown>} */
    const patch = {};

    if (src.schoolName !== undefined || src.Title !== undefined) {
        const name = norm(src.schoolName != null ? src.schoolName : src.Title);
        if (!name) return { ok: false, reason: 'empty_school_name' };
        patch.schoolName = name;
    }

    const rawSt =
        src.schoolStammdaten && typeof src.schoolStammdaten === 'object'
            ? src.schoolStammdaten
            : src.extra && typeof src.extra === 'object'
              ? src.extra
              : null;

    if (rawSt) {
        /** @type {Record<string, string>} */
        const extra = {};
        Object.keys(rawSt).forEach((key) => {
            const k = String(key || '').trim();
            if (!k || !SYNC_SET.has(k.toLowerCase())) return;
            const catalog = SCHOOL_STAMMDATEN_FIELDS.find(
                (f) => String(f.name).toLowerCase() === k.toLowerCase()
            );
            const name = catalog ? catalog.name : k;
            const v = norm(rawSt[key]);
            if (v) extra[name] = v;
        });
        if (Object.keys(extra).length) patch.extra = extra;
    }

    if (!patch.schoolName && !patch.extra) {
        return { ok: false, reason: 'empty_profile_patch' };
    }
    return { ok: true, patch };
}

/**
 * Gleiche Authz wie Tool-Site-Meldung.
 * @param {{ allowed?: boolean, contactEmails?: string[], upn?: string }} input
 * @param {(input: object) => { ok: boolean, reason: string }} canReportToolSiteUrl
 */
function canReportSchoolProfile(input, canReportToolSiteUrl) {
    return canReportToolSiteUrl(input);
}

module.exports = {
    SCHOOL_PROFILE_SYNC_FIELD_NAMES,
    pickSchoolStammdatenFromFields,
    normalizeSchoolProfileMePatch,
    canReportSchoolProfile
};
