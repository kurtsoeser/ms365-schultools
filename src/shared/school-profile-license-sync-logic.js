/**
 * Schulprofil ↔ Lizenz-Admin (Betreiber-SharePoint): Feldmapping ohne API-I/O.
 * Nur Kontakt-/Adress-/Rollen-Stammdaten – keine Schülerlisten, kein Logo.
 */

import { splitRolePersonFields } from './school-profile-logic.js';

/** Lizenz-Extra-Spalten, die die Schul-IT pflegen darf (nicht Workshop/Meta). */
export const LICENSE_PROFILE_SYNC_FIELDS = [
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

/** @type {Record<string, string>} license field → schoolProfile key */
export const LICENSE_TO_PROFILE_KEY = {
    Standort: 'standort',
    Strasse: 'street',
    PLZ: 'postalCode',
    Ort: 'city',
    Direktion: 'direktion',
    DirektionEmail: 'direktionEmail',
    Schuladministration: 'schuladministration',
    SchuladministrationEmail: 'schuladministrationEmail',
    ItM365Admin: 'itM365Admin',
    ItM365AdminEmail: 'itM365AdminEmail',
    WeitereKontakte: 'weitereKontakte',
    Telefon: 'phone',
    Schulnummer: 'schoolCode'
};

/** @param {unknown} raw */
function norm(raw) {
    return String(raw == null ? '' : raw).trim();
}

/**
 * Stammdaten-Objekt aus Lizenz-Antwort / SharePoint-Feldern.
 * @param {Record<string, unknown>|null|undefined} src
 * @returns {Record<string, string>}
 */
export function pickLicenseSchoolStammdaten(src) {
    const o = src && typeof src === 'object' ? src : {};
    /** @type {Record<string, string>} */
    const out = {};
    for (let i = 0; i < LICENSE_PROFILE_SYNC_FIELDS.length; i++) {
        const name = LICENSE_PROFILE_SYNC_FIELDS[i];
        const v = norm(o[name] != null ? o[name] : o[name.toLowerCase()]);
        if (v) out[name] = v;
    }
    return out;
}

/**
 * @param {{ schoolName?: string, primaryDomain?: string, schoolStammdaten?: Record<string, unknown>|null }|null|undefined} license
 */
export function normalizeLicenseProfileSnapshot(license) {
    const lic = license && typeof license === 'object' ? license : {};
    return {
        schoolName: norm(lic.schoolName),
        primaryDomain: norm(lic.primaryDomain),
        schoolStammdaten: pickLicenseSchoolStammdaten(lic.schoolStammdaten || lic)
    };
}

/**
 * Lizenz → lokale Profilfelder (nur Schlüssel, die syncbar sind).
 * Alte Rollenfelder mit eingebetteter E-Mail werden in Name + E-Mail getrennt.
 * @param {ReturnType<typeof normalizeLicenseProfileSnapshot>} snap
 */
export function licenseSnapshotToProfilePartial(snap) {
    const s = snap || normalizeLicenseProfileSnapshot(null);
    /** @type {Record<string, string>} */
    const partial = {};
    const st = s.schoolStammdaten || {};
    Object.keys(LICENSE_TO_PROFILE_KEY).forEach(function (licKey) {
        const profileKey = LICENSE_TO_PROFILE_KEY[licKey];
        const v = norm(st[licKey]);
        if (v) partial[profileKey] = v;
    });

    function applyRoleSplit(nameKey, emailKey) {
        const split = splitRolePersonFields(partial[nameKey], partial[emailKey]);
        if (split.name) partial[nameKey] = split.name;
        else delete partial[nameKey];
        if (split.email) partial[emailKey] = split.email;
        else delete partial[emailKey];
    }
    applyRoleSplit('direktion', 'direktionEmail');
    applyRoleSplit('schuladministration', 'schuladministrationEmail');
    applyRoleSplit('itM365Admin', 'itM365AdminEmail');

    // Standort oft Ort-ähnlich: wenn city leer, Standort als city-Hinweis mitnehmen
    if (!partial.city && partial.standort) partial.city = partial.standort;
    return partial;
}

/**
 * Leere lokale Felder aus Lizenz füllen (bestehende lokale Werte bleiben).
 * @param {{ schoolName?: string, domain?: string, profile?: Record<string, unknown> }} local
 * @param {ReturnType<typeof normalizeLicenseProfileSnapshot>} licenseSnap
 */
export function mergeLicenseIntoLocalFillEmpty(local, licenseSnap) {
    const loc = local && typeof local === 'object' ? local : {};
    const snap = normalizeLicenseProfileSnapshot(licenseSnap);
    const profileIn = loc.profile && typeof loc.profile === 'object' ? loc.profile : {};
    /** @type {Record<string, string>} */
    const profileOut = Object.assign({}, profileIn);
    const fromLic = licenseSnapshotToProfilePartial(snap);
    let profileChanged = false;
    Object.keys(fromLic).forEach(function (key) {
        if (!norm(profileOut[key]) && fromLic[key]) {
            profileOut[key] = fromLic[key];
            profileChanged = true;
        }
    });
    const schoolNameLocal = norm(loc.schoolName);
    const domainLocal = norm(loc.domain);
    const schoolName = schoolNameLocal || snap.schoolName || '';
    const domain = domainLocal || snap.primaryDomain || '';
    return {
        schoolName,
        domain,
        profile: profileOut,
        changed: profileChanged || schoolName !== schoolNameLocal || domain !== domainLocal,
        filledSchoolName: !schoolNameLocal && !!snap.schoolName,
        filledDomain: !domainLocal && !!snap.primaryDomain
    };
}

/**
 * PATCH-Body für license/me – nur nicht-leere syncbare Werte (kein Wipe der Admin-Daten).
 * @param {{ schoolName?: string, profile?: Record<string, unknown> }} local
 */
export function buildLicenseProfilePatchFromLocal(local) {
    const loc = local && typeof local === 'object' ? local : {};
    const profile = loc.profile && typeof loc.profile === 'object' ? loc.profile : {};
    /** @type {Record<string, string>} */
    const schoolStammdaten = {};
    Object.keys(LICENSE_TO_PROFILE_KEY).forEach(function (licKey) {
        const profileKey = LICENSE_TO_PROFILE_KEY[licKey];
        const v = norm(profile[profileKey]);
        if (v) schoolStammdaten[licKey] = v;
    });
    // Standort aus Ort, falls nur city gesetzt
    if (!schoolStammdaten.Standort && schoolStammdaten.Ort) {
        schoolStammdaten.Standort = schoolStammdaten.Ort;
    }
    /** @type {{ schoolName?: string, schoolStammdaten?: Record<string, string> }} */
    const body = {};
    const schoolName = norm(loc.schoolName);
    if (schoolName) body.schoolName = schoolName;
    if (Object.keys(schoolStammdaten).length) body.schoolStammdaten = schoolStammdaten;
    return body;
}

/**
 * @param {ReturnType<typeof buildLicenseProfilePatchFromLocal>} patch
 */
export function licenseProfilePatchIsEmpty(patch) {
    if (!patch || typeof patch !== 'object') return true;
    if (norm(patch.schoolName)) return false;
    const st = patch.schoolStammdaten;
    return !(st && typeof st === 'object' && Object.keys(st).length);
}
