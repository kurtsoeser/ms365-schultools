/**
 * Education-Lizenzen (Graph assignedLicenses / skuPartNumber).
 * Reine Logik: Klassifikation, Kürzel-Vorschlag, Lehrerlisten-Merge.
 *
 * Katalog angelehnt an Microsoft Learn „Education SKU reference“ und
 * gängige Office-/Microsoft-365-A1/A3/A5-SKU-IDs für Lehrpersonal.
 */

import { normCode, normEmail, normStr } from './utils/strings.js';

/** @typedef {'faculty'|'student'|'other'} LicenseAudience */
/** @typedef {'a1'|'a3'|'a5'|'apps'|'exchange'|'sharepoint'|'rooms'|'device'|'other'} LicenseFamily */

/**
 * @typedef {object} SkuInfo
 * @property {string} skuId
 * @property {string} skuPartNumber
 * @property {string} name
 * @property {string} shortLabel
 * @property {LicenseAudience} audience
 * @property {LicenseFamily} family
 * @property {boolean} userPlan  Vollwertiger Benutzerplan (A1/A3/A5/E1/E3), nicht Gerät/Add-on
 */

function sku(
    skuId,
    skuPartNumber,
    name,
    shortLabel,
    audience,
    family,
    userPlan
) {
    return {
        skuId: String(skuId).toLowerCase(),
        skuPartNumber,
        name,
        shortLabel,
        audience,
        family,
        userPlan: !!userPlan
    };
}

/** Bekannte Education-SKUs (skuId kleingeschrieben). */
export const EDU_SKU_CATALOG = [
    sku('94763226-9b3c-4e75-a931-5c89701abe66', 'STANDARDWOFFPACK_FACULTY', 'Office 365 A1 für Lehrpersonal', 'A1 Lehrpersonal', 'faculty', 'a1', true),
    sku('884e415c-1efb-4eba-9295-36bbd3148e99', 'STANDARDWOFFPACK_FACULTY_DE', 'Office 365 A1 für Lehrpersonal (DE)', 'A1 Lehrpersonal', 'faculty', 'a1', true),
    sku('78e66a63-337a-4a9a-8959-41c6654dfb56', 'STANDARDWOFFPACK_IW_FACULTY', 'Office 365 A1 Plus für Lehrpersonal', 'A1 Plus Lehrpersonal', 'faculty', 'a1', true),
    sku('a19037fc-48b4-4d57-b079-ce44b7832473', 'STANDARDPACK_FACULTY', 'Office 365 Education E1 für Lehrpersonal', 'E1 Lehrpersonal', 'faculty', 'a1', true),
    sku('43e691ad-1491-4e8c-8dc9-da6b8262c03b', 'STANDARDWOFFPACK_HOMESCHOOL_FAC', 'Office 365 Homeschool für Lehrpersonal', 'A1 Lehrpersonal', 'faculty', 'a1', true),
    sku('4b590615-0888-425a-a965-b3bf7789848d', 'M365EDU_A3_FACULTY', 'Microsoft 365 A3 für Lehrpersonal', 'A3 Lehrpersonal', 'faculty', 'a3', true),
    sku('e578b273-6db4-4691-bba0-8d691f4da603', 'ENTERPRISEPACKPLUS_FACULTY', 'Office 365 A3 für Lehrpersonal', 'A3 Lehrpersonal', 'faculty', 'a3', true),
    sku('e4fa3838-3d01-42df-aa28-5e0a4c68604b', 'ENTERPRISEPACK_FACULTY', 'Office 365 Education E3 für Lehrpersonal', 'E3 Lehrpersonal', 'faculty', 'a3', true),
    sku('eed7a755-c6cf-49b1-ace3-32eb3627c498', 'ENTERPRISEPACK_FACULTY_DE', 'Office 365 Education E3 für Lehrpersonal (DE)', 'E3 Lehrpersonal', 'faculty', 'a3', true),
    sku('e97c048c-37a4-45fb-ab50-922fbf07a370', 'M365EDU_A5_FACULTY', 'Microsoft 365 A5 für Lehrpersonal', 'A5 Lehrpersonal', 'faculty', 'a5', true),
    sku('65200ac3-f927-4407-a3d5-c63562dff461', 'M365EDU_A5_NOPSTNCONF_FACULTY', 'Microsoft 365 A5 ohne Audiokonferenz für Lehrpersonal', 'A5 Lehrpersonal', 'faculty', 'a5', true),
    sku('ea73fc9b-3f94-418d-b128-1181dc9fb125', 'M365EDU_A5_FACULTY_CALLINGMINUTES', 'Microsoft 365 A5 mit Calling Minutes für Lehrpersonal', 'A5 Lehrpersonal', 'faculty', 'a5', true),
    sku('a4585165-0533-458a-97e3-c400570268c4', 'ENTERPRISEPREMIUM_FACULTY', 'Office 365 A5 für Lehrpersonal', 'A5 Lehrpersonal', 'faculty', 'a5', true),
    sku('9a320620-ca3d-4705-a79d-27c135c96e05', 'ENTERPRISEPREMIUM_NOPSTNCONF_FACULTY', 'Office 365 A5 ohne Audiokonferenz für Lehrpersonal', 'A5 Lehrpersonal', 'faculty', 'a5', true),
    sku('1a18ad49-7cd3-4d0a-bb2a-5108dbd226ae', 'ENTERPRISEPREMIUM_FACULTY_CALLINGMINUTES', 'Office 365 A5 mit Calling Minutes für Lehrpersonal', 'A5 Lehrpersonal', 'faculty', 'a5', true),
    sku('12b8c807-2e20-48fc-b453-542b6ee9d171', 'OFFICESUBSCRIPTION_FACULTY', 'Microsoft 365 Apps für Lehrpersonal', 'Apps Lehrpersonal', 'faculty', 'apps', false),
    sku('af4e28de-6b52-4fd3-a5f4-6bf708a304d3', 'STANDARDWOFFPACK_FACULTY_DEVICE', 'Office 365 A1 für Lehrpersonal (Gerät)', 'A1 Gerät', 'faculty', 'device', false),
    sku('c07395e9-4cec-4b3e-9f63-f6933efbb6c2', 'MICROSOFT_365_A1_FOR_DEVICES_FAC', 'Microsoft 365 A1 für Geräte (Lehrpersonal)', 'A1 Gerät', 'faculty', 'device', false),
    sku('a4e376bd-c61e-4618-9901-3fc0cb1b88bb', 'Microsoft_Teams_Rooms_Basic_FAC', 'Teams Rooms Basic (Bildung)', 'Rooms', 'other', 'rooms', false),
    sku('c25e2b36-e161-4946-bef2-69239729f690', 'Microsoft_Teams_Rooms_Pro_FAC', 'Teams Rooms Pro (Bildung)', 'Rooms', 'other', 'rooms', false),

    sku('314c4481-f395-4525-be8b-2ec4bb1e9d91', 'STANDARDWOFFPACK_STUDENT', 'Office 365 A1 für Schüler:innen', 'A1 Schüler:innen', 'student', 'a1', true),
    sku('d37ba356-38c5-4c82-90da-3d714f72a382', 'STANDARDPACK_STUDENT', 'Office 365 Education E1 für Schüler:innen', 'E1 Schüler:innen', 'student', 'a1', true),
    sku('afbb89a7-db5f-45fb-8af0-1bc5c5015709', 'STANDARDWOFFPACK_HOMESCHOOL_STU', 'Office 365 Homeschool für Schüler:innen', 'A1 Schüler:innen', 'student', 'a1', true),
    sku('7cfd9a2b-e110-4c39-bf20-c6a3f36a3121', 'M365EDU_A3_STUDENT', 'Microsoft 365 A3 für Schüler:innen', 'A3 Schüler:innen', 'student', 'a3', true),
    sku('18250162-5d87-4436-a834-d795c15c80f3', 'M365EDU_A3_STUUSEBNFT', 'Microsoft 365 A3 Student Use Benefit', 'A3 Schüler:innen', 'student', 'a3', true),
    sku('98b6e773-24d4-4c0d-a968-6e787a1f8204', 'ENTERPRISEPACKPLUS_STUDENT', 'Office 365 A3 für Schüler:innen', 'A3 Schüler:innen', 'student', 'a3', true),
    sku('476aad1e-7a7f-473c-9d20-35665a5cbd4f', 'ENTERPRISEPACKPLUS_STUUSEBNFT', 'Office 365 A3 Student Use Benefit', 'A3 Schüler:innen', 'student', 'a3', true),
    sku('46c119d4-0379-4a9d-85e4-97c66d3f909e', 'M365EDU_A5_STUDENT', 'Microsoft 365 A5 für Schüler:innen', 'A5 Schüler:innen', 'student', 'a5', true),
    sku('31d57bc7-3a05-4867-ab53-97a17835a411', 'M365EDU_A5_STUUSEBNFT', 'Microsoft 365 A5 Student Use Benefit', 'A5 Schüler:innen', 'student', 'a5', true),
    sku('ee656612-49fa-43e5-b67e-cb1fdf7699df', 'ENTERPRISEPREMIUM_STUDENT', 'Office 365 A5 für Schüler:innen', 'A5 Schüler:innen', 'student', 'a5', true),
    sku('a25c01ce-bab1-47e9-a6d0-ebe939b99ff9', 'M365EDU_A5_NOPSTNCONF_STUDENT', 'Microsoft 365 A5 ohne Audiokonferenz für Schüler:innen', 'A5 Schüler:innen', 'student', 'a5', true),
    sku('81441ae1-0b31-4185-a6c0-32b6b84d419f', 'M365EDU_A5_NOPSTNCONF_STUUSEBNFT', 'Microsoft 365 A5 ohne Audiokonferenz Student Use Benefit', 'A5 Schüler:innen', 'student', 'a5', true),
    sku('1164451b-e2e5-4c9e-8fa6-e5122d90dbdc', 'ENTERPRISEPREMIUM_NOPSTNCONF_STUDENT', 'Office 365 A5 ohne Audiokonferenz für Schüler:innen', 'A5 Schüler:innen', 'student', 'a5', true),
    sku('f6e603f1-1a6d-4d32-a730-34b809cb9731', 'ENTERPRISEPREMIUM_STUUSEBNFT', 'Office 365 A5 Student Use Benefit', 'A5 Schüler:innen', 'student', 'a5', true)
];

const CATALOG_BY_ID = new Map(EDU_SKU_CATALOG.map((s) => [s.skuId, s]));
const CATALOG_BY_PART = new Map(EDU_SKU_CATALOG.map((s) => [s.skuPartNumber.toUpperCase(), s]));

const TITLE_RE = /^(mag|dr|prof|dipl|di|ing|ba|ma|phd|priv|doz|rer|nat|med|phil|iur|h|c)$/i;

function foldLettersUpper(s) {
    return String(s || '')
        .replace(/[Ää]/g, 'Ae')
        .replace(/[Öö]/g, 'Oe')
        .replace(/[Üü]/g, 'Ue')
        .replace(/ß/g, 'ss')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^A-Za-z]/g, '')
        .toUpperCase();
}

function classifyFromPartNumber(partRaw) {
    const p = String(partRaw || '').toUpperCase();
    if (!p) return null;
    const known = CATALOG_BY_PART.get(p);
    if (known) return known;

    const isStudent = /STUDENT|STUUSEBNFT|_STU(?:_|$)|HOMESCHOOL_STU/.test(p);
    const isFaculty = !isStudent && /FACULTY|_FAC(?:_|$)|HOMESCHOOL_FAC/.test(p);
    if (!isStudent && !isFaculty) return null;

    const audience = isStudent ? 'student' : 'faculty';
    const who = isStudent ? 'Schüler:innen' : 'Lehrpersonal';
    let family = 'other';
    let userPlan = false;
    let short = who;
    if (/DEVICE|FOR_DEVICES/.test(p)) {
        family = 'device';
        short = 'Gerät ' + who;
    } else if (/ROOMS/.test(p)) {
        family = 'rooms';
        short = 'Rooms';
    } else if (/EXCHANGE/.test(p)) {
        family = 'exchange';
        short = 'Exchange ' + who;
    } else if (/SHAREPOINT|STORAGE/.test(p)) {
        family = 'sharepoint';
        short = 'SharePoint ' + who;
    } else if (/OFFICESUBSCRIPTION|M365_APPS|MICROSOFT_365_APPS/.test(p)) {
        family = 'apps';
        short = 'Apps ' + who;
    } else if (/A5|ENTERPRISEPREMIUM/.test(p)) {
        family = 'a5';
        userPlan = true;
        short = 'A5 ' + who;
    } else if (/A3|ENTERPRISEPACK/.test(p)) {
        family = 'a3';
        userPlan = true;
        short = 'A3 ' + who;
    } else if (/A1|STANDARDWOFFPACK|STANDARDPACK/.test(p)) {
        family = 'a1';
        userPlan = true;
        short = 'A1 ' + who;
    }
    return {
        skuId: '',
        skuPartNumber: partRaw,
        name: short,
        shortLabel: short,
        audience,
        family,
        userPlan
    };
}

/**
 * @param {string} skuId
 * @param {string} [skuPartNumber]
 * @returns {SkuInfo}
 */
export function resolveSku(skuId, skuPartNumber) {
    const id = String(skuId || '').toLowerCase();
    if (id && CATALOG_BY_ID.has(id)) return CATALOG_BY_ID.get(id);
    const fromPart = classifyFromPartNumber(skuPartNumber);
    if (fromPart) {
        return Object.assign({}, fromPart, { skuId: id || fromPart.skuId });
    }
    const short = skuPartNumber ? String(skuPartNumber) : id ? id.slice(0, 8) + '…' : 'Unbekannt';
    return {
        skuId: id,
        skuPartNumber: String(skuPartNumber || ''),
        name: short,
        shortLabel: 'Andere',
        audience: 'other',
        family: 'other',
        userPlan: false
    };
}

/** skuIds der vollwertigen Lehrpersonal-Pläne (A1/A3/A5/E1/E3), für Graph-$filter. */
export function facultyUserPlanSkuIds() {
    return EDU_SKU_CATALOG.filter((s) => s.audience === 'faculty' && s.userPlan).map((s) => s.skuId);
}

/** skuIds der vollwertigen Schüler-Pläne (A1/A3/A5), für Graph-$filter. */
export function studentUserPlanSkuIds() {
    return EDU_SKU_CATALOG.filter((s) => s.audience === 'student' && s.userPlan).map((s) => s.skuId);
}

/**
 * @param {Array<{skuId?: string, skuPartNumber?: string}>|null|undefined} assignedLicenses
 * @param {Map<string, {skuPartNumber?: string}>|Record<string, {skuPartNumber?: string}>|null} [skuLookup]
 * @returns {SkuInfo[]}
 */
export function licensesFromAssigned(assignedLicenses, skuLookup) {
    const list = Array.isArray(assignedLicenses) ? assignedLicenses : [];
    const lookup =
        skuLookup instanceof Map
            ? skuLookup
            : skuLookup && typeof skuLookup === 'object'
              ? new Map(
                    Object.keys(skuLookup).map((k) => [String(k).toLowerCase(), skuLookup[k]])
                )
              : null;
    const out = [];
    const seen = new Set();
    for (let i = 0; i < list.length; i++) {
        const raw = list[i] || {};
        const id = String(raw.skuId || '').toLowerCase();
        if (!id || seen.has(id)) continue;
        seen.add(id);
        const extra = lookup && lookup.get(id);
        const part = raw.skuPartNumber || (extra && extra.skuPartNumber) || '';
        out.push(resolveSku(id, part));
    }
    return out;
}

function familyRank(family) {
    if (family === 'a5') return 5;
    if (family === 'a3') return 4;
    if (family === 'a1') return 3;
    if (family === 'apps') return 2;
    return 1;
}

/**
 * @param {SkuInfo[]} licenses
 */
export function summarizeLicenses(licenses) {
    const list = Array.isArray(licenses) ? licenses : [];
    const facultyUser = list.filter((l) => l.audience === 'faculty' && l.userPlan);
    const studentUser = list.filter((l) => l.audience === 'student' && l.userPlan);
    const facultyAny = list.filter((l) => l.audience === 'faculty');
    const studentAny = list.filter((l) => l.audience === 'student');
    const pick = (arr) => {
        if (!arr.length) return null;
        return arr.slice().sort((a, b) => familyRank(b.family) - familyRank(a.family))[0];
    };
    const primary = pick(facultyUser) || pick(studentUser) || pick(facultyAny) || pick(studentAny) || list[0] || null;
    const facultyFamilies = [];
    const studentFamilies = [];
    list.forEach((l) => {
        if (l.audience === 'faculty' && l.userPlan && facultyFamilies.indexOf(l.family) === -1) {
            facultyFamilies.push(l.family);
        }
        if (l.audience === 'student' && l.userPlan && studentFamilies.indexOf(l.family) === -1) {
            studentFamilies.push(l.family);
        }
    });
    return {
        licenses: list,
        primary,
        primaryLabel: primary ? primary.shortLabel : list.length ? 'Lizenz' : 'Ohne Lizenz',
        hasFacultyUserPlan: facultyUser.length > 0,
        hasStudentUserPlan: studentUser.length > 0,
        hasFaculty: facultyAny.length > 0,
        hasStudent: studentAny.length > 0,
        hasAny: list.length > 0,
        facultyFamilies,
        studentFamilies
    };
}

/**
 * @param {object} user Graph-User mit assignedLicenses
 * @param {Map<string, {skuPartNumber?: string}>|null} [skuLookup]
 */
export function summarizeUserLicenses(user, skuLookup) {
    return summarizeLicenses(licensesFromAssigned(user && user.assignedLicenses, skuLookup));
}

/**
 * Filterwert für das Personen-Modul:
 * '' | 'faculty' | 'student' | 'none' | 'other' | 'faculty-a1' | 'faculty-a3' | 'faculty-a5' | 'sku:<guid>'
 * @param {object} user
 * @param {string} filterVal
 * @param {Map<string, {skuPartNumber?: string}>|null} [skuLookup]
 */
export function userMatchesLicenseFilter(user, filterVal, skuLookup) {
    const f = String(filterVal || '').trim();
    if (!f) return true;
    const sum = summarizeUserLicenses(user, skuLookup);
    if (f === 'none') return !sum.hasAny;
    if (f === 'faculty') return sum.hasFacultyUserPlan || sum.hasFaculty;
    if (f === 'student') return sum.hasStudentUserPlan || sum.hasStudent;
    if (f === 'other') return sum.hasAny && !sum.hasFaculty && !sum.hasStudent;
    if (f === 'faculty-a1') return sum.facultyFamilies.indexOf('a1') !== -1;
    if (f === 'faculty-a3') return sum.facultyFamilies.indexOf('a3') !== -1;
    if (f === 'faculty-a5') return sum.facultyFamilies.indexOf('a5') !== -1;
    if (f === 'student-a1') return (sum.studentFamilies || []).indexOf('a1') !== -1;
    if (f === 'student-a3') return (sum.studentFamilies || []).indexOf('a3') !== -1;
    if (f === 'student-a5') return (sum.studentFamilies || []).indexOf('a5') !== -1;
    if (f.indexOf('sku:') === 0) {
        const id = f.slice(4).toLowerCase();
        return sum.licenses.some((l) => l.skuId === id);
    }
    return true;
}

/**
 * Optionen für ein Lizenz-Filter-Select, aus geladenen Benutzern.
 * @param {object[]} users
 * @param {Map<string, {skuPartNumber?: string}>|null} [skuLookup]
 */
export function buildLicenseFilterOptions(users, skuLookup) {
    const list = Array.isArray(users) ? users : [];
    let faculty = 0;
    let student = 0;
    let none = 0;
    let other = 0;
    let a1 = 0;
    let a3 = 0;
    let a5 = 0;
    let stuA1 = 0;
    let stuA3 = 0;
    let stuA5 = 0;
    const skuCounts = new Map();
    for (let i = 0; i < list.length; i++) {
        const sum = summarizeUserLicenses(list[i], skuLookup);
        if (!sum.hasAny) none++;
        if (sum.hasFacultyUserPlan || sum.hasFaculty) faculty++;
        if (sum.hasStudentUserPlan || sum.hasStudent) student++;
        if (sum.hasAny && !sum.hasFaculty && !sum.hasStudent) other++;
        if (sum.facultyFamilies.indexOf('a1') !== -1) a1++;
        if (sum.facultyFamilies.indexOf('a3') !== -1) a3++;
        if (sum.facultyFamilies.indexOf('a5') !== -1) a5++;
        const stuFam = sum.studentFamilies || [];
        if (stuFam.indexOf('a1') !== -1) stuA1++;
        if (stuFam.indexOf('a3') !== -1) stuA3++;
        if (stuFam.indexOf('a5') !== -1) stuA5++;
        sum.licenses.forEach((l) => {
            if (!l.skuId) return;
            const prev = skuCounts.get(l.skuId) || { count: 0, label: l.shortLabel, name: l.name };
            prev.count++;
            skuCounts.set(l.skuId, prev);
        });
    }
    const options = [
        { value: '', label: '(alle Lizenzen)' },
        { value: 'faculty', label: 'Lehrpersonal (Education)' + (faculty ? ' · ' + faculty : '') },
        { value: 'faculty-a1', label: 'A1 für Lehrpersonal' + (a1 ? ' · ' + a1 : '') },
        { value: 'faculty-a3', label: 'A3 für Lehrpersonal' + (a3 ? ' · ' + a3 : '') },
        { value: 'faculty-a5', label: 'A5 für Lehrpersonal' + (a5 ? ' · ' + a5 : '') },
        { value: 'student', label: 'Schüler:innen (Education)' + (student ? ' · ' + student : '') },
        { value: 'student-a1', label: 'A1 für Schüler:innen' + (stuA1 ? ' · ' + stuA1 : '') },
        { value: 'student-a3', label: 'A3 für Schüler:innen' + (stuA3 ? ' · ' + stuA3 : '') },
        { value: 'student-a5', label: 'A5 für Schüler:innen' + (stuA5 ? ' · ' + stuA5 : '') },
        { value: 'none', label: 'Ohne Lizenz' + (none ? ' · ' + none : '') },
        { value: 'other', label: 'Andere Lizenzen' + (other ? ' · ' + other : '') }
    ];
    const skuOpts = Array.from(skuCounts.entries())
        .map(([id, info]) => ({
            value: 'sku:' + id,
            label: (info.name || info.label) + ' · ' + info.count
        }))
        .sort((a, b) => a.label.localeCompare(b.label, 'de', { sensitivity: 'base' }));
    return options.concat(skuOpts);
}

export function splitPersonName(displayName, givenName, surname) {
    let given = normStr(givenName);
    let sur = normStr(surname);
    if (!sur) {
        const raw = normStr(displayName).replace(/,/g, ' ');
        const parts = raw.split(/\s+/).filter(Boolean);
        const cleaned = parts.filter((p) => !TITLE_RE.test(p.replace(/\./g, '')));
        const use = cleaned.length ? cleaned : parts;
        if (use.length >= 2) {
            given = given || use[0];
            sur = use[use.length - 1];
        } else {
            sur = use[0] || '';
        }
    }
    return { given, surname: sur };
}

function usedHas(used, code) {
    return used.has(String(code || '').toLowerCase());
}

/**
 * Kürzel: 3 Buchstaben Nachname, bei Kollision + erster Buchstabe Vorname, sonst Zähler.
 * @param {string} displayName
 * @param {string} [givenName]
 * @param {string} [surname]
 * @param {Set<string>|string[]} [usedCodes]
 */
export function suggestTeacherCode(displayName, givenName, surname, usedCodes) {
    const used = usedCodes instanceof Set ? usedCodes : new Set(
        (Array.isArray(usedCodes) ? usedCodes : []).map((c) => String(c).toLowerCase())
    );
    const names = splitPersonName(displayName, givenName, surname);
    const surL = foldLettersUpper(names.surname);
    const givenL = foldLettersUpper(names.given);
    const candidates = [];
    if (surL.length >= 3) candidates.push(surL.slice(0, 3));
    else if (surL) candidates.push(surL);
    if (surL && givenL) {
        candidates.push((surL.slice(0, 2) + givenL.slice(0, 1)).slice(0, 4));
        candidates.push((surL.slice(0, 3) + givenL.slice(0, 1)).slice(0, 4));
    }
    for (let i = 0; i < candidates.length; i++) {
        const c = candidates[i];
        if (c && !usedHas(used, c)) return c;
    }
    const base = (candidates[0] || givenL.slice(0, 3) || 'LEH').slice(0, 4);
    if (base && !usedHas(used, base)) return base;
    let n = 2;
    let code = base + String(n);
    while (usedHas(used, code)) {
        n++;
        code = base + String(n);
    }
    return code;
}

export function teacherEmailOfUser(user) {
    const mail = normEmail(user && user.mail);
    if (mail && mail.indexOf('@') !== -1) return mail;
    const upn = normEmail(user && user.userPrincipalName);
    if (upn && upn.indexOf('@') !== -1) return upn;
    return '';
}

/**
 * @param {object[]} users Graph-User
 * @param {Array<{code?: string, name?: string, email?: string}>} existingTeachers
 * @param {Map<string, {skuPartNumber?: string}>|null} [skuLookup]
 * @param {{ activeOnly?: boolean, guests?: boolean, families?: string[] }} [opts]
 */
export function buildTeacherImportPreview(users, existingTeachers, skuLookup, opts) {
    const opt = opts || {};
    const families = Array.isArray(opt.families) ? opt.families : ['a1', 'a3', 'a5'];
    const familySet = new Set(families);
    const existing = Array.isArray(existingTeachers) ? existingTeachers : [];
    const emailToExisting = new Map();
    const usedCodes = new Set();
    existing.forEach((t) => {
        const em = normEmail(t && t.email);
        if (em) emailToExisting.set(em, t);
        if (t && t.code) usedCodes.add(String(t.code).toLowerCase());
    });

    const rows = [];
    const seenUser = new Set();
    (Array.isArray(users) ? users : []).forEach((u) => {
        if (!u || !u.id || seenUser.has(u.id)) return;
        if (opt.activeOnly && u.accountEnabled === false) return;
        if (opt.guests === false && String(u.userType || '').toLowerCase() === 'guest') return;
        const sum = summarizeUserLicenses(u, skuLookup);
        if (!sum.hasFacultyUserPlan) return;
        const hitFamily = sum.facultyFamilies.some((f) => familySet.has(f));
        if (!hitFamily) return;
        seenUser.add(u.id);
        const email = teacherEmailOfUser(u);
        const existingRow = email ? emailToExisting.get(email) : null;
        let code;
        if (existingRow && existingRow.code) {
            code = existingRow.code;
        } else {
            code = suggestTeacherCode(u.displayName, u.givenName, u.surname, usedCodes);
            usedCodes.add(String(code).toLowerCase());
        }
        rows.push({
            graphUserId: String(u.id),
            displayName: normStr(u.displayName),
            givenName: normStr(u.givenName),
            surname: normStr(u.surname),
            userPrincipalName: normStr(u.userPrincipalName),
            accountEnabled: u.accountEnabled !== false,
            userType: String(u.userType || 'Member'),
            email,
            code,
            name: normStr(u.displayName),
            licenseLabel: sum.primaryLabel,
            facultyFamilies: sum.facultyFamilies.slice(),
            alreadyInList: !!existingRow,
            selected: !existingRow && !!email
        });
    });
    rows.sort((a, b) => String(a.name || '').localeCompare(String(b.name || ''), 'de', { sensitivity: 'base' }));
    return rows;
}

/**
 * @param {Array<{code?: string, name?: string, email?: string}>} existingTeachers
 * @param {Array<{selected?: boolean, code?: string, name?: string, email?: string, graphUserId?: string, displayName?: string, userPrincipalName?: string}>} previewRows
 */
export function applyTeacherImportSelection(existingTeachers, previewRows) {
    const out = (Array.isArray(existingTeachers) ? existingTeachers : []).map((t) => ({
        code: normCode(t.code),
        name: normStr(t.name),
        email: normEmail(t.email)
    }));
    const emailIndex = new Map();
    const usedCodes = new Set();
    out.forEach((t, i) => {
        if (t.email) emailIndex.set(t.email, i);
        if (t.code) usedCodes.add(String(t.code).toLowerCase());
    });
    const added = [];
    const updated = [];
    const skipped = [];
    const directoryMatches = {};
    const iso = new Date().toISOString();

    (Array.isArray(previewRows) ? previewRows : []).forEach((row) => {
        if (!row || !row.selected) return;
        const email = normEmail(row.email);
        const name = normStr(row.name || row.displayName);
        if (!email || email.indexOf('@') === -1) {
            skipped.push(row);
            return;
        }
        if (row.graphUserId) {
            directoryMatches[email] = {
                graphUserId: String(row.graphUserId),
                displayName: name,
                userPrincipalName: normStr(row.userPrincipalName),
                notFound: false,
                checkedAt: iso
            };
            const upn = normEmail(row.userPrincipalName);
            if (upn && upn !== email) {
                directoryMatches[upn] = directoryMatches[email];
            }
        }
        if (emailIndex.has(email)) {
            const i = emailIndex.get(email);
            if (name && name !== out[i].name) {
                out[i] = { code: out[i].code, name, email };
                updated.push(out[i]);
            } else {
                skipped.push(row);
            }
            return;
        }
        let code = normCode(row.code);
        if (!code || usedCodes.has(code.toLowerCase())) {
            code = suggestTeacherCode(name, row.givenName, row.surname, usedCodes);
        }
        usedCodes.add(code.toLowerCase());
        const next = { code, name, email };
        emailIndex.set(email, out.length);
        out.push(next);
        added.push(next);
    });

    return { teachers: out, added, updated, skipped, directoryMatches };
}

/**
 * Klasse aus Entra-Feldern raten (Abteilung/Standort), sonst leer.
 * @param {object} user
 */
export function suggestKlasseFromUser(user) {
    const candidates = [user && user.department, user && user.officeLocation];
    for (let i = 0; i < candidates.length; i++) {
        const s = normStr(candidates[i]);
        if (!s) continue;
        const compact = s.replace(/\s+/g, '');
        if (/^(sch[uü]ler(?:innen)?|students?|klasse|class|education|bildung)/i.test(s) && compact.length > 10) {
            continue;
        }
        if (/^[0-9]{1,2}[.\s-]?[A-Za-zÄÖÜäöü][A-Za-z0-9ÄÖÜäöü]{0,10}$/.test(compact)) {
            return compact.replace(/[.\s-]/g, '').toUpperCase();
        }
        if (compact.length <= 8 && !/\s/.test(s) && /[A-Za-z0-9]/.test(compact)) {
            return compact.toUpperCase();
        }
    }
    return '';
}

function directoryMatchPayload(row, iso) {
    const email = normEmail(row.email);
    const name = normStr(row.name || row.displayName);
    if (!row.graphUserId || !email) return {};
    const payload = {
        graphUserId: String(row.graphUserId),
        displayName: name,
        userPrincipalName: normStr(row.userPrincipalName),
        notFound: false,
        checkedAt: iso
    };
    const out = {};
    out[email] = payload;
    const upn = normEmail(row.userPrincipalName);
    if (upn && upn !== email) out[upn] = payload;
    return out;
}

/**
 * @param {object[]} users Graph-User
 * @param {Array<{klasse?: string, name?: string, email?: string}>} existingStudents
 * @param {Map<string, {skuPartNumber?: string}>|null} [skuLookup]
 * @param {{ activeOnly?: boolean, guests?: boolean, families?: string[] }} [opts]
 */
export function buildStudentImportPreview(users, existingStudents, skuLookup, opts) {
    const opt = opts || {};
    const families = Array.isArray(opt.families) ? opt.families : ['a1', 'a3', 'a5'];
    const familySet = new Set(families);
    const existing = Array.isArray(existingStudents) ? existingStudents : [];
    const emailToExisting = new Map();
    existing.forEach((t) => {
        const em = normEmail(t && t.email);
        if (em) emailToExisting.set(em, t);
    });

    const rows = [];
    const seenUser = new Set();
    (Array.isArray(users) ? users : []).forEach((u) => {
        if (!u || !u.id || seenUser.has(u.id)) return;
        if (opt.activeOnly && u.accountEnabled === false) return;
        if (opt.guests === false && String(u.userType || '').toLowerCase() === 'guest') return;
        const sum = summarizeUserLicenses(u, skuLookup);
        if (!sum.hasStudentUserPlan) return;
        const hitFamily = (sum.studentFamilies || []).some((f) => familySet.has(f));
        if (!hitFamily) return;
        seenUser.add(u.id);
        const email = teacherEmailOfUser(u);
        const existingRow = email ? emailToExisting.get(email) : null;
        const guessed = suggestKlasseFromUser(u);
        const klasse = existingRow && existingRow.klasse ? existingRow.klasse : guessed;
        rows.push({
            graphUserId: String(u.id),
            displayName: normStr(u.displayName),
            userPrincipalName: normStr(u.userPrincipalName),
            accountEnabled: u.accountEnabled !== false,
            userType: String(u.userType || 'Member'),
            email,
            klasse,
            name: normStr(u.displayName),
            licenseLabel: sum.primaryLabel,
            studentFamilies: (sum.studentFamilies || []).slice(),
            alreadyInList: !!existingRow,
            selected: !existingRow && !!email
        });
    });
    rows.sort((a, b) => String(a.name || '').localeCompare(String(b.name || ''), 'de', { sensitivity: 'base' }));
    return rows;
}

/**
 * @param {Array<{klasse?: string, name?: string, email?: string}>} existingStudents
 * @param {Array<{selected?: boolean, klasse?: string, name?: string, email?: string, graphUserId?: string, displayName?: string, userPrincipalName?: string}>} previewRows
 */
export function applyStudentImportSelection(existingStudents, previewRows) {
    const out = (Array.isArray(existingStudents) ? existingStudents : []).map((t) => ({
        klasse: normStr(t.klasse),
        name: normStr(t.name),
        email: normEmail(t.email)
    }));
    const emailIndex = new Map();
    out.forEach((t, i) => {
        if (t.email) emailIndex.set(t.email, i);
    });
    const added = [];
    const updated = [];
    const skipped = [];
    const directoryMatches = {};
    const iso = new Date().toISOString();

    (Array.isArray(previewRows) ? previewRows : []).forEach((row) => {
        if (!row || !row.selected) return;
        const email = normEmail(row.email);
        const name = normStr(row.name || row.displayName);
        const klasse = normStr(row.klasse);
        if (!email || email.indexOf('@') === -1) {
            skipped.push(row);
            return;
        }
        Object.assign(directoryMatches, directoryMatchPayload(row, iso));
        if (emailIndex.has(email)) {
            const i = emailIndex.get(email);
            const prev = out[i];
            const nextKlasse = klasse && !prev.klasse ? klasse : prev.klasse;
            const nextName = name && name !== prev.name ? name : prev.name;
            if (nextKlasse !== prev.klasse || nextName !== prev.name) {
                out[i] = { klasse: nextKlasse, name: nextName, email };
                updated.push(out[i]);
            } else {
                skipped.push(row);
            }
            return;
        }
        const next = { klasse, name, email };
        emailIndex.set(email, out.length);
        out.push(next);
        added.push(next);
    });

    return { students: out, added, updated, skipped, directoryMatches };
}

export function countFacultyFamilies(previewRows) {
    const counts = { a1: 0, a3: 0, a5: 0, total: 0, neu: 0, vorhanden: 0 };
    (Array.isArray(previewRows) ? previewRows : []).forEach((r) => {
        counts.total++;
        if (r.alreadyInList) counts.vorhanden++;
        else counts.neu++;
        (r.facultyFamilies || []).forEach((f) => {
            if (counts[f] != null) counts[f]++;
        });
    });
    return counts;
}

/**
 * Freie Lizenzen einer subscribedSku (enabled − consumed).
 * @param {{ prepaidUnits?: { enabled?: number }, consumedUnits?: number }|null} sku
 * @returns {number|null}
 */
export function remainingPrepaidUnits(sku) {
    if (!sku || typeof sku !== 'object') return null;
    const prepaid = sku.prepaidUnits || {};
    const enabled = Number(prepaid.enabled);
    if (!Number.isFinite(enabled)) return null;
    const consumed = Number(sku.consumedUnits);
    return Math.max(0, enabled - (Number.isFinite(consumed) ? consumed : 0));
}

/**
 * SKUs, die einer Person noch zugewiesen werden können.
 * Mit Tenant-subscribedSkus: nur vorhandene Pläne mit freien Sitzen.
 * Ohne Tenant-Liste: Education-Katalog (A1/A3/A5) als Fallback.
 * @param {object[]} subscribedSkus
 * @param {string[]} assignedSkuIds
 * @param {{ fallbackCatalog?: boolean }} [opts]
 */
export function buildAssignableSkuOptions(subscribedSkus, assignedSkuIds, opts) {
    const assigned = new Set(
        (Array.isArray(assignedSkuIds) ? assignedSkuIds : []).map(function (id) {
            return String(id || '').toLowerCase();
        })
    );
    const list = Array.isArray(subscribedSkus) ? subscribedSkus : [];
    const out = [];
    const seen = new Set();

    function pushOpt(id, part, remaining) {
        if (!id || assigned.has(id) || seen.has(id)) return;
        seen.add(id);
        const info = resolveSku(id, part);
        out.push({
            skuId: id,
            skuPartNumber: part || info.skuPartNumber,
            name: info.name,
            shortLabel: info.shortLabel,
            audience: info.audience,
            family: info.family,
            remaining: remaining == null ? null : remaining,
            disabled: remaining === 0
        });
    }

    if (list.length) {
        list.forEach(function (s) {
            const id = String((s && s.skuId) || '').toLowerCase();
            const cap = String((s && s.capabilityStatus) || '').toLowerCase();
            if (cap && cap !== 'enabled' && cap !== 'warning') return;
            const remaining = remainingPrepaidUnits(s);
            if (remaining === 0) return;
            pushOpt(id, s && s.skuPartNumber, remaining);
        });
    } else if (!opts || opts.fallbackCatalog !== false) {
        EDU_SKU_CATALOG.forEach(function (s) {
            if (!s.userPlan) return;
            pushOpt(s.skuId, s.skuPartNumber, null);
        });
    }

    out.sort(function (a, b) {
        const rank = function (x) {
            if (x.audience === 'faculty') return 0;
            if (x.audience === 'student') return 1;
            return 2;
        };
        const d = rank(a) - rank(b);
        if (d) return d;
        const fd = familyRank(b.family) - familyRank(a.family);
        if (fd) return fd;
        return String(a.name || '').localeCompare(String(b.name || ''), 'de', { sensitivity: 'base' });
    });
    return out;
}

const api = {
    EDU_SKU_CATALOG,
    resolveSku,
    facultyUserPlanSkuIds,
    studentUserPlanSkuIds,
    licensesFromAssigned,
    summarizeLicenses,
    summarizeUserLicenses,
    userMatchesLicenseFilter,
    buildLicenseFilterOptions,
    splitPersonName,
    suggestTeacherCode,
    suggestKlasseFromUser,
    teacherEmailOfUser,
    buildTeacherImportPreview,
    applyTeacherImportSelection,
    buildStudentImportPreview,
    applyStudentImportSelection,
    countFacultyFamilies,
    remainingPrepaidUnits,
    buildAssignableSkuOptions
};

export default api;

if (typeof window !== 'undefined') {
    window.ms365GraphLicenses = api;
}
