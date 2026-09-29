/**
 * Naming-/Mail-Nick-/Schuljahr-Helfer fuer „Schulstruktur-Sync".
 *
 * Schuljahr: kanonisch Sep–Aug via `shared/utils/school-year.js`.
 * Mail-Nick: Umlaute via `shared/utils/mail-nickname.js`.
 */

import {
    parseSchoolYearStartYear,
    currentSchoolYearLabel,
    nextSchoolYearLabel
} from '../../shared/utils/school-year.js';
import { normalizeMailNickname } from '../../shared/utils/mail-nickname.js';

export { parseSchoolYearStartYear, currentSchoolYearLabel, nextSchoolYearLabel };

/**
 * Berechnet die aktuelle Schulstufe aus Abschlussjahr und Schuljahr.
 *
 * @param {string|number} gradYear        4-stelliges Abschlussjahr.
 * @param {string} schoolYearLabel        z. B. `"2025/26"`.
 * @param {number} maxStufen              Anzahl Schulstufen (1..12, Default 5).
 * @returns {number} Schulstufe oder `NaN` bei ungueltigen Eingaben.
 */
export function gradeFromGraduationYear(gradYear, schoolYearLabel, maxStufen) {
    const gy = String(gradYear || '').trim();
    const sy = parseSchoolYearStartYear(schoolYearLabel);
    const gyi = /^\d{4}$/.test(gy) ? parseInt(gy, 10) : NaN;
    if (!isFinite(gyi) || !isFinite(sy)) return NaN;
    const ms = isFinite(maxStufen) ? Math.max(1, Math.min(12, Math.round(maxStufen))) : 5;
    /*
     * Abschlussjahr = Ende der hoechsten Stufe. In Schuljahr sy/sy+1 gilt:
     *   grade = (maxStufen + 1) - (abschlussjahr - sy)
     */
    return (ms + 1) - (gyi - sy);
}

/**
 * Ersetzt die fuehrende 1–2-stellige Zahl in `label` durch `nextGrade`
 * (z. B. `"1A"` → `"2A"`). Labels ohne fuehrende Zahl bleiben unveraendert.
 *
 * @param {string} label
 * @param {number} nextGrade
 */
export function replaceLeadingNumber(label, nextGrade) {
    const s = String(label || '').trim();
    if (!s) return s;
    const g = String(Math.round(nextGrade));
    if (/^\d{1,2}/.test(s)) return s.replace(/^\d{1,2}/, g);
    return s;
}

/** Erlaubt nur `[A-Za-z0-9-]`, behaelt Gross-/Kleinschreibung. */
export function normNickPart(s) {
    return String(s || '').trim().replace(/[^A-Za-z0-9-]/g, '');
}

/**
 * Normalisiert ein Praefix als lower-case, `[a-z0-9]`-only.
 * Liefert `fallback` (selbst normalisiert), wenn `s` leer wird.
 */
export function normNickPrefixLower(s, fallback) {
    const t = String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    return t || String(fallback || '').trim() || '';
}

/** Liefert `s` in UPPER- oder lower-case, je nach `upper`-Flag. */
export function maybeUpperByFlag(s, upper) {
    const v = String(s || '');
    return upper ? v.toUpperCase() : v.toLowerCase();
}

/**
 * Setzt `{yearPrefix} | {klasse} | {fach}`-Platzhalter im Template ein.
 * Default-Template wird verwendet, wenn `tpl` leer ist.
 */
export function buildKursteamNameFromTemplate(tpl, ctx) {
    const template = String(tpl || '').trim() || '{yearPrefix} | {klasse} | {fach}';
    return template
        .replaceAll('{yearPrefix}', String(ctx.yearPrefix || ''))
        .replaceAll('{klasse}', String(ctx.klasse || ''))
        .replaceAll('{fach}', String(ctx.fach || ''))
        .replaceAll('{gruppe}', String(ctx.gruppe || ''));
}

/**
 * Wie {@link buildKursteamNameFromTemplate}, aber liefert einen
 * Mail-Nick-tauglichen Slug (`a-z0-9-`).
 */
export function buildKursteamMailNickFromTemplate(tpl, ctx) {
    const template = String(tpl || '').trim() || 'kt-{yearPrefix}-{klasse}-{fach}';
    const raw = template
        .replaceAll('{yearPrefix}', String(ctx.yearPrefix || ''))
        .replaceAll('{klasse}', String(ctx.klasse || ''))
        .replaceAll('{fach}', String(ctx.fach || ''))
        .replaceAll('{gruppe}', String(ctx.gruppe || ''));
    return normalizeMailNickname(raw);
}

/**
 * Mail-Nick fuer Jahrgang. Schema: `<prefix><year>[-<suffix>]`,
 * z. B. `jg2025-A`.
 */
export function buildJgMailNick(schema, year, suffix) {
    const prefix = normNickPrefixLower(schema?.jgPrefix, 'jg');
    const suf = maybeUpperByFlag(normNickPart(suffix), !!schema?.jgUpper);
    const y = String(year || '').trim().replace(/[^0-9]/g, '').slice(0, 4);
    const sep = suf ? '-' : '';
    return (prefix + y + sep + suf).replace(/[^A-Za-z0-9-]/g, '');
}

/**
 * Mail-Nick fuer Arbeitsgemeinschaft. Schema: `<prefix>[-<shortCode>]`,
 * z. B. `arge-FUSS`.
 */
export function buildArgeMailNick(schema, shortCode) {
    const prefix = normNickPrefixLower(schema?.argePrefix, 'arge');
    const code = maybeUpperByFlag(normNickPart(shortCode), !!schema?.argeUpper);
    const sep = code ? '-' : '';
    return (prefix + sep + code).replace(/[^A-Za-z0-9-]/g, '');
}

/**
 * Generischer Slug aus einem Bezeichner (Umlaute → ae/oe/ue/ss).
 */
export function buildMailNickFromLabel(label) {
    return normalizeMailNickname(label);
}

/**
 * Leitet einen Mail-Nick aus dem Local-Part eines UPN/Mail-Adresse ab.
 * Fallback: zufaelliger Wert `u<hex>`. Max. 64 Zeichen.
 */
export function mailNicknameFromUpn(upn) {
    const u = String(upn || '').trim().toLowerCase();
    const local = (u.split('@')[0] || '').trim();
    let nick = buildMailNickFromLabel(local);
    if (!nick) nick = 'u' + String(Math.random()).toString(16).slice(2, 10);
    if (nick.length > 64) nick = nick.slice(0, 64);
    return nick;
}

/**
 * Generiert ein temporaeres Graph-API-Passwort:
 * 14 alphanumerische Zeichen + 1 Sonderzeichen + Garantie-Suffix `1aA`
 * (erfuellt Standard-Komplexitaetsregeln).
 */
export function generateGraphTempPassword() {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
    const sym = '!@#$%';
    let s = '';
    for (let i = 0; i < 14; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
    return s + sym.charAt(Math.floor(Math.random() * sym.length)) + '1aA';
}
