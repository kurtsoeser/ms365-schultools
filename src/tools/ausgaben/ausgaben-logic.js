/**
 * Domänenlogik Ausgabenverwaltung.
 */

import {
    EMPFAENGER_TYP,
    VORGANG_STATUS,
    normalizeGegenstand,
    normalizeTyp,
    normalizeVorgang,
    toIsoDateOnly
} from './ausgaben-model.js';

/**
 * @param {object} draft
 */
export function validateGegenstand(draft) {
    const errors = [];
    const g = normalizeGegenstand(draft);
    if (!g.kennung) errors.push('Kennung fehlt.');
    if (!g.typ) errors.push('Typ fehlt.');
    return { ok: errors.length === 0, errors };
}

/**
 * @param {object} draft
 */
export function validateVorgang(draft) {
    const errors = [];
    const v = normalizeVorgang(draft);
    if (!v.gegenstandId) errors.push('Gegenstand fehlt.');
    if (!v.empfaengerName) errors.push('Empfänger-Name fehlt.');
    if (!Object.values(EMPFAENGER_TYP).includes(v.empfaengerTyp)) {
        errors.push('Empfänger-Typ ungültig.');
    }
    if (!Object.values(VORGANG_STATUS).includes(v.status)) {
        errors.push('Status ungültig.');
    }
    return { ok: errors.length === 0, errors };
}

/**
 * @param {object} draft
 */
export function validateTyp(draft) {
    const errors = [];
    const t = normalizeTyp(draft);
    if (!t.id) errors.push('Typ-ID fehlt.');
    if (!t.label) errors.push('Bezeichnung fehlt.');
    return { ok: errors.length === 0, errors };
}

/**
 * Offener Vorgang für Gegenstand?
 * @param {string} gegenstandId
 * @param {object[]} vorgaenge
 * @param {string} [ignoreId]
 */
export function findOpenVorgang(gegenstandId, vorgaenge, ignoreId) {
    const gid = String(gegenstandId || '').trim();
    const ignore = String(ignoreId || '').trim();
    return (vorgaenge || []).find(
        (v) =>
            v &&
            v.gegenstandId === gid &&
            v.id !== ignore &&
            v.status === VORGANG_STATUS.AUSGEGEBEN
    );
}

/**
 * @param {object[]} gegenstaende
 * @param {object[]} vorgaenge
 * @param {{ typ?: string, q?: string, nurAktiv?: boolean, nurVerfuegbar?: boolean }} filters
 */
export function filterGegenstaende(gegenstaende, vorgaenge, filters = {}) {
    const typ = String(filters.typ || '')
        .trim()
        .toLowerCase();
    const q = String(filters.q || '')
        .trim()
        .toLowerCase();
    return (gegenstaende || []).filter((g) => {
        if (!g) return false;
        if (filters.nurAktiv && !g.aktiv) return false;
        if (typ && g.typ !== typ) return false;
        if (q) {
            const hay = (g.kennung + ' ' + g.bezeichnung + ' ' + g.ortZone + ' ' + g.typLabel)
                .toLowerCase();
            if (!hay.includes(q)) return false;
        }
        if (filters.nurVerfuegbar) {
            if (findOpenVorgang(g.id, vorgaenge)) return false;
        }
        return true;
    });
}

/**
 * @param {object[]} vorgaenge
 * @param {{ status?: string, typ?: string, q?: string, ueberfaellig?: boolean, day?: string, meEmail?: string, admin?: boolean, showAll?: boolean }} filters
 * @param {object[]} [gegenstaende]
 */
export function filterVorgaenge(vorgaenge, filters = {}, gegenstaende = []) {
    const status = String(filters.status || '')
        .trim()
        .toLowerCase();
    const typ = String(filters.typ || '')
        .trim()
        .toLowerCase();
    const q = String(filters.q || '')
        .trim()
        .toLowerCase();
    const me = String(filters.meEmail || '')
        .trim()
        .toLowerCase();
    // explizit admin:false → nur eigene; admin:true → showAll default an; ohne Flag → Legacy ohne Scope
    const canSeeAll =
        filters.admin === true
            ? filters.showAll !== false
            : filters.admin === false
              ? false
              : true;
    const today = toIsoDateOnly(filters.day || new Date());
    const byId = new Map((gegenstaende || []).map((g) => [g.id, g]));

    return (vorgaenge || []).filter((v) => {
        if (!v) return false;
        if (status && v.status !== status) return false;
        if (typ) {
            const g = byId.get(v.gegenstandId);
            if (!g || g.typ !== typ) return false;
        }
        if (!canSeeAll) {
            if (!me || String(v.empfaengerEmail || '').toLowerCase() !== me) return false;
        }
        if (q) {
            const hay = (
                v.kennung +
                ' ' +
                v.empfaengerName +
                ' ' +
                v.empfaengerEmail +
                ' ' +
                v.klasseCode
            ).toLowerCase();
            if (!hay.includes(q)) return false;
        }
        if (filters.ueberfaellig) {
            if (v.status !== VORGANG_STATUS.AUSGEGEBEN) return false;
            if (!v.faelligAm || v.faelligAm >= today) return false;
        }
        return true;
    });
}

/**
 * @param {object[]} gegenstaende
 * @param {object[]} vorgaenge
 */
export function inventarStats(gegenstaende, vorgaenge) {
    const list = gegenstaende || [];
    const aktiv = list.filter((g) => g.aktiv).length;
    const offen = (vorgaenge || []).filter((v) => v.status === VORGANG_STATUS.AUSGEGEBEN);
    const heute = toIsoDateOnly(new Date());
    const ueberfaellig = offen.filter((v) => v.faelligAm && v.faelligAm < heute).length;
    const verfuegbar = list.filter((g) => g.aktiv && !findOpenVorgang(g.id, vorgaenge)).length;
    return {
        total: list.length,
        aktiv,
        ausgegeben: offen.length,
        verfuegbar,
        ueberfaellig
    };
}

/**
 * Einfaches Übergabeprotokoll (Text).
 * @param {object} gegenstand
 * @param {object} vorgang
 */
export function buildUebergabeText(gegenstand, vorgang) {
    const g = normalizeGegenstand(gegenstand || {});
    const v = normalizeVorgang(vorgang || {});
    return [
        'Übergabeprotokoll – Ausgabenverwaltung',
        '--------------------------------------',
        'Gegenstand: ' + (g.bezeichnung || g.kennung) + ' (' + (g.typLabel || g.typ) + ')',
        'Kennung: ' + g.kennung,
        'Ort/Zone: ' + (g.ortZone || '–'),
        '',
        'Empfänger: ' + v.empfaengerName + ' (' + v.empfaengerTyp + ')',
        'E-Mail: ' + (v.empfaengerEmail || '–'),
        'Klasse: ' + (v.klasseCode || '–'),
        'Ausgegeben am: ' + v.ausgegebenAm,
        'Fällig am: ' + (v.faelligAm || '–'),
        'Kaution (Hinweis): ' + (v.kautionHinweis || '–'),
        'Unterschrift Empfang: ' + (v.unterschriftEmpfang ? 'ja' : 'nein'),
        'Notiz: ' + (v.notiz || '–'),
        'Ausgegeben von: ' + (v.ausgegebenVon || '–'),
        '',
        'Unterschrift Empfänger: __________________'
    ].join('\n');
}

/**
 * CSV Inventar.
 * @param {object[]} gegenstaende
 */
export function gegenstaendeToCsv(gegenstaende) {
    const lines = ['Typ;TypLabel;Kennung;Bezeichnung;OrtZone;Aktiv;Notiz'];
    for (const g of gegenstaende || []) {
        lines.push(
            [g.typ, g.typLabel, g.kennung, g.bezeichnung, g.ortZone, g.aktiv ? 'ja' : 'nein', g.notiz]
                .map(csv)
                .join(';')
        );
    }
    return lines.join('\n');
}

/**
 * CSV-Import: Typ;Kennung;Bezeichnung;OrtZone
 * @param {string} text
 */
export function parseGegenstaendeCsv(text) {
    const lines = String(text || '')
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean);
    /** @type {Array<{ typ: string, kennung: string, bezeichnung: string, ortZone: string }>} */
    const out = [];
    for (const line of lines) {
        if (/^typ[;,\t]/i.test(line)) continue;
        const parts = line.split(/[;,\t]/).map((p) => p.trim());
        if (parts.length < 2) continue;
        out.push({
            typ: (parts[0] || 'sonstiges').toLowerCase(),
            kennung: parts[1] || '',
            bezeichnung: parts[2] || '',
            ortZone: parts[3] || ''
        });
    }
    return out;
}

function csv(v) {
    const s = String(v ?? '');
    if (/[;"\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
}
