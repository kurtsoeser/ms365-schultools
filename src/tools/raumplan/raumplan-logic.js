/**
 * Domänenlogik Raumplan.
 */

import { RAUM_TYP, RAUM_TYP_LABEL, normalizeRaum } from './raumplan-model.js';

/**
 * @param {object} draft
 */
export function validateRaum(draft) {
    const errors = [];
    const r = normalizeRaum(draft);
    if (!r.code) errors.push('Code fehlt.');
    if (!r.name) errors.push('Name fehlt.');
    return { ok: errors.length === 0, errors };
}

/**
 * @param {string} typ
 */
export function typLabel(typ) {
    const t = String(typ || '')
        .trim()
        .toLowerCase();
    return RAUM_TYP_LABEL[t] || t || '–';
}

/**
 * @param {object[]} raeume
 * @param {{ q?: string, typ?: string, gebaeude?: string }} filters
 */
export function filterRaeume(raeume, filters = {}) {
    const q = String(filters.q || '')
        .trim()
        .toLowerCase();
    const typ = String(filters.typ || '')
        .trim()
        .toLowerCase();
    const geb = String(filters.gebaeude || '')
        .trim()
        .toLowerCase();
    return (raeume || []).filter((r) => {
        if (!r) return false;
        if (typ && r.typ !== typ) return false;
        if (geb && String(r.gebaeude || '').toLowerCase() !== geb) return false;
        if (q) {
            const hay = (
                r.code +
                ' ' +
                r.name +
                ' ' +
                r.gebaeude +
                ' ' +
                r.stockwerk +
                ' ' +
                r.typ +
                ' ' +
                r.notiz
            ).toLowerCase();
            if (!hay.includes(q)) return false;
        }
        return true;
    });
}

/**
 * @param {object[]} raeume
 */
export function raumStats(raeume) {
    const list = raeume || [];
    const byTyp = {};
    const gebaeude = new Set();
    let kapSum = 0;
    for (const r of list) {
        byTyp[r.typ] = (byTyp[r.typ] || 0) + 1;
        if (r.gebaeude) gebaeude.add(r.gebaeude);
        kapSum += Number(r.kapazitaet) || 0;
    }
    return {
        total: list.length,
        gebaeude: gebaeude.size,
        kapazitaet: kapSum,
        fachraeume: byTyp[RAUM_TYP.FACHRAUM] || 0,
        klassen: byTyp[RAUM_TYP.KLASSE] || 0,
        byTyp
    };
}

/**
 * CSV-Import: Code;Name;Gebaeude;Stockwerk;Kapazitaet;Typ
 * @param {string} text
 */
export function parseRaeumeCsv(text) {
    const lines = String(text || '')
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean);
    /** @type {Array<{ code: string, name: string, gebaeude: string, stockwerk: string, kapazitaet: number, typ: string }>} */
    const out = [];
    for (const line of lines) {
        if (/^code[;,\t]/i.test(line)) continue;
        const parts = line.split(/[;,\t]/).map((p) => p.trim());
        if (parts.length < 2) continue;
        const kap = Number(parts[4] || 0);
        const typRaw = String(parts[5] || RAUM_TYP.SONSTIGES)
            .trim()
            .toLowerCase();
        out.push({
            code: (parts[0] || '').toUpperCase(),
            name: parts[1] || '',
            gebaeude: parts[2] || '',
            stockwerk: parts[3] || '',
            kapazitaet: Number.isFinite(kap) && kap >= 0 ? Math.round(kap) : 0,
            typ: Object.values(RAUM_TYP).includes(typRaw) ? typRaw : RAUM_TYP.SONSTIGES
        });
    }
    return out;
}

/**
 * @param {object[]} raeume
 */
export function raeumeToCsv(raeume) {
    const lines = ['Code;Name;Gebaeude;Stockwerk;Kapazitaet;Typ;Notiz'];
    for (const r of raeume || []) {
        lines.push(
            [r.code, r.name, r.gebaeude, r.stockwerk, r.kapazitaet, r.typ, r.notiz]
                .map(csv)
                .join(';')
        );
    }
    return lines.join('\n');
}

/**
 * Deep-Link zu Technik-Meldungen.
 * @param {string} code
 */
export function technikMeldungUrl(code) {
    const c = String(code || '').trim();
    return 'technik-meldungen.html?raum=' + encodeURIComponent(c);
}

function csv(v) {
    const s = String(v ?? '');
    if (/[;"\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
}
