/**
 * Validierung / Hilfen für Schulaktivitäten – ohne DOM/fetch.
 */

/** @typedef {'beantragt'|'genehmigt'|'abgelehnt'} AktStatus */

/**
 * @typedef {object} Aktivitaet
 * @property {string} [aktivitaetId]
 * @property {string} titel
 * @property {string} typ
 * @property {string} klasseCode
 * @property {string} lehrerCode
 * @property {string} startdatum
 * @property {string} enddatum
 * @property {AktStatus} [status]
 */

/**
 * @typedef {object} Regelwerk
 * @property {number} minVorlaufTage
 * @property {number} maxGleichzeitigProKlasse
 */

export const DEFAULT_RULES = {
    minVorlaufTage: 7,
    maxGleichzeitigProKlasse: 1
};

/**
 * @param {unknown} value
 * @returns {string|null}
 */
export function toIsoDateOnly(value) {
    if (value == null) return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        const y = value.getFullYear();
        const m = String(value.getMonth() + 1).padStart(2, '0');
        const d = String(value.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }
    const s = String(value).trim();
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
    if (!m) return null;
    return `${m[1]}-${m[2]}-${m[3]}`;
}

function utcNoonMs(iso) {
    const s = toIsoDateOnly(iso);
    if (!s) return NaN;
    const [y, m, d] = s.split('-').map(Number);
    return Date.UTC(y, m - 1, d, 12, 0, 0);
}

/**
 * @param {string} fromIso
 * @param {string} toIso
 */
export function daysBetween(fromIso, toIso) {
    const a = utcNoonMs(fromIso);
    const b = utcNoonMs(toIso);
    if (Number.isNaN(a) || Number.isNaN(b)) return null;
    return Math.round((b - a) / 86400000);
}

/**
 * Inklusiver Datumsbereich überschneidet sich?
 * @param {string} a0
 * @param {string} a1
 * @param {string} b0
 * @param {string} b1
 */
export function rangesOverlap(a0, a1, b0, b1) {
    const A0 = utcNoonMs(a0);
    const A1 = utcNoonMs(a1 || a0);
    const B0 = utcNoonMs(b0);
    const B1 = utcNoonMs(b1 || b0);
    if ([A0, A1, B0, B1].some((n) => Number.isNaN(n))) return false;
    return A0 <= B1 && B0 <= A1;
}

/**
 * @param {object} opts
 * @param {Aktivitaet} opts.draft
 * @param {Aktivitaet[]} opts.existing
 * @param {Regelwerk} [opts.rules]
 * @param {string} [opts.today]
 */
export function validateAktivitaet(opts) {
    const draft = (opts && opts.draft) || {};
    const existing = Array.isArray(opts && opts.existing) ? opts.existing : [];
    const rules = { ...DEFAULT_RULES, ...((opts && opts.rules) || {}) };
    const today = toIsoDateOnly((opts && opts.today) || new Date()) || '';
    /** @type {string[]} */
    const errors = [];
    /** @type {string[]} */
    const warnings = [];

    const titel = String(draft.titel || '').trim();
    if (!titel) errors.push('Bitte einen Titel angeben.');

    const typ = String(draft.typ || '').trim();
    if (!typ) errors.push('Bitte einen Typ wählen.');

    const klasse = String(draft.klasseCode || '').trim();
    if (!klasse) errors.push('Bitte eine Klasse wählen.');

    const start = toIsoDateOnly(draft.startdatum);
    const end = toIsoDateOnly(draft.enddatum) || start;
    if (!start) errors.push('Bitte ein Startdatum angeben.');
    if (start && end && utcNoonMs(end) < utcNoonMs(start)) {
        errors.push('Das Enddatum darf nicht vor dem Startdatum liegen.');
    }

    if (start && today) {
        const vorlauf = daysBetween(today, start);
        if (vorlauf != null && vorlauf < Number(rules.minVorlaufTage || 0)) {
            errors.push(
                `Mindest-Vorlauf ${rules.minVorlaufTage} Tag(e) vor Beginn (Antrag heute → Start in ${vorlauf} Tag(en)).`
            );
        }
    }

    const status = String(draft.status || 'beantragt').toLowerCase();
    const draftId = String(draft.aktivitaetId || '').trim();
    if (klasse && start && end) {
        const relevant = existing.filter((e) => {
            if (String(e.klasseCode || '').trim() !== klasse) return false;
            const st = String(e.status || '').toLowerCase();
            if (st === 'abgelehnt') return false;
            if (draftId && String(e.aktivitaetId || '') === draftId) return false;
            return rangesOverlap(start, end, e.startdatum, e.enddatum || e.startdatum);
        });
        const max = Number(rules.maxGleichzeitigProKlasse);
        if (Number.isFinite(max) && max > 0 && relevant.length >= max) {
            const msg = `In dieser Klasse überschneiden sich bereits ${relevant.length} Aktivität(en) (Limit ${max}).`;
            if (status === 'genehmigt') errors.push(msg);
            else warnings.push(msg);
        }
    }

    const ort = String(draft.ort || '').trim();
    if (!ort) warnings.push('Ort/Ziel ist leer – für die Genehmigung meist hilfreich.');

    return { errors, warnings, ok: errors.length === 0 };
}

/**
 * @param {Aktivitaet[]} items
 * @param {string} [today]
 */
export function computeDashboardKpis(items, today) {
    const list = Array.isArray(items) ? items : [];
    const t = toIsoDateOnly(today || new Date()) || '';
    let offen = 0;
    let genehmigt = 0;
    let abgelehnt = 0;
    let demnaechst = 0;
    list.forEach((it) => {
        const st = String(it.status || '').toLowerCase();
        if (st === 'beantragt') offen++;
        else if (st === 'genehmigt') genehmigt++;
        else if (st === 'abgelehnt') abgelehnt++;
        if (st === 'genehmigt' && t && it.startdatum) {
            const d = daysBetween(t, it.startdatum);
            if (d != null && d >= 0 && d <= 14) demnaechst++;
        }
    });
    return { offen, genehmigt, abgelehnt, demnaechst, gesamt: list.length };
}

/**
 * @param {number} year
 * @param {number} month1to12
 * @returns {string[]} ISO dates covering calendar grid (Mon-start weeks)
 */
export function monthGridDates(year, month1to12) {
    const first = new Date(Date.UTC(year, month1to12 - 1, 1, 12));
    let dow = first.getUTCDay(); // 0 Sun
    dow = dow === 0 ? 6 : dow - 1; // Mon=0
    const start = new Date(first);
    start.setUTCDate(first.getUTCDate() - dow);
    const out = [];
    for (let i = 0; i < 42; i++) {
        const d = new Date(start);
        d.setUTCDate(start.getUTCDate() + i);
        const y = d.getUTCFullYear();
        const m = String(d.getUTCMonth() + 1).padStart(2, '0');
        const day = String(d.getUTCDate()).padStart(2, '0');
        out.push(`${y}-${m}-${day}`);
    }
    return out;
}

/**
 * @param {Aktivitaet} item
 * @param {string} dayIso
 */
export function itemCoversDay(item, dayIso) {
    const s = toIsoDateOnly(item && item.startdatum);
    const e = toIsoDateOnly((item && item.enddatum) || s);
    if (!s || !dayIso) return false;
    return rangesOverlap(s, e || s, dayIso, dayIso);
}
