/**
 * Validierung / Hilfen für Schulaktivitäten – ohne DOM/fetch.
 */
import {
    STATUS_CHOICES,
    TEILNAHME_MODUS_CHOICES,
    VORLAUF_MODUS_CHOICES,
    GENEHMIGUNG_PFAD_CHOICES
} from './schulaktivitaeten-planer-schema.js';
import { timesFromUnterrichtseinheiten, parseUeNumber } from './schulaktivitaeten-planer-zeitraster.js';

/** @typedef {'beantragt'|'kv_stellungnahme'|'nachforderung'|'genehmigt'|'abgelehnt'} AktStatus */
/** @typedef {'klasse'|'klassen'|'gruppe'|'schueler'} TeilnahmeModus */
/** @typedef {'direktion'|'kv_dann_direktion'} GenehmigungPfad */
/** @typedef {'sperre'|'warnung'|'ausnahme'} VorlaufModus */

/**
 * @typedef {object} Aktivitaet
 * @property {string} [aktivitaetId]
 * @property {string} titel
 * @property {string} typ
 * @property {string} [klasseCode]
 * @property {string} [klasseCodes]
 * @property {TeilnahmeModus|string} [teilnahmeModus]
 * @property {string} [gruppenText]
 * @property {string} [teilnehmendeText]
 * @property {string} [nichtTeilnehmende]
 * @property {string} lehrerCode
 * @property {string} startdatum
 * @property {string} enddatum
 * @property {number|string} [startUe]
 * @property {number|string} [endUe]
 * @property {string} [startZeit]
 * @property {string} [endZeit]
 * @property {AktStatus} [status]
 */

/**
 * @typedef {object} Regelwerk
 * @property {number} minVorlaufTage
 * @property {number} maxGleichzeitigProKlasse
 * @property {GenehmigungPfad|string} [genehmigungPfad]
 * @property {VorlaufModus|string} [vorlaufModus]
 */

export const DEFAULT_RULES = {
    minVorlaufTage: 7,
    maxGleichzeitigProKlasse: 1,
    genehmigungPfad: 'direktion',
    vorlaufModus: 'sperre',
    /** @type {{ typ: string, okBis: number, warnBis: number }[]} */
    kostenLimits: []
};

/** Orientierungskosten je Veranstaltungstyp (€ pro Teilnehmer) – nur Ampel, kein Payment. */
export const DEFAULT_KOSTEN_LIMITS = [
    { typ: 'Projekttage', okBis: 200, warnBis: 300 },
    { typ: 'Skikurs', okBis: 300, warnBis: 400 },
    { typ: 'Sportwoche', okBis: 400, warnBis: 500 },
    { typ: 'Wien-Tage', okBis: 200, warnBis: 300 },
    { typ: 'Sprachwoche', okBis: 600, warnBis: 1000 },
    { typ: 'Exkursion', okBis: 50, warnBis: 100 },
    { typ: 'Abschlusstage', okBis: 200, warnBis: 300 }
];

/**
 * @param {unknown} rules
 * @returns {{ typ: string, okBis: number, warnBis: number }[]}
 */
export function resolveKostenLimits(rules) {
    const custom = rules && Array.isArray(rules.kostenLimits) ? rules.kostenLimits : [];
    if (custom.length) {
        return custom
            .map((r) => ({
                typ: String((r && r.typ) || '').trim(),
                okBis: Number(r && r.okBis) || 0,
                warnBis: Number(r && r.warnBis) || 0
            }))
            .filter((r) => r.typ);
    }
    return DEFAULT_KOSTEN_LIMITS.map((r) => ({ ...r }));
}

/**
 * @param {string} typ
 * @param {number|string} betrag
 * @param {unknown} [rules]
 * @returns {'ok'|'warn'|'teuer'|'unknown'}
 */
export function classifyKostenAmpel(typ, betrag, rules) {
    const n = Number(betrag);
    if (!Number.isFinite(n) || n < 0) return 'unknown';
    const limits = resolveKostenLimits(rules);
    const t = String(typ || '').trim().toLowerCase();
    const row =
        limits.find((l) => l.typ.toLowerCase() === t) ||
        limits.find((l) => t.includes(l.typ.toLowerCase()) || l.typ.toLowerCase().includes(t));
    if (!row) return 'unknown';
    if (n < row.okBis) return 'ok';
    if (n <= row.warnBis) return 'warn';
    return 'teuer';
}

/**
 * Lehrkräfte vorschlagen, deren Klassenkürzel zu den gewählten Klassen passt
 * (Stammdaten: teacher.klassen / classes / classCodes – best effort).
 * @param {string[]|string} klasseCodes
 * @param {Array<{ code?: string, email?: string, name?: string, klassen?: string[], classes?: string[], classCodes?: string }>} teachers
 * @returns {string[]} teacher codes
 */
export function suggestTeachersForKlassen(klasseCodes, teachers) {
    const wanted = new Set(parseCodeList(klasseCodes).map((c) => c.toUpperCase()));
    if (!wanted.size || !Array.isArray(teachers)) return [];
    const out = [];
    teachers.forEach((t) => {
        if (!t || !t.code) return;
        const raw = []
            .concat(t.klassen || [])
            .concat(t.classes || [])
            .concat(parseCodeList(t.classCodes));
        const hit = raw.some((c) => wanted.has(String(c || '').trim().toUpperCase()));
        if (hit) out.push(String(t.code).trim());
    });
    return out.filter((c, i, a) => c && a.indexOf(c) === i);
}

/**
 * @param {unknown} raw
 * @returns {GenehmigungPfad}
 */
export function normalizeGenehmigungPfad(raw) {
    const s = String(raw || '')
        .trim()
        .toLowerCase();
    return GENEHMIGUNG_PFAD_CHOICES.includes(s) ? /** @type {GenehmigungPfad} */ (s) : 'direktion';
}

/**
 * @param {unknown} raw
 * @returns {VorlaufModus}
 */
export function normalizeVorlaufModus(raw) {
    const s = String(raw || '')
        .trim()
        .toLowerCase();
    return VORLAUF_MODUS_CHOICES.includes(s) ? /** @type {VorlaufModus} */ (s) : 'sperre';
}

/**
 * Initialstatus nach Einreichen – abhängig vom Schul-Genehmigungspfad.
 * @param {Regelwerk|object} [rules]
 */
export function initialStatusForSubmit(rules) {
    return normalizeGenehmigungPfad(rules && rules.genehmigungPfad) === 'kv_dann_direktion'
        ? 'kv_stellungnahme'
        : 'beantragt';
}

/**
 * @param {unknown} value
 * @returns {string[]}
 */
export function parseCodeList(value) {
    return String(value || '')
        .split(/[;,\n]+/)
        .map((s) => s.trim())
        .filter(Boolean);
}

/**
 * @param {string[]|string} value
 */
export function joinCodeList(value) {
    if (Array.isArray(value)) {
        return value
            .map((s) => String(s || '').trim())
            .filter(Boolean)
            .join('; ');
    }
    return parseCodeList(value).join('; ');
}

/**
 * @param {Aktivitaet|object} akt
 * @returns {string[]}
 */
export function klasseCodesOf(akt) {
    const multi = parseCodeList(akt && akt.klasseCodes);
    if (multi.length) return multi;
    const one = String((akt && akt.klasseCode) || '').trim();
    return one ? [one] : [];
}

/**
 * @param {Aktivitaet|object} akt
 */
export function normalizeTeilnahme(akt) {
    const raw = String((akt && akt.teilnahmeModus) || '').toLowerCase();
    const modus = TEILNAHME_MODUS_CHOICES.includes(raw) ? raw : 'klasse';
    const codes = klasseCodesOf(akt);
    const klasseCodes = joinCodeList(codes);
    const klasseCode = codes[0] || '';
    return {
        teilnahmeModus: /** @type {TeilnahmeModus} */ (modus),
        klasseCode,
        klasseCodes,
        gruppenText: String((akt && akt.gruppenText) || '').trim(),
        teilnehmendeText: String((akt && akt.teilnehmendeText) || '').trim(),
        nichtTeilnehmende: String((akt && akt.nichtTeilnehmende) || '').trim()
    };
}

/**
 * @param {string} prev
 * @param {string} entry
 * @param {string} [iso]
 */
export function appendAenderungsLog(prev, entry, iso) {
    const line = `${String(iso || new Date().toISOString()).slice(0, 16).replace('T', ' ')} · ${String(entry || '').trim()}`;
    const old = String(prev || '').trim();
    return old ? `${old}\n${line}` : line;
}

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
    if (!titel) errors.push('Bitte gib der Aktivität einen Titel.');

    const typ = String(draft.typ || '').trim();
    if (!typ) errors.push('Bitte wähle die Art der Aktivität.');

    const teilnahme = normalizeTeilnahme(draft);
    const klassen = klasseCodesOf(teilnahme);
    if (teilnahme.teilnahmeModus === 'gruppe') {
        if (!teilnahme.gruppenText) errors.push('Bitte benenne die Unterrichtsgruppe.');
        if (!klassen.length) warnings.push('Für Gruppen bitte die betroffenen Klassen mit angeben.');
    } else if (teilnahme.teilnahmeModus === 'schueler') {
        if (!teilnahme.teilnehmendeText) errors.push('Bitte nenne die teilnehmenden Schüler:innen.');
        if (!klassen.length) warnings.push('Für Einzelpersonen bitte die Klassen mit angeben.');
    } else if (!klassen.length) {
        errors.push('Bitte wähle mindestens eine Klasse.');
    }

    if (teilnahme.teilnahmeModus === 'klassen' && klassen.length < 2) {
        warnings.push('Mehrere Klassen gewählt – idealerweise mindestens zwei Klassen.');
    }

    const start = toIsoDateOnly(draft.startdatum);
    const end = toIsoDateOnly(draft.enddatum) || start;
    if (!start) errors.push('Bitte gib an, wann die Aktivität beginnt.');
    if (start && end && utcNoonMs(end) < utcNoonMs(start)) {
        errors.push('Das Enddatum darf nicht vor dem Startdatum liegen.');
    }

    if (start && today) {
        const vorlauf = daysBetween(today, start);
        const minTage = Number(rules.minVorlaufTage || 0);
        if (vorlauf != null && minTage > 0 && vorlauf < minTage) {
            const msg =
                `Dieser Termin liegt innerhalb der ${minTage}-Tage-Frist. ` +
                `Für Schulaktivitäten gilt normalerweise ein Vorlauf von mindestens ${minTage} Tagen.`;
            const modus = normalizeVorlaufModus(rules.vorlaufModus);
            if (modus === 'warnung') {
                warnings.push(msg);
            } else if (modus === 'ausnahme') {
                const note = String(draft.vorlaufAusnahmeGrund || draft.notiz || '').trim();
                if (!note) {
                    errors.push(msg + ' Bitte begründe die Ausnahme unter Hinweise.');
                } else {
                    warnings.push(msg + ' Ausnahme mit Begründung akzeptiert.');
                }
            } else {
                errors.push(msg);
            }
        }
    }

    if (String(draft.typ || '') === 'Klassenabwesenheit' && !klassen.length) {
        errors.push('Bei Klassenabwesenheit bitte die betroffene(n) Klasse(n) angeben.');
    }

    const status = String(draft.status || 'beantragt').toLowerCase();
    if (status && !STATUS_CHOICES.includes(status)) {
        warnings.push('Unbekannter Status – erwartete Werte: ' + STATUS_CHOICES.join(', '));
    }

    const draftId = String(draft.aktivitaetId || '').trim();
    if (klassen.length && start && end) {
        const max = Number(rules.maxGleichzeitigProKlasse);
        if (Number.isFinite(max) && max > 0) {
            for (const klasse of klassen) {
                const relevant = existing.filter((e) => {
                    const otherKlassen = klasseCodesOf(e);
                    if (!otherKlassen.includes(klasse)) return false;
                    const st = String(e.status || '').toLowerCase();
                    if (st === 'abgelehnt') return false;
                    if (draftId && String(e.aktivitaetId || '') === draftId) return false;
                    return rangesOverlap(start, end, e.startdatum, e.enddatum || e.startdatum);
                });
                if (relevant.length >= max) {
                    const msg = `In Klasse ${klasse} überschneiden sich bereits ${relevant.length} Aktivität(en) (Limit ${max}).`;
                    if (status === 'genehmigt') errors.push(msg);
                    else warnings.push(msg);
                }
            }
        }
    }

    const ort = String(draft.ort || '').trim();
    if (!ort) warnings.push('Ein Ort oder Ziel hilft bei der Genehmigung.');

    const ueCheck = timesFromUnterrichtseinheiten({
        startUe: draft.startUe,
        endUe: draft.endUe,
        raster: opts && opts.zeitraster
    });
    if (!ueCheck.ok && ueCheck.error) errors.push(ueCheck.error);
    else if (parseUeNumber(draft.startUe) == null && !String(draft.startZeit || '').trim()) {
        warnings.push('Wann findet die Aktivität statt? Unterrichtseinheiten oder Uhrzeit helfen bei der Planung.');
    }

    return { errors, warnings, ok: errors.length === 0, teilnahme, ue: ueCheck };
}

/**
 * Menschliche Vorlauf-Anzeige am Datumsfeld (ohne Debug-Formulierung).
 * @param {object} draft
 * @param {object} [rules]
 * @param {string} [today]
 */
export function vorlaufStatusForDraft(draft, rules, today) {
    const r = { ...DEFAULT_RULES, ...(rules || {}) };
    const minTage = Number(r.minVorlaufTage || 0);
    const start = toIsoDateOnly(draft && draft.startdatum);
    const day = toIsoDateOnly(today || new Date()) || '';
    if (!minTage || !start || !day) {
        return { kind: 'none', minTage, message: '' };
    }
    const vorlauf = daysBetween(day, start);
    if (vorlauf == null) return { kind: 'none', minTage, message: '' };
    if (vorlauf < minTage) {
        return {
            kind: 'warn',
            minTage,
            vorlauf,
            message:
                `Dieser Termin liegt innerhalb der ${minTage}-Tage-Frist. ` +
                `Für Schulaktivitäten gilt normalerweise ein Vorlauf von mindestens ${minTage} Tagen.`
        };
    }
    return {
        kind: 'ok',
        minTage,
        vorlauf,
        message: 'Der Termin kann regulär beantragt werden.'
    };
}

/**
 * Fehler → Wizard-Schritt + Kurzlabel für handlungsorientierte Review.
 * @param {string} msg
 * @returns {{ step: 1|2|3|4, label: string, action: string }}
 */
export function mapValidationIssueToStep(msg) {
    const m = String(msg || '').toLowerCase();
    if (m.includes('titel') || m.includes('art der') || m.includes('typ') || m.includes('ort') || m.includes('datum') || m.includes('beginn') || m.includes('frist') || m.includes('vorlauf')) {
        return { step: 1, label: m.includes('frist') || m.includes('vorlauf') ? 'Termin prüfen' : 'Aktivität ergänzen', action: 'Aktivität ergänzen →' };
    }
    if (m.includes('klasse') || m.includes('teilnehm') || m.includes('gruppe')) {
        return { step: 2, label: 'Teilnahme ergänzen', action: 'Teilnahme ergänzen →' };
    }
    if (m.includes('zeit') || m.includes('unterrichtseinheit') || m.includes('begleit')) {
        return { step: 3, label: 'Organisation ergänzen', action: 'Organisation ergänzen →' };
    }
    return { step: 4, label: 'Angaben prüfen', action: 'Angaben prüfen →' };
}

/**
 * Österreichisches Schuljahr: Startjahr (z. B. 2026 → 2026/27, 1.9.–31.8.).
 * @param {number|string|Date} [ref]
 */
export function schoolYearStartFromDate(ref) {
    const d = ref instanceof Date ? ref : ref ? new Date(String(ref)) : new Date();
    if (Number.isNaN(d.getTime())) {
        const now = new Date();
        return now.getMonth() + 1 >= 9 ? now.getFullYear() : now.getFullYear() - 1;
    }
    const month = d.getMonth() + 1;
    const year = d.getFullYear();
    return month >= 9 ? year : year - 1;
}

/**
 * @param {number|string} startYear
 */
export function schoolYearLabel(startYear) {
    const y = Number(startYear);
    if (!Number.isFinite(y)) return '–';
    return `${y}/${String(y + 1).slice(-2)}`;
}

/**
 * @param {number|string} startYear
 * @returns {{ from: string, to: string }}
 */
export function schoolYearRange(startYear) {
    const y = Number(startYear);
    return { from: `${y}-09-01`, to: `${y + 1}-08-31` };
}

/**
 * @param {string} iso
 * @param {number|string} startYear
 */
export function isoInSchoolYear(iso, startYear) {
    const start = toIsoDateOnly(iso) || '';
    if (!start) return false;
    const { from, to } = schoolYearRange(startYear);
    return start >= from && start <= to;
}

/**
 * Einfache Auswertungen für Jahresübersicht / Statistik.
 * @param {Aktivitaet[]} items
 * @param {{ year?: number, schoolYearStart?: number|string }} [opts]
 */
export function computeAktivitaetStats(items, opts) {
    const list = Array.isArray(items) ? items : [];
    const year = opts && opts.year != null ? Number(opts.year) : null;
    const schoolYearStart =
        opts && opts.schoolYearStart != null ? Number(opts.schoolYearStart) : null;
    const filtered = list.filter((it) => {
        const start = toIsoDateOnly(it.startdatum) || '';
        if (schoolYearStart != null && Number.isFinite(schoolYearStart)) {
            return isoInSchoolYear(start, schoolYearStart);
        }
        if (!year) return true;
        return start.startsWith(String(year) + '-');
    });
    /** @type {Record<string, number>} */
    const byTyp = {};
    /** @type {Record<string, number>} */
    const byStatus = {};
    /** @type {Record<string, number>} */
    const byMonth = {};
    /** @type {Record<string, number>} */
    const byKlasse = {};
    let mitNichtteilnehmern = 0;
    let untisOffen = 0;
    filtered.forEach((it) => {
        const typ = String(it.typ || 'Sonstiges');
        byTyp[typ] = (byTyp[typ] || 0) + 1;
        const st = String(it.status || 'beantragt').toLowerCase();
        byStatus[st] = (byStatus[st] || 0) + 1;
        const start = toIsoDateOnly(it.startdatum) || '';
        const month = start.length >= 7 ? start.slice(0, 7) : 'unbekannt';
        byMonth[month] = (byMonth[month] || 0) + 1;
        klasseCodesOf(it).forEach((k) => {
            byKlasse[k] = (byKlasse[k] || 0) + 1;
        });
        if (String(it.nichtTeilnehmende || '').trim()) mitNichtteilnehmern++;
        if (st === 'genehmigt' && !it.untisErledigt) untisOffen++;
    });
    return {
        gesamt: filtered.length,
        byTyp,
        byStatus,
        byMonth,
        byKlasse,
        mitNichtteilnehmern,
        untisOffen,
        year: year || null,
        schoolYearStart: schoolYearStart || null,
        schoolYearLabel: schoolYearStart != null ? schoolYearLabel(schoolYearStart) : null
    };
}

/**
 * @param {Aktivitaet[]} items
 * @param {string} [today]
 */
export function computeDashboardKpis(items, today) {
    const list = Array.isArray(items) ? items : [];
    const t = toIsoDateOnly(today || new Date()) || '';
    let offen = 0;
    let nachforderung = 0;
    let genehmigt = 0;
    let abgelehnt = 0;
    let demnaechst = 0;
    let untisOffen = 0;
    list.forEach((it) => {
        const st = String(it.status || '').toLowerCase();
        if (st === 'beantragt') offen++;
        else if (st === 'nachforderung') nachforderung++;
        else if (st === 'genehmigt') {
            genehmigt++;
            if (!it.untisErledigt) untisOffen++;
        } else if (st === 'abgelehnt') abgelehnt++;
        if (st === 'genehmigt' && t && it.startdatum) {
            const d = daysBetween(t, it.startdatum);
            if (d != null && d >= 0 && d <= 14) demnaechst++;
        }
    });
    return {
        offen,
        nachforderung,
        genehmigt,
        abgelehnt,
        demnaechst,
        untisOffen,
        gesamt: list.length
    };
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

/**
 * Kürzel aus E-Mail ableiten, wenn Stammdaten kein Match liefern.
 * @param {string} email
 */
export function deriveLehrerCodeFromEmail(email) {
    const local = String(email || '')
        .trim()
        .toLowerCase()
        .split('@')[0] || '';
    const cleaned = local.replace(/[^a-z0-9]/gi, '');
    if (!cleaned) return 'LEHR';
    return cleaned.slice(0, 12).toUpperCase();
}

/**
 * Ownership für SharePoint erzwingen (BeantragtVon / LehrerEmail / LehrerCode).
 * Ohne diese Felder bleibt „Meine Aktivitäten“ leer.
 *
 * @param {object} akt
 * @param {{ accountEmail?: string, teachers?: Array<{ code?: string, email?: string, upn?: string }> }} [ctx]
 * @returns {{ beantragtVon: string, lehrerEmail: string, lehrerCode: string }}
 */
export function resolveAktOwnership(akt, ctx) {
    const a = akt && typeof akt === 'object' ? akt : {};
    const teachers = Array.isArray(ctx && ctx.teachers) ? ctx.teachers : [];
    const account = String((ctx && ctx.accountEmail) || '')
        .trim()
        .toLowerCase();

    let beantragtVon = String(a.beantragtVon || '')
        .trim()
        .toLowerCase();
    let lehrerEmail = String(a.lehrerEmail || '')
        .trim()
        .toLowerCase();
    let lehrerCode = String(a.lehrerCode || '').trim();

    if (!beantragtVon) beantragtVon = account || lehrerEmail;
    if (!lehrerEmail) lehrerEmail = account || beantragtVon;

    const byEmail =
        lehrerEmail &&
        teachers.find((t) => {
            const em = String((t && t.email) || '')
                .trim()
                .toLowerCase();
            const upn = String((t && t.upn) || '')
                .trim()
                .toLowerCase();
            return em === lehrerEmail || upn === lehrerEmail;
        });
    const byCode =
        lehrerCode &&
        teachers.find(
            (t) =>
                String((t && t.code) || '')
                    .trim()
                    .toUpperCase() === lehrerCode.toUpperCase()
        );

    if (!lehrerCode && byEmail && byEmail.code) {
        lehrerCode = String(byEmail.code).trim();
    }
    if (byCode && byCode.email && !String(a.lehrerEmail || '').trim()) {
        lehrerEmail = String(byCode.email).trim().toLowerCase();
    }
    if (!lehrerCode && lehrerEmail) {
        lehrerCode = deriveLehrerCodeFromEmail(lehrerEmail);
    }
    if (!beantragtVon) beantragtVon = lehrerEmail;
    if (!lehrerEmail) lehrerEmail = beantragtVon;

    return {
        beantragtVon: beantragtVon || '',
        lehrerEmail: lehrerEmail || '',
        lehrerCode: lehrerCode || ''
    };
}

/**
 * @param {{ beantragtVon?: string, lehrerEmail?: string, lehrerCode?: string }} ownership
 */
export function assertAktOwnership(ownership) {
    const email = String((ownership && (ownership.lehrerEmail || ownership.beantragtVon)) || '')
        .trim()
        .toLowerCase();
    if (!email || email.indexOf('@') < 0) {
        throw new Error(
            'Lehrer-E-Mail fehlt im Antrag (Anmeldung nicht erkannt). Bitte Seite neu laden oder ab-/anmelden und erneut einreichen.'
        );
    }
    if (!String((ownership && ownership.beantragtVon) || '').trim()) {
        throw new Error('Feld „Beantragt von“ fehlt – bitte erneut einreichen.');
    }
    if (!String((ownership && ownership.lehrerCode) || '').trim()) {
        throw new Error('Lehrer-Kürzel fehlt – bitte erneut einreichen.');
    }
}
