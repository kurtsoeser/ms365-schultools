/**
 * Regel-Engine für Schularbeiten (SchUG § 17 / LBVO § 7, HAK).
 * Rein, ohne DOM/fetch – Vitest-tauglich.
 */

/** @typedef {'beantragt'|'fixiert'|'abgelehnt'} SaStatus */
/** @typedef {'gesperrt'|'erlaubt'} FensterTyp */

/**
 * @typedef {object} Regelwerk
 * @property {number} maxProTag
 * @property {number} maxProWoche
 * @property {number} ankuendigungsfristTage
 * @property {number} sperreVorNotenkonferenzTage
 */

/**
 * @typedef {object} Terminfenster
 * @property {string} titel
 * @property {FensterTyp} typ
 * @property {string} startdatum ISO YYYY-MM-DD
 * @property {string} enddatum ISO YYYY-MM-DD
 */

/**
 * @typedef {object} Schularbeit
 * @property {string} [schularbeitId]
 * @property {string} fachCode
 * @property {string} klasseCode
 * @property {string} lehrerCode
 * @property {string} datum ISO YYYY-MM-DD
 * @property {number} dauerMinuten
 * @property {string} [semester] WS|SS
 * @property {SaStatus} [status]
 */

/**
 * @typedef {object} ValidateOptions
 * @property {Schularbeit} draft
 * @property {Schularbeit[]} existing
 * @property {Regelwerk} rules
 * @property {Terminfenster[]} windows
 * @property {string} [today] ISO YYYY-MM-DD (Default: lokales Heute)
 * @property {{ proSemester?: number, minDauer?: number, maxDauer?: number }} [fachMeta]
 */

export const DEFAULT_RULES = {
    maxProTag: 1,
    maxProWoche: 2,
    ankuendigungsfristTage: 7,
    sperreVorNotenkonferenzTage: 7
};

/**
 * @param {unknown} value
 * @returns {string|null} YYYY-MM-DD
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

/**
 * @param {string} iso
 * @returns {{ y: number, m: number, d: number }|null}
 */
export function parseIsoDateParts(iso) {
    const s = toIsoDateOnly(iso);
    if (!s) return null;
    const [y, m, d] = s.split('-').map(Number);
    if (!y || !m || !d) return null;
    return { y, m, d };
}

/** UTC-Mittag für DST-sichere Tagesdifferenz. */
function utcNoonMs(iso) {
    const p = parseIsoDateParts(iso);
    if (!p) return NaN;
    return Date.UTC(p.y, p.m - 1, p.d, 12, 0, 0);
}

/**
 * @param {string} fromIso
 * @param {string} toIso
 * @returns {number|null}
 */
export function daysBetween(fromIso, toIso) {
    const a = utcNoonMs(fromIso);
    const b = utcNoonMs(toIso);
    if (Number.isNaN(a) || Number.isNaN(b)) return null;
    return Math.round((b - a) / 86400000);
}

/**
 * ISO-Kalenderwochen-Schlüssel, z. B. 2026-W42
 * @param {string} iso
 */
export function isoWeekKey(iso) {
    const p = parseIsoDateParts(iso);
    if (!p) return '';
    const date = new Date(Date.UTC(p.y, p.m - 1, p.d));
    // Donnerstag der aktuellen Woche bestimmt das ISO-Jahr
    const day = date.getUTCDay() || 7;
    date.setUTCDate(date.getUTCDate() + 4 - day);
    const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
    const week = Math.ceil(((date - yearStart) / 86400000 + 1) / 7);
    const isoYear = date.getUTCFullYear();
    return `${isoYear}-W${String(week).padStart(2, '0')}`;
}

/**
 * @param {string} iso
 * @param {number} deltaDays
 */
export function addDays(iso, deltaDays) {
    const ms = utcNoonMs(iso);
    if (Number.isNaN(ms)) return null;
    const next = new Date(ms + deltaDays * 86400000);
    return toIsoDateOnly(next);
}

/**
 * @param {string} iso
 * @param {Terminfenster} win
 */
export function dateInWindow(iso, win) {
    const d = toIsoDateOnly(iso);
    const a = toIsoDateOnly(win && win.startdatum);
    const b = toIsoDateOnly(win && win.enddatum);
    if (!d || !a || !b) return false;
    return d >= a && d <= b;
}

/**
 * @param {string} iso
 * @param {Terminfenster[]} windows
 * @param {FensterTyp} typ
 */
export function findWindowsCovering(iso, windows, typ) {
    const list = Array.isArray(windows) ? windows : [];
    return list.filter((w) => w && w.typ === typ && dateInWindow(iso, w));
}

/**
 * @param {Schularbeit} sa
 */
function isActiveStatus(sa) {
    const st = String((sa && sa.status) || 'beantragt').toLowerCase();
    return st === 'beantragt' || st === 'fixiert';
}

/**
 * @param {ValidateOptions} opts
 * @returns {{ errors: string[], warnings: string[], canSubmit: boolean }}
 */
export function validateSchularbeit(opts) {
    const errors = [];
    const warnings = [];
    const draft = (opts && opts.draft) || {};
    const rules = { ...DEFAULT_RULES, ...(opts && opts.rules) };
    const windows = Array.isArray(opts && opts.windows) ? opts.windows : [];
    const existing = Array.isArray(opts && opts.existing) ? opts.existing : [];
    const fachMeta = (opts && opts.fachMeta) || {};
    const today = toIsoDateOnly(opts && opts.today) || toIsoDateOnly(new Date());

    const klasse = String(draft.klasseCode || '').trim();
    const fach = String(draft.fachCode || '').trim();
    const datum = toIsoDateOnly(draft.datum);
    const dauer = Number(draft.dauerMinuten);
    const draftId = String(draft.schularbeitId || '').trim();

    if (!klasse) errors.push('Bitte eine Klasse wählen.');
    if (!fach) errors.push('Bitte ein Fach wählen.');
    if (!datum) errors.push('Bitte ein gültiges Datum (JJJJ-MM-TT) angeben.');
    if (!Number.isFinite(dauer) || dauer <= 0) errors.push('Bitte eine gültige Dauer in Minuten angeben.');

    if (errors.length) {
        return { errors, warnings, canSubmit: false };
    }

    const peers = existing.filter((sa) => {
        if (!sa || !isActiveStatus(sa)) return false;
        if (String(sa.klasseCode || '').trim() !== klasse) return false;
        const id = String(sa.schularbeitId || '').trim();
        if (draftId && id && id === draftId) return false;
        return true;
    });

    const sameDay = peers.filter((sa) => toIsoDateOnly(sa.datum) === datum);
    const maxTag = Math.max(1, Number(rules.maxProTag) || 1);
    if (sameDay.length >= maxTag) {
        errors.push(
            `Maximal ${maxTag} Schularbeit${maxTag === 1 ? '' : 'en'} pro Tag und Klasse (bereits ${sameDay.length}).`
        );
    }

    const week = isoWeekKey(datum);
    const sameWeek = peers.filter((sa) => isoWeekKey(sa.datum) === week);
    const maxWoche = Math.max(1, Number(rules.maxProWoche) || 1);
    if (sameWeek.length >= maxWoche) {
        errors.push(
            `Maximal ${maxWoche} Schularbeit${maxWoche === 1 ? '' : 'en'} pro Kalenderwoche und Klasse (bereits ${sameWeek.length} in ${week}).`
        );
    }

    const frist = Math.max(0, Number(rules.ankuendigungsfristTage) || 0);
    const lead = daysBetween(today, datum);
    if (lead != null && lead < frist) {
        errors.push(`Ankündigungsfrist: mindestens ${frist} Tage Vorlauf (aktuell ${lead} Tag${lead === 1 ? '' : 'e'}).`);
    }
    if (lead != null && lead < 0) {
        errors.push('Das Datum liegt in der Vergangenheit.');
    }

    const blocked = findWindowsCovering(datum, windows, 'gesperrt');
    if (blocked.length) {
        const titles = blocked.map((w) => w.titel || 'Sperrzeit').join(', ');
        errors.push(`Termin liegt in einer Sperrzeit: ${titles}.`);
    }

    const prev = addDays(datum, -1);
    if (prev) {
        const prevBlocked = findWindowsCovering(prev, windows, 'gesperrt');
        if (prevBlocked.length) {
            warnings.push(
                'Tag nach schulfreiem Zeitraum (Sperrzeit) – laut LBVO möglichst vermeiden.'
            );
        }
    }

    const proSem = Number(fachMeta.proSemester);
    if (Number.isFinite(proSem) && proSem > 0 && draft.semester) {
        const sem = String(draft.semester).toUpperCase();
        const sameFachSem = peers.filter(
            (sa) =>
                String(sa.fachCode || '').trim() === fach &&
                String(sa.semester || '').toUpperCase() === sem
        );
        if (sameFachSem.length >= proSem) {
            warnings.push(
                `Kontingent Fach/Semester: üblich max. ${proSem} (bereits ${sameFachSem.length}).`
            );
        }
    }

    const minD = Number(fachMeta.minDauer) || 50;
    const maxD = Number(fachMeta.maxDauer) || 150;
    if (Number.isFinite(dauer) && (dauer < minD || dauer > maxD)) {
        warnings.push(`Dauer üblicherweise ${minD}–${maxD} Minuten (aktuell ${dauer}).`);
    }

    return {
        errors,
        warnings,
        canSubmit: errors.length === 0
    };
}

/**
 * KPI-Helfer für Dashboard.
 * @param {Schularbeit[]} items
 * @param {string} [todayIso]
 * @param {Regelwerk} [rules]
 */
export function computeDashboardKpis(items, todayIso, rules) {
    const today = toIsoDateOnly(todayIso) || toIsoDateOnly(new Date());
    const list = Array.isArray(items) ? items : [];
    const inTwoWeeks = addDays(today, 14);
    const rw = { ...DEFAULT_RULES, ...(rules || {}) };

    let offen = 0;
    let fixiertNaechste2Wochen = 0;
    let dieseWoche = 0;
    const weekNow = isoWeekKey(today);

    list.forEach((sa) => {
        const st = String((sa && sa.status) || '').toLowerCase();
        const d = toIsoDateOnly(sa && sa.datum);
        if (st === 'beantragt') offen++;
        if (st === 'fixiert' && d && today && inTwoWeeks && d >= today && d <= inTwoWeeks) {
            fixiertNaechste2Wochen++;
        }
        if ((st === 'beantragt' || st === 'fixiert') && d && isoWeekKey(d) === weekNow) {
            dieseWoche++;
        }
    });

    return {
        offen,
        fixiertNaechste2Wochen,
        konflikte: countRuleConflicts(list, rw),
        dieseWoche
    };
}

/**
 * Anzahl Konflikt-Gruppen (Klasse+Tag oder Klasse+Woche über Limit).
 * @param {Schularbeit[]} items
 * @param {Regelwerk} rules
 */
export function countRuleConflicts(items, rules) {
    const rw = { ...DEFAULT_RULES, ...(rules || {}) };
    const maxTag = Math.max(1, Number(rw.maxProTag) || 1);
    const maxWoche = Math.max(1, Number(rw.maxProWoche) || 1);
    const active = (Array.isArray(items) ? items : []).filter((sa) => {
        const st = String((sa && sa.status) || '').toLowerCase();
        return st === 'beantragt' || st === 'fixiert';
    });

    const byDay = new Map();
    const byWeek = new Map();
    active.forEach((sa) => {
        const klasse = String(sa.klasseCode || '').trim();
        const d = toIsoDateOnly(sa.datum);
        if (!klasse || !d) return;
        const dayKey = klasse + '|' + d;
        const weekKey = klasse + '|' + isoWeekKey(d);
        byDay.set(dayKey, (byDay.get(dayKey) || 0) + 1);
        byWeek.set(weekKey, (byWeek.get(weekKey) || 0) + 1);
    });

    let n = 0;
    byDay.forEach((count) => {
        if (count > maxTag) n++;
    });
    byWeek.forEach((count) => {
        if (count > maxWoche) n++;
    });
    return n;
}

/**
 * Wochenverteilung für Dashboard-Chart.
 * @param {Schularbeit[]} items
 * @param {{ today?: string, weekCount?: number }} [opts]
 * @returns {{ weeks: { key: string, label: string, total: number, byFach: Record<string, number> }[], faecher: string[] }}
 */
export function buildWeeklyDistribution(items, opts) {
    const today = toIsoDateOnly(opts && opts.today) || toIsoDateOnly(new Date());
    const weekCount = Math.max(1, Number(opts && opts.weekCount) || 8);
    const active = (Array.isArray(items) ? items : []).filter((sa) => {
        const st = String((sa && sa.status) || '').toLowerCase();
        return (st === 'beantragt' || st === 'fixiert') && toIsoDateOnly(sa.datum);
    });

    // Anker: Montag der aktuellen ISO-Woche
    const anchorMonday = mondayOfIsoWeek(today);
    const weekKeys = [];
    for (let i = 0; i < weekCount; i++) {
        const monday = addDays(anchorMonday, i * 7);
        const key = isoWeekKey(monday);
        weekKeys.push({ key, monday, label: key.replace(/^\d{4}-/, '') });
    }
    const keySet = new Set(weekKeys.map((w) => w.key));

    /** @type {Map<string, Record<string, number>>} */
    const map = new Map();
    weekKeys.forEach((w) => map.set(w.key, {}));

    const fachSet = new Set();
    active.forEach((sa) => {
        const key = isoWeekKey(sa.datum);
        if (!keySet.has(key)) return;
        const fach = String(sa.fachCode || '').trim() || '?';
        fachSet.add(fach);
        const bucket = map.get(key) || {};
        bucket[fach] = (bucket[fach] || 0) + 1;
        map.set(key, bucket);
    });

    const faecher = Array.from(fachSet).sort((a, b) => a.localeCompare(b, 'de'));
    const weeks = weekKeys.map((w) => {
        const byFach = map.get(w.key) || {};
        let total = 0;
        Object.keys(byFach).forEach((f) => {
            total += byFach[f];
        });
        return { key: w.key, label: w.label, total, byFach };
    });

    return { weeks, faecher };
}

/**
 * @param {string} iso
 */
function mondayOfIsoWeek(iso) {
    const p = parseIsoDateParts(iso);
    if (!p) return iso;
    const dt = new Date(Date.UTC(p.y, p.m - 1, p.d, 12));
    const day = dt.getUTCDay() || 7; // So=7
    dt.setUTCDate(dt.getUTCDate() - (day - 1));
    return toIsoDateOnly(dt);
}
