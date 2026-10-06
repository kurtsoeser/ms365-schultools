/**
 * Validierung / Genehmigungspfad für Freistellungen – ohne DOM/fetch.
 */
import { MULTI_DAY_THRESHOLD, STATUS_CHOICES, KATEGORIE_CHOICES } from './freistellung-planer-schema.js';
import { isAllowedKategorie } from './freistellung-planer-kategorien.js';
import { itemMatchesJahrgangClassCodes } from './freistellung-planer-jahrgang-scope.js';

/**
 * @typedef {object} FreistellungDraft
 * @property {string} [titel]
 * @property {string} [schuelerName]
 * @property {string} klasse
 * @property {string} beginn
 * @property {string} ende
 * @property {string} [kategorie]
 * @property {string} [beschreibung]
 * @property {string} [kvEmail]
 * @property {string} [status]
 */

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
 * Inklusive Kalendertage (Beginn und Ende zählen).
 * @param {string} fromIso
 * @param {string} toIso
 * @returns {number|null}
 */
export function inclusiveDayCount(fromIso, toIso) {
    const a = utcNoonMs(fromIso);
    const b = utcNoonMs(toIso || fromIso);
    if (Number.isNaN(a) || Number.isNaN(b)) return null;
    if (b < a) return null;
    return Math.round((b - a) / 86400000) + 1;
}

/**
 * @param {string} beginn
 * @param {string} ende
 * @param {number} [threshold]
 */
export function isMultiDay(beginn, ende, threshold) {
    const days = inclusiveDayCount(beginn, ende);
    const t = Number.isFinite(threshold) ? Number(threshold) : MULTI_DAY_THRESHOLD;
    return days != null && days >= t;
}

/**
 * Genehmigungspfad laut Produktlogik:
 * - 1 Tag: Klassenvorstand
 * - mehrtägig: sequentiell KV → Direktion
 * (Der importierte PA-Flow genehmigt derzeit immer sequentiell KV→Direktion;
 *  die UI/Logik kennzeichnet den gewünschten Pfad klar.)
 *
 * @param {string} beginn
 * @param {string} ende
 * @returns {{ multiDay: boolean, dayCount: number|null, steps: string[], label: string }}
 */
export function approvalPath(beginn, ende) {
    const dayCount = inclusiveDayCount(beginn, ende);
    const multiDay = dayCount != null && dayCount >= MULTI_DAY_THRESHOLD;
    if (multiDay) {
        return {
            multiDay: true,
            dayCount,
            steps: ['Klassenvorstand', 'Direktion'],
            label: 'Sequentiell: Klassenvorstand → Direktion'
        };
    }
    return {
        multiDay: false,
        dayCount,
        steps: ['Klassenvorstand'],
        label: 'Klassenvorstand'
    };
}

/**
 * @param {object} opts
 * @param {FreistellungDraft} opts.draft
 * @param {string} [opts.today]
 */
export function validateFreistellung(opts) {
    const draft = (opts && opts.draft) || {};
    const today = toIsoDateOnly((opts && opts.today) || new Date()) || '';
    /** @type {string[]} */
    const errors = [];
    /** @type {string[]} */
    const warnings = [];

    const name = String(draft.schuelerName || draft.titel || '').trim();
    if (!name) errors.push('Bitte den Namen der Schülerin / des Schülers angeben.');

    const klasse = String(draft.klasse || '').trim();
    if (!klasse) errors.push('Bitte eine Klasse wählen.');

    const beginn = toIsoDateOnly(draft.beginn);
    const ende = toIsoDateOnly(draft.ende) || beginn;
    if (!beginn) errors.push('Bitte ein Beginndatum angeben.');
    if (beginn && ende && utcNoonMs(ende) < utcNoonMs(beginn)) {
        errors.push('Das Endedatum darf nicht vor dem Beginn liegen.');
    }

    if (beginn && today && utcNoonMs(beginn) < utcNoonMs(today)) {
        warnings.push('Beginn liegt in der Vergangenheit – nur sinnvoll bei nachträglicher Dokumentation.');
    }

    const kat = String(draft.kategorie || '').trim();
    const allowedKat = (draft && draft._allowedKategorien) || KATEGORIE_CHOICES;
    if (!kat) errors.push('Bitte eine Kategorie wählen.');
    else if (!isAllowedKategorie(kat, allowedKat)) {
        errors.push('Kategorie ist ungültig.');
    }

    const beschreibung = String(draft.beschreibung || '').trim();
    if (!beschreibung) warnings.push('Eine kurze Begründung erleichtert die Genehmigung.');

    const kvEmail = String(draft.kvEmail || '')
        .trim()
        .toLowerCase();
    if (!kvEmail || !kvEmail.includes('@')) {
        errors.push('Klassenvorstand (E-Mail) fehlt – bitte Klasse mit KV in den Stammdaten pflegen oder manuell setzen.');
    }

    const path = approvalPath(beginn || '', ende || beginn || '');
    if (path.multiDay) {
        warnings.push(
            `Mehrtägige Freistellung (${path.dayCount} Tage): Genehmigung durch Klassenvorstand und Direktion.`
        );
    }

    const status = String(draft.status || 'Ausstehend').trim();
    if (status && !STATUS_CHOICES.includes(status)) {
        warnings.push('Unbekannter Status – erwartete Werte: ' + STATUS_CHOICES.join(', '));
    }

    return {
        errors,
        warnings,
        ok: errors.length === 0,
        path,
        dayCount: path.dayCount
    };
}

/**
 * @param {Array<{ status?: string, beginn?: string, ende?: string }>} items
 * @param {string} [today]
 */
export function computeDashboardKpis(items, today) {
    const list = Array.isArray(items) ? items : [];
    const now = toIsoDateOnly(today || new Date()) || '';
    let ausstehend = 0;
    let genehmigt = 0;
    let abgelehnt = 0;
    let mehrtage = 0;
    let demnaechst = 0;

    list.forEach((it) => {
        const st = String(it.status || '').trim();
        if (st === 'Ausstehend') ausstehend++;
        else if (st === 'Genehmigt') genehmigt++;
        else if (st === 'Abgelehnt') abgelehnt++;
        if (isMultiDay(it.beginn, it.ende)) mehrtage++;
        const start = toIsoDateOnly(it.beginn);
        if (start && now) {
            const a = utcNoonMs(now);
            const b = utcNoonMs(start);
            const diff = Math.round((b - a) / 86400000);
            if (diff >= 0 && diff <= 14) demnaechst++;
        }
    });

    return {
        gesamt: list.length,
        ausstehend,
        genehmigt,
        abgelehnt,
        mehrtage,
        demnaechst
    };
}

/**
 * @param {Array<object>} items
 * @param {object} filters
 * @param {object} [scope]
 */
export function filterFreistellungen(items, filters, scope) {
    const f = filters || {};
    const sc = scope || {};
    const list = Array.isArray(items) ? items : [];
    return list.filter((it) => {
        if (sc.onlyMine && sc.accountEmail) {
            const author = String(it.authorEmail || it.beantragtVon || '')
                .trim()
                .toLowerCase();
            if (author !== String(sc.accountEmail).toLowerCase()) return false;
        }
        if (sc.jahrgangOrKv && sc.accountEmail) {
            const kv = String(it.kvEmail || '')
                .trim()
                .toLowerCase();
            const kvHit = kv === String(sc.accountEmail).toLowerCase();
            const jgHit = itemMatchesJahrgangClassCodes(sc.jahrgangClassCodes, it.klasse);
            if (!kvHit && !jgHit) return false;
        } else if (sc.onlyKv && sc.accountEmail) {
            const kv = String(it.kvEmail || '')
                .trim()
                .toLowerCase();
            if (kv !== String(sc.accountEmail).toLowerCase()) return false;
        } else if (sc.jahrgangClassCodes && sc.jahrgangClassCodes.size) {
            if (!itemMatchesJahrgangClassCodes(sc.jahrgangClassCodes, it.klasse)) return false;
        }
        if (f.klasse && String(it.klasse || '') !== String(f.klasse)) return false;
        if (f.status && String(it.status || '') !== String(f.status)) return false;
        if (f.kategorie && String(it.kategorie || '') !== String(f.kategorie)) return false;
        if (f.multiDay === '1' && !isMultiDay(it.beginn, it.ende)) return false;
        if (f.multiDay === '0' && isMultiDay(it.beginn, it.ende)) return false;
        if (f.q) {
            const q = String(f.q).trim().toLowerCase();
            const hay = [it.titel, it.schuelerName, it.klasse, it.beschreibung, it.kategorie, it.kvEmail]
                .map((x) => String(x || '').toLowerCase())
                .join(' ');
            if (!hay.includes(q)) return false;
        }
        return true;
    });
}

export { STATUS_CHOICES, KATEGORIE_CHOICES, MULTI_DAY_THRESHOLD };
