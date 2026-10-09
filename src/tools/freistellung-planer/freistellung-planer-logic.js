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

/**
 * Lokales Datum+Uhrzeit für &lt;input type="datetime-local"&gt;: `YYYY-MM-DDTHH:mm`.
 * @param {unknown} value
 * @returns {string|null}
 */
export function toIsoDateTimeLocal(value) {
    if (value == null || value === '') return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return formatLocalDateTimeParts(value);
    }
    const s = String(value).trim();
    const local = /^(\d{4}-\d{2}-\d{2})T(\d{2}):(\d{2})/.exec(s);
    if (local && !/[zZ]|[+-]\d{2}:?\d{2}$/.test(s)) {
        return local[1] + 'T' + local[2] + ':' + local[3];
    }
    if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
        const d = new Date(s);
        if (!Number.isNaN(d.getTime())) return formatLocalDateTimeParts(d);
    }
    const dateOnly = toIsoDateOnly(s);
    return dateOnly ? dateOnly + 'T00:00' : null;
}

/**
 * @param {Date} d
 */
function formatLocalDateTimeParts(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const h = String(d.getHours()).padStart(2, '0');
    const min = String(d.getMinutes()).padStart(2, '0');
    return y + '-' + m + '-' + day + 'T' + h + ':' + min;
}

/**
 * @param {unknown} iso
 * @returns {boolean}
 */
export function hasClockTime(iso) {
    const local = toIsoDateTimeLocal(iso);
    if (!local || !local.includes('T')) return false;
    const t = local.split('T')[1] || '';
    return t !== '00:00';
}

/**
 * Millisekunden für Vergleich (lokale Wanduhr bei `YYYY-MM-DDTHH:mm`).
 * @param {unknown} iso
 */
export function toDateTimeMs(iso) {
    const local = toIsoDateTimeLocal(iso);
    if (!local) return NaN;
    const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(local);
    if (!m) return NaN;
    return new Date(
        Number(m[1]),
        Number(m[2]) - 1,
        Number(m[3]),
        Number(m[4]),
        Number(m[5]),
        0,
        0
    ).getTime();
}

/**
 * SharePoint Graph dateTime-Feld (ohne Zeitzonen-Suffix → lokale Wandzeit).
 * @param {unknown} value
 * @returns {string|null}
 */
export function toSharePointDateTime(value) {
    const local = toIsoDateTimeLocal(value);
    if (!local) return null;
    return local.length === 16 ? local + ':00' : local;
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

    const beginnDt = toIsoDateTimeLocal(draft.beginn);
    const endeDt = toIsoDateTimeLocal(draft.ende) || beginnDt;
    const beginn = toIsoDateOnly(beginnDt);
    const ende = toIsoDateOnly(endeDt) || beginn;
    if (!beginnDt) errors.push('Bitte Beginn (Datum und Uhrzeit) angeben.');
    if (beginnDt && endeDt && toDateTimeMs(endeDt) < toDateTimeMs(beginnDt)) {
        errors.push('Ende darf nicht vor dem Beginn liegen.');
    }

    if (beginn && today && utcNoonMs(beginn) < utcNoonMs(today)) {
        warnings.push('Beginn liegt in der Vergangenheit – nur sinnvoll bei nachträglicher Dokumentation.');
    }

    if (
        beginnDt &&
        endeDt &&
        beginn === ende &&
        hasClockTime(beginnDt) &&
        toDateTimeMs(endeDt) > toDateTimeMs(beginnDt)
    ) {
        const hours = Math.round(((toDateTimeMs(endeDt) - toDateTimeMs(beginnDt)) / 3600000) * 10) / 10;
        if (hours > 0 && hours < 24) {
            warnings.push('Stundenweise Freistellung (' + hours + ' Std.) – Genehmigung durch Klassenvorstand.');
        }
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
 * @param {string} [raw]
 * @returns {string}
 */
export function normalizeFreistellungClassCode(raw) {
    const s = String(raw || '').trim();
    if (!s) return '';
    const m = /\b(\d{1,2}[A-Za-zÄÖÜäöü]{1,4})\b/.exec(s);
    if (m) return m[1].toUpperCase();
    return s.toUpperCase();
}

/**
 * @param {string} a
 * @param {string} b
 */
export function personNamesLooselyMatch(a, b) {
    const norm = (s) =>
        String(s || '')
            .trim()
            .toLowerCase()
            .replace(/\s+/g, ' ');
    const aa = norm(a);
    const bb = norm(b);
    if (!aa || !bb) return false;
    if (aa === bb) return true;
    const tokens = (s) =>
        s
            .replace(/[,;]/g, ' ')
            .split(/\s+/)
            .map((t) => t.trim())
            .filter(Boolean)
            .sort();
    const ta = tokens(aa);
    const tb = tokens(bb);
    if (ta.length >= 2 && tb.length >= 2 && ta.join(' ') === tb.join(' ')) return true;
    return false;
}

/**
 * @param {object} item
 * @param {string} accountEmail
 * @param {string} [accountName]
 */
export function freistellungItemHasKvAccount(item, accountEmail, accountName) {
    const acc = String(accountEmail || '').trim().toLowerCase();
    if (!acc) return false;
    const kv = String(item && item.kvEmail ? item.kvEmail : '')
        .trim()
        .toLowerCase();
    if (kv && kv === acc) return true;
    const accName = String(accountName || '').trim();
    const kvName = String(item && item.kvName ? item.kvName : '').trim();
    if (accName && kvName && personNamesLooselyMatch(accName, kvName)) return true;
    return false;
}

/**
 * Klassen, in denen dieses Konto auf mindestens einem Antrag als KV steht → alle Anträge dieser Klasse anzeigen.
 * @param {object[]} items
 * @param {string} accountEmail
 * @param {string} [accountName]
 * @returns {Set<string>}
 */
export function deriveKvClassCodesFromFreistellungItems(items, accountEmail, accountName) {
    const codes = new Set();
    const list = Array.isArray(items) ? items : [];
    for (let i = 0; i < list.length; i++) {
        const it = list[i];
        if (!freistellungItemHasKvAccount(it, accountEmail, accountName)) continue;
        const code = normalizeFreistellungClassCode(it.klasse);
        if (code) codes.add(code);
    }
    return codes;
}

/**
 * @param {Set<string>|undefined} classCodes
 * @param {string} itemKlasse
 */
export function freistellungClassCodeInSet(classCodes, itemKlasse) {
    if (!classCodes || !classCodes.size) return false;
    const ik = normalizeFreistellungClassCode(itemKlasse);
    if (!ik) return false;
    for (const c of classCodes) {
        if (normalizeFreistellungClassCode(c) === ik) return true;
    }
    return false;
}

/**
 * @param {object} item
 * @param {object} scope onlyKv, accountEmail, accountName, kvClassCodes, kvResolver(klasse)->email
 */
export function freistellungMatchesKvScope(item, scope) {
    const sc = scope || {};
    const account = String(sc.accountEmail || '')
        .trim()
        .toLowerCase();
    if (!account || !item) return false;
    const kv = String(item.kvEmail || '')
        .trim()
        .toLowerCase();
    if (kv && kv === account) return true;
    if (sc.kvClassCodes && sc.kvClassCodes.size && freistellungClassCodeInSet(sc.kvClassCodes, item.klasse)) {
        return true;
    }
    if (typeof sc.kvResolver === 'function' && item.klasse) {
        const em = String(sc.kvResolver(item.klasse) || '')
            .trim()
            .toLowerCase();
        if (em && em === account) return true;
    }
    const accName = String(sc.accountName || '').trim();
    const kvName = String(item.kvName || '').trim();
    if (accName && kvName && personNamesLooselyMatch(accName, kvName)) return true;
    return false;
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
            if (!freistellungMatchesKvScope(it, sc)) return false;
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

/**
 * @param {object} item
 * @param {string} dayIso YYYY-MM-DD
 */
export function itemCoversDay(item, dayIso) {
    const day = toIsoDateOnly(dayIso);
    if (!day || !item) return false;
    const b = toIsoDateOnly(item.beginn);
    const e = toIsoDateOnly(item.ende) || b;
    if (!b) return false;
    return day >= b && day <= e;
}

/**
 * @param {number} year
 * @param {number} month 1–12
 * @returns {string[]}
 */
export function monthGridDates(year, month) {
    const first = new Date(year, month - 1, 1);
    let dow = first.getDay();
    if (dow === 0) dow = 7;
    const start = new Date(year, month - 1, 1 - (dow - 1));
    const out = [];
    for (let i = 0; i < 42; i++) {
        const d = new Date(start.getFullYear(), start.getMonth(), start.getDate() + i);
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        out.push(`${y}-${m}-${day}`);
    }
    return out;
}

export { STATUS_CHOICES, KATEGORIE_CHOICES, MULTI_DAY_THRESHOLD };
