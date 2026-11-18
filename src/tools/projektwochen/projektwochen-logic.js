/**
 * Reine Geschäftslogik Projektwochen (ohne DOM/fetch) – Vitest.
 */

export const STATUS_ANGEBOT = ['entwurf', 'beantragt', 'freigegeben', 'abgelehnt', 'abgesagt'];
/** Formular/Schema: inkl. ganztags (Anzeige im Plan über PLAN_SLOTS gespannt). */
export const SLOTS = ['ganztags', 'vormittag', 'nachmittag', 'abend'];
/** Wochenplan-Spalten (ohne ganztags – ganztägige Angebote spannen alle drei). */
export const PLAN_SLOTS = ['vormittag', 'nachmittag', 'abend'];
export const TAGE = ['Mo', 'Di', 'Mi', 'Do', 'Fr'];
export const KATEGORIEN = ['exkursion', 'workshop', 'kultur', 'sport', 'sonstiges'];

/**
 * @param {unknown} v
 * @returns {string} YYYY-MM-DD oder ''
 */
export function toIsoDateOnly(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date && !Number.isNaN(v.getTime())) {
        const y = v.getFullYear();
        const m = String(v.getMonth() + 1).padStart(2, '0');
        const d = String(v.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + d;
    }
    const s = String(v).trim();
    const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : '';
}

/**
 * @param {unknown} v
 * @returns {string} YYYY-MM-DDTHH:mm oder ''
 */
export function toDatetimeLocalValue(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date && !Number.isNaN(v.getTime())) {
        const pad = (n) => String(n).padStart(2, '0');
        return (
            v.getFullYear() +
            '-' +
            pad(v.getMonth() + 1) +
            '-' +
            pad(v.getDate()) +
            'T' +
            pad(v.getHours()) +
            ':' +
            pad(v.getMinutes())
        );
    }
    const s = String(v).trim();
    const m = s.match(/^(\d{4}-\d{2}-\d{2})[T ](\d{2}:\d{2})/);
    if (m) return m[1] + 'T' + m[2];
    const d = toIsoDateOnly(s);
    return d ? d + 'T08:00' : '';
}

/**
 * @param {string} isoDate
 * @returns {'Mo'|'Di'|'Mi'|'Do'|'Fr'|''}
 */
export function weekdayLabelDeFromIso(isoDate) {
    const d = toIsoDateOnly(isoDate);
    if (!d) return '';
    const dt = new Date(d + 'T12:00:00');
    if (Number.isNaN(dt.getTime())) return '';
    return TAGE[dt.getDay() === 0 ? 6 : dt.getDay() - 1] || '';
}

/**
 * @param {string} a
 * @param {string} b
 */
export function daysBetween(a, b) {
    const da = toIsoDateOnly(a);
    const db = toIsoDateOnly(b);
    if (!da || !db) return null;
    const t0 = new Date(da + 'T12:00:00').getTime();
    const t1 = new Date(db + 'T12:00:00').getTime();
    if (Number.isNaN(t0) || Number.isNaN(t1)) return null;
    return Math.round((t1 - t0) / 86400000);
}

/**
 * Effektiver Buchungsstart: Angebot.BuchungAb oder Aktion.BuchungAbDefault.
 * @param {{ buchungAb?: string }} angebot
 * @param {{ buchungAbDefault?: string }|null} aktion
 */
export function effectiveBuchungAb(angebot, aktion) {
    const fromOffer = String((angebot && angebot.buchungAb) || '').trim();
    if (fromOffer) return fromOffer;
    return String((aktion && aktion.buchungAbDefault) || '').trim();
}

/**
 * @param {string} buchungAb
 * @param {Date} [now]
 */
export function isBookingOpen(buchungAb, now) {
    const raw = String(buchungAb || '').trim();
    if (!raw) return true;
    const t = Date.parse(raw.length === 10 ? raw + 'T00:00:00' : raw);
    if (Number.isNaN(t)) return true;
    const n = now instanceof Date ? now.getTime() : Date.now();
    return n >= t;
}

/**
 * @param {object} opts
 * @param {object} opts.draft
 * @param {object[]} [opts.existing]
 * @param {object|null} [opts.aktion]
 * @param {string[]} [opts.classCodes]
 * @param {Date} [opts.now]
 * @returns {{ errors: string[], warnings: string[], canSubmit: boolean }}
 */
export function validateAngebot(opts) {
    const draft = (opts && opts.draft) || {};
    const existing = Array.isArray(opts && opts.existing) ? opts.existing : [];
    const aktion = (opts && opts.aktion) || null;
    const classCodes = Array.isArray(opts && opts.classCodes) ? opts.classCodes : [];
    const errors = [];
    const warnings = [];

    const title = String(draft.title || draft.titel || '').trim();
    if (!title) errors.push('Titel fehlt.');

    const datum = toIsoDateOnly(draft.datum);
    if (!datum) errors.push('Datum fehlt.');

    const kap = Number(draft.kapazitaet);
    if (!Number.isFinite(kap) || kap < 1) errors.push('Kapazität muss mindestens 1 sein.');
    else if (kap > 60) warnings.push('Kapazität über 60 – bitte prüfen.');

    const preis = Number(draft.preisEuro);
    if (draft.preisEuro !== '' && draft.preisEuro != null && (!Number.isFinite(preis) || preis < 0)) {
        errors.push('Preis darf nicht negativ sein.');
    }

    if (!String(draft.lehrerCode || '').trim() && !String(draft.lehrerEmail || '').trim()) {
        errors.push('Verantwortliche Lehrkraft (Kürzel oder E-Mail) fehlt.');
    }

    const start = String(draft.startzeit || '').trim();
    const end = String(draft.endzeit || '').trim();
    if (start && end && start >= end) errors.push('Endzeit muss nach der Startzeit liegen.');

    if (aktion && aktion.startdatum && aktion.enddatum && datum) {
        if (datum < toIsoDateOnly(aktion.startdatum) || datum > toIsoDateOnly(aktion.enddatum)) {
            errors.push('Datum liegt außerhalb der Projektwoche (' + aktion.startdatum + ' – ' + aktion.enddatum + ').');
        }
    }

    const buchungAb = String(draft.buchungAb || '').trim();
    if (buchungAb && datum) {
        const openDay = toIsoDateOnly(buchungAb);
        if (openDay && openDay > datum) {
            warnings.push('Buchungsstart liegt nach dem Angebotsdatum.');
        }
    }

    const ziel = String(draft.zielklassen || '').trim();
    if (ziel && ziel.toLowerCase() !== 'alle' && classCodes.length) {
        const parts = ziel.split(/[,;]/).map((s) => s.trim()).filter(Boolean);
        const unknown = parts.filter((c) => classCodes.indexOf(c) === -1);
        if (unknown.length) warnings.push('Unbekannte Zielklassen: ' + unknown.join(', '));
    }

    const slot = String(draft.slot || '').trim();
    const lehrerCode = String(draft.lehrerCode || '').trim().toLowerCase();
    const selfId = String(draft.angebotId || draft.itemId || '');
    if (lehrerCode && datum && slot) {
        const clash = existing.filter((o) => {
            if (!o) return false;
            if (String(o.angebotId || o.itemId || '') === selfId) return false;
            const st = String(o.status || '');
            if (st === 'abgelehnt' || st === 'abgesagt') return false;
            return (
                String(o.lehrerCode || '').trim().toLowerCase() === lehrerCode &&
                toIsoDateOnly(o.datum) === datum &&
                String(o.slot || '') === slot
            );
        });
        if (clash.length) {
            warnings.push('Dieselbe Lehrkraft hat am gleichen Tag/Slot bereits „' + (clash[0].title || 'anderes Angebot') + '“.');
        }
    }

    if (title && datum) {
        const dup = existing.filter((o) => {
            if (!o) return false;
            if (String(o.angebotId || o.itemId || '') === selfId) return false;
            const st = String(o.status || '');
            if (st === 'abgelehnt' || st === 'abgesagt') return false;
            return String(o.title || '').trim().toLowerCase() === title.toLowerCase() && toIsoDateOnly(o.datum) === datum;
        });
        if (dup.length) warnings.push('Gleicher Titel am selben Tag existiert bereits.');
    }

    return { errors, warnings, canSubmit: errors.length === 0 };
}

/**
 * @param {object[]} angebote
 * @param {object|null} aktion
 * @param {Date} [now]
 */
export function computeDashboardKpis(angebote, aktion, now) {
    const list = Array.isArray(angebote) ? angebote : [];
    const beantragt = list.filter((a) => a.status === 'beantragt').length;
    const freigegeben = list.filter((a) => a.status === 'freigegeben').length;
    const abgelehnt = list.filter((a) => a.status === 'abgelehnt').length;
    const kapTotal = list
        .filter((a) => a.status === 'freigegeben' || a.status === 'beantragt')
        .reduce((s, a) => s + (Number(a.kapazitaet) || 0), 0);
    let buchungOffen = 0;
    let buchungGesperrt = 0;
    for (let i = 0; i < list.length; i++) {
        const a = list[i];
        if (a.status !== 'freigegeben') continue;
        if (isBookingOpen(effectiveBuchungAb(a, aktion), now)) buchungOffen++;
        else buchungGesperrt++;
    }
    return {
        total: list.length,
        beantragt,
        freigegeben,
        abgelehnt,
        kapTotal,
        buchungOffen,
        buchungGesperrt
    };
}

/**
 * Wochenraster: Tag × Slot → Angebote
 * @param {object[]} angebote
 */
export function buildWeekPlan(angebote) {
    /** @type {Record<string, Record<string, object[]>>} */
    const grid = {};
    for (let t = 0; t < TAGE.length; t++) {
        grid[TAGE[t]] = {};
        for (let s = 0; s < SLOTS.length; s++) grid[TAGE[t]][SLOTS[s]] = [];
    }
    const list = Array.isArray(angebote) ? angebote : [];
    for (let i = 0; i < list.length; i++) {
        const a = list[i];
        let tag = String(a.tag || '').trim();
        if (!tag || TAGE.indexOf(tag) === -1) tag = weekdayLabelDeFromIso(a.datum) || '';
        const slot = SLOTS.indexOf(String(a.slot || '')) >= 0 ? String(a.slot) : 'ganztags';
        if (!tag || !grid[tag]) continue;
        grid[tag][slot].push(a);
    }
    return grid;
}

/**
 * Zeilen für den Wochenplan: ganztags spannt vormittag–abend (kein eigene Spalte).
 * Bei Mischung am selben Tag: zuerst ganztags-Band, darunter Slot-Zellen.
 *
 * @param {ReturnType<typeof buildWeekPlan>} grid
 * @returns {{ tag: string, bands: Array<{ type: 'ganztags', items: object[] }|{ type: 'slots', bySlot: Record<string, object[]> }> }[]}
 */
export function buildWeekPlanDisplayRows(grid) {
    const g = grid && typeof grid === 'object' ? grid : {};
    return TAGE.map((tag) => {
        const day = g[tag] || {};
        const ganztags = Array.isArray(day.ganztags) ? day.ganztags : [];
        /** @type {Record<string, object[]>} */
        const bySlot = {};
        let halfCount = 0;
        for (let i = 0; i < PLAN_SLOTS.length; i++) {
            const slot = PLAN_SLOTS[i];
            const list = Array.isArray(day[slot]) ? day[slot] : [];
            bySlot[slot] = list;
            halfCount += list.length;
        }
        /** @type {Array<{ type: 'ganztags', items: object[] }|{ type: 'slots', bySlot: Record<string, object[]> }>} */
        const bands = [];
        if (ganztags.length) bands.push({ type: 'ganztags', items: ganztags });
        if (halfCount > 0 || !ganztags.length) bands.push({ type: 'slots', bySlot });
        return { tag, bands };
    });
}

/**
 * Kalender-Zellen für einen Monat.
 * @param {number} year
 * @param {number} month 0-11
 * @param {object[]} angebote
 */
export function buildMonthCells(year, month, angebote) {
    const first = new Date(year, month, 1);
    const startPad = (first.getDay() + 6) % 7; // Mo=0
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    /** @type {{ iso: string|null, day: number|null, items: object[] }[]} */
    const cells = [];
    for (let i = 0; i < startPad; i++) cells.push({ iso: null, day: null, items: [] });
    const byDate = {};
    const list = Array.isArray(angebote) ? angebote : [];
    for (let i = 0; i < list.length; i++) {
        const iso = toIsoDateOnly(list[i].datum);
        if (!iso) continue;
        if (!byDate[iso]) byDate[iso] = [];
        byDate[iso].push(list[i]);
    }
    for (let d = 1; d <= daysInMonth; d++) {
        const iso =
            year +
            '-' +
            String(month + 1).padStart(2, '0') +
            '-' +
            String(d).padStart(2, '0');
        cells.push({ iso, day: d, items: byDate[iso] || [] });
    }
    while (cells.length % 7) cells.push({ iso: null, day: null, items: [] });
    return cells;
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
 * Ohne diese Felder bleibt „Meine Angebote“ leer.
 *
 * @param {object} angebot
 * @param {{ accountEmail?: string, teachers?: Array<{ code?: string, email?: string, upn?: string }> }} [ctx]
 * @returns {{ beantragtVon: string, lehrerEmail: string, lehrerCode: string }}
 */
export function resolvePwOwnership(angebot, ctx) {
    const a = angebot && typeof angebot === 'object' ? angebot : {};
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
export function assertPwOwnership(ownership) {
    const email = String((ownership && (ownership.lehrerEmail || ownership.beantragtVon)) || '')
        .trim()
        .toLowerCase();
    if (!email || email.indexOf('@') < 0) {
        throw new Error(
            'Lehrer-E-Mail fehlt im Angebot (Anmeldung nicht erkannt). Bitte Seite neu laden oder ab-/anmelden und erneut speichern.'
        );
    }
    if (!String((ownership && ownership.beantragtVon) || '').trim()) {
        throw new Error('Feld „Beantragt von“ fehlt – bitte erneut speichern.');
    }
    if (!String((ownership && ownership.lehrerCode) || '').trim()) {
        throw new Error('Lehrer-Kürzel fehlt – bitte erneut speichern.');
    }
}

/**
 * @param {unknown} msg
 */
export function isPwAccessDeniedError(msg) {
    return /403|accessdenied|forbidden/i.test(String(msg || ''));
}

/**
 * Endnutzer-freundliche SPO-/Graph-Meldung (kein Roh-403).
 * @param {unknown} msg
 */
export function friendlyPwSpoError(msg) {
    const s = String(msg || '');
    if (isPwAccessDeniedError(s)) {
        return 'Kein SharePoint-Zugriff – lokale bzw. Demo-Daten bleiben erhalten. IT: Listenrechte prüfen.';
    }
    if (/nicht angemeldet|anmelden|login|interaction_required/i.test(s)) {
        return 'Bitte anmelden und erneut laden.';
    }
    if (/site|website|not found|404/i.test(s) && /site|liste|list/i.test(s)) {
        return 'Site oder Listen nicht gefunden – IT: SharePoint-URL und PW-Listen prüfen.';
    }
    return s.length > 220 ? s.slice(0, 217) + '…' : s;
}

/**
 * Aufmerksamkeit für Calm Action-Banner (ohne Domäne zu verbiegen).
 * @param {{ role?: string, items?: object[], accountEmail?: string, teacherMatch?: { code?: string }|null }} stateLike
 */
export function computePwAttention(stateLike) {
    const role = String((stateLike && stateLike.role) || 'lehrer');
    const items = Array.isArray(stateLike && stateLike.items) ? stateLike.items : [];
    const em = String((stateLike && stateLike.accountEmail) || '')
        .trim()
        .toLowerCase();
    const code = String((stateLike && stateLike.teacherMatch && stateLike.teacherMatch.code) || '')
        .trim()
        .toLowerCase();

    const beantragt = items.filter((a) => a && a.status === 'beantragt').length;
    const freigegeben = items.filter((a) => a && a.status === 'freigegeben');
    const pendingSync = freigegeben.filter((a) => {
        const st = String(a.syncStatus || '').toLowerCase();
        if (st === 'fehler') return true;
        return !String(a.bookingsServiceId || '').trim();
    }).length;
    const mine = items.filter((a) => {
        if (!a) return false;
        if (code && String(a.lehrerCode || '').toLowerCase() === code) return true;
        if (
            em &&
            (String(a.lehrerEmail || '').toLowerCase() === em ||
                String(a.beantragtVon || '').toLowerCase() === em)
        ) {
            return true;
        }
        return false;
    });
    const myOpen = mine.filter((a) => a.status === 'entwurf' || a.status === 'beantragt').length;

    return {
        role,
        beantragt,
        freigegeben: freigegeben.length,
        pendingSync,
        myCount: mine.length,
        myOpen,
        empty: items.length === 0
    };
}
