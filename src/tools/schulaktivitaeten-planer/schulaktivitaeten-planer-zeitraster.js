/**
 * Standard-Zeitraster (Unterrichtseinheiten → Uhrzeiten) für Schulaktivitäten.
 * Schul-spezifische Raster können später über Stammdaten/Regelwerk ergänzt werden.
 */

/**
 * @typedef {{ nr: number, label: string, start: string, end: string }} Unterrichtseinheit
 */

/** Typisches österreichisches Sekundarstufen-Raster (50 Min + Pausen). */
export const DEFAULT_ZEITRASTER = Object.freeze([
    { nr: 1, label: '1. UE', start: '07:45', end: '08:35' },
    { nr: 2, label: '2. UE', start: '08:40', end: '09:30' },
    { nr: 3, label: '3. UE', start: '09:45', end: '10:35' },
    { nr: 4, label: '4. UE', start: '10:40', end: '11:30' },
    { nr: 5, label: '5. UE', start: '11:45', end: '12:35' },
    { nr: 6, label: '6. UE', start: '12:40', end: '13:30' },
    { nr: 7, label: '7. UE', start: '13:35', end: '14:25' },
    { nr: 8, label: '8. UE', start: '14:30', end: '15:20' },
    { nr: 9, label: '9. UE', start: '15:25', end: '16:15' },
    { nr: 10, label: '10. UE', start: '16:20', end: '17:10' }
]);

/**
 * @param {unknown} value
 * @returns {number|null}
 */
export function parseUeNumber(value) {
    if (value == null || value === '') return null;
    const n = Number(String(value).trim().replace(/[^\d]/g, ''));
    if (!Number.isFinite(n) || n < 1 || n > 20) return null;
    return Math.round(n);
}

/**
 * @param {Unterrichtseinheit[]|null|undefined} raster
 */
export function resolveZeitraster(raster) {
    if (Array.isArray(raster) && raster.length) {
        return raster
            .map((r) => ({
                nr: parseUeNumber(r && r.nr),
                label: String((r && (r.label || r.name)) || '').trim(),
                start: String((r && r.start) || '').trim(),
                end: String((r && r.end) || '').trim()
            }))
            .filter((r) => r.nr != null && /^\d{1,2}:\d{2}$/.test(r.start) && /^\d{1,2}:\d{2}$/.test(r.end));
    }
    return DEFAULT_ZEITRASTER.slice();
}

/**
 * @param {number|string|null|undefined} ue
 * @param {Unterrichtseinheit[]|null|undefined} raster
 */
export function ueSlot(ue, raster) {
    const nr = parseUeNumber(ue);
    if (nr == null) return null;
    const list = resolveZeitraster(raster);
    return list.find((r) => r.nr === nr) || null;
}

/**
 * Start-/Endzeit aus UE-Bereich ableiten (Beginn der Start-UE, Ende der End-UE).
 * @param {object} opts
 * @param {number|string|null|undefined} opts.startUe
 * @param {number|string|null|undefined} opts.endUe
 * @param {Unterrichtseinheit[]|null|undefined} [opts.raster]
 * @returns {{ startZeit: string, endZeit: string, startUe: number|null, endUe: number|null, ok: boolean, error?: string }}
 */
export function timesFromUnterrichtseinheiten(opts) {
    const o = opts || {};
    let startUe = parseUeNumber(o.startUe);
    let endUe = parseUeNumber(o.endUe);
    if (startUe == null && endUe == null) {
        return { startZeit: '', endZeit: '', startUe: null, endUe: null, ok: true };
    }
    if (startUe == null) startUe = endUe;
    if (endUe == null) endUe = startUe;
    if (endUe < startUe) {
        return {
            startZeit: '',
            endZeit: '',
            startUe,
            endUe,
            ok: false,
            error: 'Die End-UE darf nicht vor der Start-UE liegen.'
        };
    }
    const raster = resolveZeitraster(o.raster);
    const startSlot = ueSlot(startUe, raster);
    const endSlot = ueSlot(endUe, raster);
    if (!startSlot || !endSlot) {
        return {
            startZeit: '',
            endZeit: '',
            startUe,
            endUe,
            ok: false,
            error: 'Unterrichtseinheit liegt außerhalb des Zeitrasters.'
        };
    }
    return {
        startZeit: startSlot.start,
        endZeit: endSlot.end,
        startUe,
        endUe,
        ok: true
    };
}

/**
 * Formular-/Datensatz: UE setzen und Uhrzeiten ableiten (manuelle Uhrzeit bleibt, wenn keine UE).
 * @param {object} draft
 * @param {Unterrichtseinheit[]|null|undefined} [raster]
 */
export function applyUeToDraft(draft, raster) {
    const d = { ...(draft || {}) };
    const derived = timesFromUnterrichtseinheiten({
        startUe: d.startUe,
        endUe: d.endUe,
        raster
    });
    d.startUe = derived.startUe == null ? '' : derived.startUe;
    d.endUe = derived.endUe == null ? '' : derived.endUe;
    if (derived.startUe != null) {
        d.startZeit = derived.startZeit;
        d.endZeit = derived.endZeit;
    }
    return { draft: d, derived };
}

/**
 * Anzeige z. B. „1.–4. UE (07:45–11:30)“.
 * @param {object} item
 * @param {Unterrichtseinheit[]|null|undefined} [raster]
 */
export function formatUeZeitraum(item, raster) {
    const startUe = parseUeNumber(item && item.startUe);
    const endUe = parseUeNumber(item && item.endUe);
    const startZeit = String((item && item.startZeit) || '').trim();
    const endZeit = String((item && item.endZeit) || '').trim();
    if (startUe != null) {
        const ueLabel =
            endUe != null && endUe !== startUe ? `${startUe}.–${endUe}. UE` : `${startUe}. UE`;
        const clock =
            startZeit && endZeit ? ` (${startZeit}–${endZeit})` : startZeit ? ` (${startZeit})` : '';
        return ueLabel + clock;
    }
    if (startZeit || endZeit) return [startZeit, endZeit].filter(Boolean).join('–');
    const derived = timesFromUnterrichtseinheiten({ startUe, endUe, raster });
    if (derived.ok && derived.startZeit) {
        return `${derived.startUe}.–${derived.endUe}. UE (${derived.startZeit}–${derived.endZeit})`;
    }
    return '';
}

export function ueChoices(raster) {
    return resolveZeitraster(raster).map((r) => ({
        code: String(r.nr),
        name: `${r.label} (${r.start}–${r.end})`
    }));
}
