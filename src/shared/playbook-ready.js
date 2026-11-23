/**
 * Playbook Ready-Checks (Wave 0 Foundation).
 *
 * Hybrid-Modell: manuelle Checkbox bleibt; optionaler Ready-Check zeigt
 * status ok | warn | missing | unknown pro Schritt (Saison 1 ausbauen).
 *
 * evaluateReadyChecks() liefert Record<stepId, ReadyResult>.
 * wirePlaybookPage schreibt in [data-cp-ready-hint] und setzt
 * cp-step--ready-{status} am Schritt-Artikel.
 */

/** @typedef {'ok'|'warn'|'missing'|'unknown'} PlaybookReadyStatus */
/** @typedef {{ status: PlaybookReadyStatus, message?: string }} PlaybookReadyResult */

export const PLAYBOOK_READY_STATUS = Object.freeze({
    ok: 'ok',
    warn: 'warn',
    missing: 'missing',
    unknown: 'unknown'
});

/**
 * @param {unknown} raw
 * @returns {PlaybookReadyResult}
 */
export function normalizeReadyResult(raw) {
    if (!raw || typeof raw !== 'object') {
        return { status: 'unknown', message: '' };
    }
    const o = /** @type {Record<string, unknown>} */ (raw);
    let status = String(o.status || '').trim().toLowerCase();
    if (status !== 'ok' && status !== 'warn' && status !== 'missing' && status !== 'unknown') {
        if (o.ok === true) status = 'ok';
        else if (o.ok === false) status = 'missing';
        else status = 'unknown';
    }
    return {
        status: /** @type {PlaybookReadyStatus} */ (status),
        message: String(o.message || '').trim()
    };
}

/**
 * Legacy-Gates ({ ok, message }) → Ready-Map.
 * @param {Record<string, { ok?: boolean, message?: string }>|null|undefined} gates
 * @returns {Record<string, PlaybookReadyResult>}
 */
export function gatesToReadyMap(gates) {
    /** @type {Record<string, PlaybookReadyResult>} */
    const out = {};
    if (!gates || typeof gates !== 'object') return out;
    Object.keys(gates).forEach(function (id) {
        const g = gates[id];
        if (!g || typeof g !== 'object') {
            out[id] = { status: 'unknown', message: '' };
            return;
        }
        out[id] = {
            status: g.ok ? 'ok' : 'missing',
            message: String(g.message || '').trim()
        };
    });
    return out;
}

/**
 * @param {PlaybookReadyResult|null|undefined} result
 * @returns {string}
 */
export function formatReadyHint(result) {
    const r = normalizeReadyResult(result);
    const msg = r.message || '';
    if (r.status === 'ok') {
        return msg ? 'Bereit: ' + msg : 'Bereit';
    }
    if (r.status === 'warn') {
        return msg ? 'Hinweis: ' + msg : 'Hinweis';
    }
    if (r.status === 'missing') {
        return msg ? 'Fehlt: ' + msg : 'Noch nicht bereit';
    }
    return msg || '';
}

/**
 * @param {Record<string, PlaybookReadyResult>|null|undefined} map
 * @param {string} stepId
 * @returns {PlaybookReadyResult}
 */
export function readyForStep(map, stepId) {
    if (!map || typeof map !== 'object') return { status: 'unknown', message: '' };
    return normalizeReadyResult(map[stepId]);
}

if (typeof window !== 'undefined') {
    window.ms365PlaybookReady = {
        PLAYBOOK_READY_STATUS,
        normalizeReadyResult,
        gatesToReadyMap,
        formatReadyHint,
        readyForStep
    };
}
