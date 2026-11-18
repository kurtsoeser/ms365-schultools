/**
 * Domänenlogik Technik-Meldungen (DOM-frei).
 */

import {
    KATEGORIE,
    PRIORITAET,
    STATUS,
    STATUS_LABEL,
    normalizeMeldung
} from './technik-meldungen-model.js';

/**
 * @param {object} draft
 */
export function validateMeldung(draft) {
    const errors = [];
    const m = normalizeMeldung(draft);
    if (!m.title) errors.push('Titel fehlt.');
    if (!m.raumCode && !m.raumName) errors.push('Raum (Code oder Name) fehlt.');
    if (!Object.values(KATEGORIE).includes(m.kategorie)) {
        errors.push('Kategorie ungültig.');
    }
    if (!Object.values(PRIORITAET).includes(m.prioritaet)) {
        errors.push('Priorität ungültig.');
    }
    if (!Object.values(STATUS).includes(m.status)) {
        errors.push('Status ungültig.');
    }
    return { ok: errors.length === 0, errors, item: m };
}

/**
 * @param {object[]} meldungen
 * @param {{ status?: string, kategorie?: string, q?: string, nurOffen?: boolean, meEmail?: string, admin?: boolean, showAll?: boolean }} filters
 */
export function filterMeldungen(meldungen, filters = {}) {
    const status = String(filters.status || '')
        .trim()
        .toLowerCase();
    const kategorie = String(filters.kategorie || '')
        .trim()
        .toLowerCase();
    const q = String(filters.q || '')
        .trim()
        .toLowerCase();
    const me = String(filters.meEmail || '')
        .trim()
        .toLowerCase();
    const canSeeAll = !!(filters.admin && filters.showAll);

    return (meldungen || []).filter((m) => {
        if (!m) return false;
        if (filters.nurOffen) {
            if (m.status !== STATUS.OFFEN && m.status !== STATUS.IN_ARBEIT) return false;
        }
        if (status && m.status !== status) return false;
        if (kategorie && m.kategorie !== kategorie) return false;
        if (!canSeeAll) {
            if (!me || String(m.melderEmail || '').toLowerCase() !== me) return false;
        }
        if (q) {
            const hay = (
                m.title +
                ' ' +
                m.beschreibung +
                ' ' +
                m.raumCode +
                ' ' +
                m.raumName +
                ' ' +
                m.melderName +
                ' ' +
                m.kommentar
            ).toLowerCase();
            if (!hay.includes(q)) return false;
        }
        return true;
    });
}

/**
 * @param {object[]} meldungen
 */
export function meldungenStats(meldungen) {
    const list = meldungen || [];
    const offen = list.filter((m) => m.status === STATUS.OFFEN || m.status === STATUS.IN_ARBEIT);
    return {
        total: list.length,
        offen: offen.length,
        inArbeit: list.filter((m) => m.status === STATUS.IN_ARBEIT).length,
        erledigt: list.filter((m) => m.status === STATUS.ERLEDIGT).length,
        dringend: offen.filter((m) => m.prioritaet === PRIORITAET.DRINGEND).length
    };
}

/**
 * @param {object[]} meldungen
 */
export function sortMeldungen(meldungen) {
    const prioRank = { dringend: 0, hoch: 1, mittel: 2, niedrig: 3 };
    return [...(meldungen || [])].sort((a, b) => {
        const pa = prioRank[a.prioritaet] ?? 9;
        const pb = prioRank[b.prioritaet] ?? 9;
        if (pa !== pb) return pa - pb;
        const ga = a.gemeldetAm || '';
        const gb = b.gemeldetAm || '';
        if (ga !== gb) return ga > gb ? -1 : 1;
        return String(a.title || '').localeCompare(String(b.title || ''), 'de');
    });
}

/**
 * Status-Pipeline-Aktionen für die Admin-Inbox.
 * @param {string} status
 * @returns {{ to: string, label: string, kind?: string }[]}
 */
export function statusPipelineActions(status) {
    const st = String(status || '')
        .trim()
        .toLowerCase();
    if (st === STATUS.ERLEDIGT || st === STATUS.ABGELEHNT) {
        return [{ to: STATUS.OFFEN, label: 'Wieder öffnen' }];
    }
    if (st === STATUS.IN_ARBEIT) {
        return [
            { to: STATUS.ERLEDIGT, label: 'Erledigt', kind: 'success' },
            { to: STATUS.ABGELEHNT, label: 'Ablehnen' },
            { to: STATUS.OFFEN, label: 'Zurück auf offen' }
        ];
    }
    return [
        { to: STATUS.IN_ARBEIT, label: 'In Arbeit' },
        { to: STATUS.ERLEDIGT, label: 'Erledigt', kind: 'success' },
        { to: STATUS.ABGELEHNT, label: 'Ablehnen' }
    ];
}

/**
 * @param {object|null|undefined} meldung
 * @param {{ admin?: boolean }} [ctx]
 */
export function canChangeStatus(meldung, ctx = {}) {
    return !!(meldung && ctx.admin);
}

/**
 * @param {object|null|undefined} meldung
 * @param {{ admin?: boolean, meEmail?: string }} [ctx]
 */
export function canCommentOnMeldung(meldung, ctx = {}) {
    if (!meldung) return false;
    if (ctx.admin) return true;
    const me = String(ctx.meEmail || '')
        .trim()
        .toLowerCase();
    return !!me && String(meldung.melderEmail || '').toLowerCase() === me;
}

/**
 * @param {string} status
 */
export function statusLabel(status) {
    return STATUS_LABEL[status] || status || '–';
}
